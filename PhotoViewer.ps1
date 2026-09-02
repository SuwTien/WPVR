#Requires -Version 5.1
<#
    Photo Viewer - Etape 2 : switch entre repertoire "existant" (lecture seule) et "fraiches" (lecture + ecriture).

    IMPORTANT : ce script doit etre lance via "powershell.exe" (Windows PowerShell), PAS "pwsh.exe".
    WPF exige un thread STA ; powershell.exe est STA par defaut, pwsh.exe (PowerShell 7) est MTA par defaut
    et plantera au chargement de la fenetre sans l'option -sta.
#>

# ------------------------------------------------------------------
# Configuration : deux repertoires (existant = lecture seule, fraiches = lecture + ecriture)
# Les chemins reels sont choisis par l'utilisateur via le bouton "Dossier..." et persistes
# dans un fichier JSON a cote du script (l'appli ne cree jamais de dossier elle-meme).
# ------------------------------------------------------------------
$script:configPath = Join-Path $PSScriptRoot 'PhotoViewer.config.json'
$ExistingPhotoDirectory = "$env:USERPROFILE\Pictures"
$FreshPhotoDirectory = "$env:USERPROFILE\Pictures\Fraiches"
$ThumbnailCellWidth = 158   # largeur du bouton (150) + marges (4+4)
$ThumbnailDecodePixelWidth = 220
$MaxConcurrentThumbnailLoads = 4
$LongPressThresholdMs = 500

if (Test-Path -LiteralPath $script:configPath -PathType Leaf) {
    try {
        $saved = Get-Content -LiteralPath $script:configPath -Raw | ConvertFrom-Json
        if ($saved.ExistingPhotoDirectory) { $ExistingPhotoDirectory = $saved.ExistingPhotoDirectory }
        if ($saved.FreshPhotoDirectory) { $FreshPhotoDirectory = $saved.FreshPhotoDirectory }
    } catch {
        # Config illisible/corrompue : on retombe sur les valeurs par defaut ci-dessus.
    }
}

# Repertoire actif au demarrage : "fraiches" (workflow principal de renommage/tri).
# Les chemins peuvent ne pas (encore) exister sur le disque : la grille reste vide
# jusqu'a ce que l'utilisateur choisisse un dossier valide via le bouton "Dossier...".
$script:directoryConfigs = @(
    [PSCustomObject]@{ Path = $FreshPhotoDirectory; IsReadOnly = $false }
    [PSCustomObject]@{ Path = $ExistingPhotoDirectory; IsReadOnly = $true }
)
$script:currentDirectoryIndex = 0

function Save-DirectoryConfig {
    [PSCustomObject]@{
        ExistingPhotoDirectory = $script:directoryConfigs[1].Path
        FreshPhotoDirectory = $script:directoryConfigs[0].Path
    } | ConvertTo-Json | Set-Content -LiteralPath $script:configPath -Encoding UTF8
}

Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, System.Xaml, System.Windows.Forms

# ------------------------------------------------------------------
# Types C# compiles a la volee (Add-Type) : modele de donnees, commande,
# et detection d'appui long (base reutilisable pour le mode plein ecran, etape 5).
# ------------------------------------------------------------------
Add-Type -ReferencedAssemblies PresentationCore, PresentationFramework, WindowsBase, System.Xaml -TypeDefinition @"
using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace PhotoViewer.Models
{
    public class PhotoItem : INotifyPropertyChanged
    {
        private string _fullPath;
        public string FullPath
        {
            get { return _fullPath; }
            set { _fullPath = value; OnPropertyChanged("FullPath"); }
        }

        private string _fileName;
        public string FileName
        {
            get { return _fileName; }
            set { _fileName = value; OnPropertyChanged("FileName"); }
        }

        private ImageSource _thumbnail;
        public ImageSource Thumbnail
        {
            get { return _thumbnail; }
            set { _thumbnail = value; OnPropertyChanged("Thumbnail"); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string name)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(name));
        }
    }
}

namespace PhotoViewer.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private string _directoryLabel;
        public string DirectoryLabel
        {
            get { return _directoryLabel; }
            set { _directoryLabel = value; OnPropertyChanged("DirectoryLabel"); }
        }

        private bool _isReadOnly;
        public bool IsReadOnly
        {
            get { return _isReadOnly; }
            set { _isReadOnly = value; OnPropertyChanged("IsReadOnly"); }
        }

        public ICommand TapCommand { get; set; }
        public ICommand ToggleDirectoryCommand { get; set; }
        public ICommand ChooseDirectoryCommand { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string name)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(name));
        }
    }
}

namespace PhotoViewer.Commands
{
    public class RelayCommand : ICommand
    {
        private readonly Action<object> _execute;
        private readonly Func<object, bool> _canExecute;

        public RelayCommand(Action<object> execute, Func<object, bool> canExecute = null)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter) { return _canExecute == null || _canExecute(parameter); }
        public void Execute(object parameter) { _execute(parameter); }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
    }
}

namespace PhotoViewer.Gestures
{
    using PhotoViewer.Models;

    // Validation technique de l'approche Add-Type pour les gestes : detection d'un appui long
    // par element (via hit-test sur le DataContext), reutilisable pour le mode plein ecran (etape 5).
    public static class LongPressGesture
    {
        public static void Attach(FrameworkElement root, Action<PhotoItem> onLongPress, int thresholdMs)
        {
            DispatcherTimer timer = null;
            PhotoItem pressedItem = null;
            Point downPos = default(Point);
            bool longPressFired = false;

            Action<DependencyObject, Point> startPress = (source, pos) =>
            {
                var element = FindPhotoItemElement(source);
                if (element == null) return;

                pressedItem = element.DataContext as PhotoItem;
                downPos = pos;
                longPressFired = false;

                timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(thresholdMs) };
                timer.Tick += (s, e) =>
                {
                    timer.Stop();
                    var item = pressedItem;
                    pressedItem = null;
                    longPressFired = true;
                    if (item != null) onLongPress(item);
                };
                timer.Start();
            };

            Action cancelPress = () =>
            {
                if (timer != null) { timer.Stop(); timer = null; }
                pressedItem = null;
            };

            root.PreviewMouseLeftButtonDown += (s, e) => startPress(e.OriginalSource as DependencyObject, e.GetPosition(root));
            root.PreviewMouseMove += (s, e) =>
            {
                if (pressedItem != null && e.LeftButton == MouseButtonState.Pressed)
                {
                    var p = e.GetPosition(root);
                    if ((p - downPos).Length > 15) cancelPress();
                }
            };
            root.PreviewMouseLeftButtonUp += (s, e) =>
            {
                // L'appui long a deja declenche son action : on avale ce relachement pour eviter
                // que le Click normal du bouton (tap -> renommage) se declenche juste apres.
                if (longPressFired) { e.Handled = true; longPressFired = false; }
                cancelPress();
            };
            root.PreviewTouchDown += (s, e) => startPress(e.OriginalSource as DependencyObject, e.GetTouchPoint(root).Position);
            root.PreviewTouchUp += (s, e) =>
            {
                if (longPressFired) { e.Handled = true; longPressFired = false; }
                cancelPress();
            };
        }

        private static FrameworkElement FindPhotoItemElement(DependencyObject d)
        {
            while (d != null)
            {
                var fe = d as FrameworkElement;
                if (fe != null && fe.DataContext is PhotoItem) return fe;
                d = VisualTreeHelper.GetParent(d);
            }
            return null;
        }
    }

    // Zoom (pincement tactile ou molette) + pan + swipe gauche/droite (mode plein ecran, etape 5).
    // Le swipe n'est detecte que si l'image n'est pas zoomee (comportement standard des visionneuses photo).
    public static class ZoomPanGesture
    {
        private const double MinScale = 1.0;
        private const double MaxScale = 5.0;
        private const double SwipeThreshold = 60.0;

        public static void Attach(FrameworkElement target, ScaleTransform scale, TranslateTransform translate, Action onSwipeLeft, Action onSwipeRight)
        {
            target.IsManipulationEnabled = true;
            double accumulatedX = 0;

            target.ManipulationDelta += (s, e) =>
            {
                double newScale = Math.Max(MinScale, Math.Min(MaxScale, scale.ScaleX * e.DeltaManipulation.Scale.X));
                scale.ScaleX = newScale;
                scale.ScaleY = newScale;

                if (newScale > MinScale + 0.01)
                {
                    translate.X += e.DeltaManipulation.Translation.X;
                    translate.Y += e.DeltaManipulation.Translation.Y;
                }
                else
                {
                    accumulatedX += e.DeltaManipulation.Translation.X;
                }
                e.Handled = true;
            };

            target.ManipulationCompleted += (s, e) =>
            {
                if (scale.ScaleX <= MinScale + 0.01)
                {
                    if (accumulatedX <= -SwipeThreshold && onSwipeLeft != null) onSwipeLeft();
                    else if (accumulatedX >= SwipeThreshold && onSwipeRight != null) onSwipeRight();
                    translate.X = 0;
                    translate.Y = 0;
                }
                accumulatedX = 0;
            };

            target.PreviewMouseWheel += (s, e) =>
            {
                double factor = e.Delta > 0 ? 1.1 : (1 / 1.1);
                double newScale = Math.Max(MinScale, Math.Min(MaxScale, scale.ScaleX * factor));
                scale.ScaleX = newScale;
                scale.ScaleY = newScale;
                if (newScale <= MinScale + 0.01) { translate.X = 0; translate.Y = 0; }
                e.Handled = true;
            };

            // Deplacement (pan) a la souris une fois zoome : la Manipulation ne couvre que le tactile.
            bool dragging = false;
            Point dragStart = default(Point);
            double startTranslateX = 0, startTranslateY = 0;

            target.PreviewMouseLeftButtonDown += (s, e) =>
            {
                if (scale.ScaleX <= MinScale + 0.01) return;
                dragging = true;
                dragStart = e.GetPosition(target);
                startTranslateX = translate.X;
                startTranslateY = translate.Y;
                target.CaptureMouse();
                e.Handled = true;
            };
            target.PreviewMouseMove += (s, e) =>
            {
                if (!dragging) return;
                var pos = e.GetPosition(target);
                translate.X = startTranslateX + (pos.X - dragStart.X);
                translate.Y = startTranslateY + (pos.Y - dragStart.Y);
            };
            target.PreviewMouseLeftButtonUp += (s, e) =>
            {
                if (!dragging) return;
                dragging = false;
                target.ReleaseMouseCapture();
            };
        }

        public static void Reset(ScaleTransform scale, TranslateTransform translate)
        {
            scale.ScaleX = 1;
            scale.ScaleY = 1;
            translate.X = 0;
            translate.Y = 0;
        }
    }
}

namespace PhotoViewer.Interop
{
    // Bascule fiable du clavier tactile Windows (TabTip.exe ne s'affiche pas toujours
    // via un simple Start-Process s'il tourne deja en tache de fond).
    [ComImport, Guid("4CE576FA-83DC-4F88-951C-9D0782B4E376")]
    internal class UIHostNoLaunch { }

    [ComImport, Guid("37c994e7-432b-4834-a2f7-dce1f13b834b"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ITipInvocation
    {
        void Toggle(IntPtr hwnd);
    }

    public static class TouchKeyboard
    {
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("dwmapi.dll")]
        private static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out int pvAttribute, int cbAttribute);

        private const int DWMWA_CLOAKED = 14;

        // IsWindowVisible ne suffit pas : le clavier tactile moderne (UWP) reste "visible" au sens Win32
        // meme masque, Windows le cache via le "cloaking" DWM plutot qu'un simple ShowWindow. Sans cette
        // verification, on le croit ferme alors qu'il est ouvert, et Toggle() le referme par erreur.
        private static bool IsActuallyVisible(IntPtr hWnd)
        {
            if (!IsWindowVisible(hWnd)) return false;
            int cloaked;
            int hr = DwmGetWindowAttribute(hWnd, DWMWA_CLOAKED, out cloaked, sizeof(int));
            if (hr == 0 && cloaked != 0) return false;
            return true;
        }

        // ITipInvocation.Toggle() bascule montre/cache : on ne l'appelle donc que si le clavier
        // n'est PAS deja visible, sinon un appel alors qu'il est ouvert le referme (comportement surprenant).
        public static void Show(IntPtr ownerHandle)
        {
            try
            {
                IntPtr tipWindow = FindWindow("IPTip_Main_Window", null);
                if (tipWindow != IntPtr.Zero && IsActuallyVisible(tipWindow)) return;
                var invocation = (ITipInvocation)new UIHostNoLaunch();
                invocation.Toggle(ownerHandle);
            }
            catch
            {
                // Clavier tactile indisponible (ex. Bureau classique sans TabTip) : ignore.
            }
        }
    }
}

namespace PhotoViewer.Services
{
    using PhotoViewer.Models;

    // Compteur de generation thread-safe : invalide les chargements de miniatures encore en vol
    // depuis un repertoire quitte (pas de dependance a une variable PowerShell, lue depuis un thread du pool).
    public class GenerationGate
    {
        private int _current;
        public int Current { get { return Volatile.Read(ref _current); } }
        public int Next() { return Interlocked.Increment(ref _current); }
    }

    // Chargement asynchrone des miniatures entierement en C# : un ScriptBlock PowerShell ne peut PAS
    // s'executer sur un thread brut du pool .NET (Task.Run) faute de Runspace attache a ce thread
    // ("There is no Runspace available to run scripts in this thread") - d'ou l'echec silencieux
    // total de la version precedente. Ce code est compile, donc utilisable depuis n'importe quel thread.
    public static class ThumbnailLoader
    {
        private static SemaphoreSlim _semaphore = new SemaphoreSlim(4, 4);
        private static readonly object _logLock = new object();

        public static void Configure(int maxConcurrent)
        {
            _semaphore = new SemaphoreSlim(maxConcurrent, maxConcurrent);
        }

        public static void LoadAsync(PhotoItem photo, int decodePixelWidth, Dispatcher dispatcher, GenerationGate gate, int generation, string errorLogPath)
        {
            Task.Run(() =>
            {
                if (gate.Current != generation) return;
                _semaphore.Wait();
                try
                {
                    if (gate.Current != generation) return;
                    BitmapImage bmp;
                    using (var stream = File.OpenRead(photo.FullPath))
                    {
                        bmp = new BitmapImage();
                        bmp.BeginInit();
                        bmp.CacheOption = BitmapCacheOption.OnLoad;
                        bmp.StreamSource = stream;
                        bmp.DecodePixelWidth = decodePixelWidth;
                        bmp.EndInit();
                    }
                    bmp.Freeze();
                    dispatcher.Invoke((Action)(() =>
                    {
                        if (gate.Current == generation) photo.Thumbnail = bmp;
                    }));
                }
                catch (Exception ex)
                {
                    try
                    {
                        lock (_logLock)
                        {
                            File.AppendAllText(errorLogPath, string.Format("[{0:yyyy-MM-dd HH:mm:ss}] Miniature KO '{1}': {2}{3}", DateTime.Now, photo.FullPath, ex.Message, Environment.NewLine));
                        }
                    }
                    catch
                    {
                        // Log indisponible (droits, disque plein...) : on abandonne silencieusement ce log.
                    }
                }
                finally
                {
                    _semaphore.Release();
                }
            });
        }

        // Chargement d'une image plein ecran (pas de limite de concurrence : une seule a la fois en pratique).
        public static void LoadFullImageAsync(string fullPath, int decodePixelWidth, Dispatcher dispatcher, GenerationGate gate, int generation, Action<BitmapImage> onLoaded)
        {
            Task.Run(() =>
            {
                if (gate.Current != generation) return;
                try
                {
                    BitmapImage bmp;
                    using (var stream = File.OpenRead(fullPath))
                    {
                        bmp = new BitmapImage();
                        bmp.BeginInit();
                        bmp.CacheOption = BitmapCacheOption.OnLoad;
                        bmp.StreamSource = stream;
                        if (decodePixelWidth > 0) bmp.DecodePixelWidth = decodePixelWidth;
                        bmp.EndInit();
                    }
                    bmp.Freeze();
                    dispatcher.Invoke((Action)(() =>
                    {
                        if (gate.Current == generation) onLoaded(bmp);
                    }));
                }
                catch
                {
                    // Image plein ecran illisible : la vue reste vide, l'utilisateur peut naviguer ailleurs.
                }
            });
        }
    }

    // Surveillance d'un repertoire (detection auto des nouvelles photos, etape 6). Ecrit entierement en C#
    // pour la meme raison que ThumbnailLoader : Register-ObjectEvent ne se declenche JAMAIS tant que le
    // thread principal est bloque dans Application.Run() (celui-ci pompe les messages WPF, pas la queue
    // d'evenements PowerShell). FileSystemWatcher leve ses evenements sur un thread du pool ; on rejoint
    // le thread UI via Dispatcher.Invoke (qui LUI est activement pompe par Application.Run) avant d'appeler
    // le callback PowerShell fourni - c'est la meme technique que le chargement des miniatures.
    public static class DirectoryWatcher
    {
        public static FileSystemWatcher Start(string path, Dispatcher dispatcher, Action onChanged, int debounceMs)
        {
            var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(debounceMs) };
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                onChanged();
            };

            FileSystemEventHandler onEvent = (s, e) =>
            {
                dispatcher.Invoke((Action)(() =>
                {
                    timer.Stop();
                    timer.Start();
                }));
            };

            var watcher = new FileSystemWatcher(path);
            watcher.IncludeSubdirectories = false;
            watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite;
            watcher.Created += onEvent;
            watcher.Renamed += (s, e) => onEvent(s, e);
            watcher.EnableRaisingEvents = true;
            return watcher;
        }
    }
}
"@

# ------------------------------------------------------------------
# Chargement de la fenetre depuis le fichier XAML
# ------------------------------------------------------------------
$xamlPath = Join-Path $PSScriptRoot 'MainWindow.xaml'
$reader = [System.Xml.XmlReader]::Create($xamlPath)
$window = [Windows.Markup.XamlReader]::Load($reader)
$reader.Close()

$directoryLabel = $window.FindName('DirectoryLabel')
$scrollViewer = $window.FindName('GridScrollViewer')
$rowsItemsControl = $window.FindName('RowsItemsControl')
$refreshButton = $window.FindName('RefreshButton')

$fullScreenOverlay = $window.FindName('FullScreenOverlay')
$fullScreenImageContainer = $window.FindName('FullScreenImageContainer')
$fullScreenImage = $window.FindName('FullScreenImage')
$fullScreenScale = $window.FindName('FullScreenScale')
$fullScreenTranslate = $window.FindName('FullScreenTranslate')
$fullScreenPrevButton = $window.FindName('FullScreenPrevButton')
$fullScreenNextButton = $window.FindName('FullScreenNextButton')
$fullScreenCloseButton = $window.FindName('FullScreenCloseButton')
$fullScreenFileNameButton = $window.FindName('FullScreenFileNameButton')

$uiDispatcher = $window.Dispatcher

# ------------------------------------------------------------------
# Clavier tactile Windows : InputScope numerique par defaut (l'utilisateur bascule vers
# les lettres via le bouton propre au clavier tactile lui-meme si besoin - pas de bascule custom).
# ------------------------------------------------------------------
function Set-NumericInputScope {
    param($TextBox)
    $scope = New-Object System.Windows.Input.InputScope
    $scopeName = New-Object System.Windows.Input.InputScopeName
    $scopeName.NameValue = [System.Windows.Input.InputScopeNameValue]::Number
    $scope.Names.Add($scopeName)
    $TextBox.InputScope = $scope
}

# ------------------------------------------------------------------
# Popup de renommage (repertoire "fraiches" uniquement)
# ------------------------------------------------------------------
function Show-RenamePopup {
    param([PhotoViewer.Models.PhotoItem]$Photo)

    $popupXamlPath = Join-Path $PSScriptRoot 'RenamePopup.xaml'
    $popupReader = [System.Xml.XmlReader]::Create($popupXamlPath)
    $popup = [Windows.Markup.XamlReader]::Load($popupReader)
    $popupReader.Close()
    $popup.Owner = $window

    $nameTextBox = $popup.FindName('NameTextBox')
    $extensionLabel = $popup.FindName('ExtensionLabel')
    $trashButton = $popup.FindName('TrashButton')
    $validateButton = $popup.FindName('ValidateButton')
    $cancelButton = $popup.FindName('CancelButton')

    $extension = [System.IO.Path]::GetExtension($Photo.FileName)
    $nameTextBox.Text = [System.IO.Path]::GetFileNameWithoutExtension($Photo.FileName)
    $extensionLabel.Text = $extension
    Set-NumericInputScope -TextBox $nameTextBox

    # HWND de la popup (force la creation du handle avant meme l'affichage) pour piloter le clavier tactile.
    $popupHandleHelper = New-Object System.Windows.Interop.WindowInteropHelper($popup)
    $popupHandleHelper.EnsureHandle() | Out-Null
    $popupHandle = $popupHandleHelper.Handle

    # Selection du nom + ouverture du clavier tactile une fois la popup reellement affichee.
    $popup.Add_Loaded({
        param($s, $e)
        $nameTextBox.Focus()
        $nameTextBox.SelectAll()
        [PhotoViewer.Interop.TouchKeyboard]::Show($popupHandle)
    })

    # Bouton corbeille : confirmation obligatoire puis suppression reelle du fichier.
    $trashButton.Add_Click({
        param($s, $e)
        $confirm = [System.Windows.MessageBox]::Show(
            "Supprimer '$($Photo.FileName)' ?`nCette action est definitive.",
            "Photo Viewer", 'YesNo', 'Warning')
        if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }

        try {
            Remove-Item -LiteralPath $Photo.FullPath -ErrorAction Stop
        } catch {
            [System.Windows.MessageBox]::Show("Erreur lors de la suppression : $_", "Photo Viewer", 'OK', 'Error') | Out-Null
            return
        }

        $allPhotos.Remove($Photo) | Out-Null
        Update-PhotoRows -Columns $script:currentColumnCount -Force
        $popup.DialogResult = $false
    })

    $cancelButton.Add_Click({
        param($s, $e)
        $popup.DialogResult = $false
    })

    $validateButton.Add_Click({
        param($s, $e)
        $newBaseName = $nameTextBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($newBaseName)) {
            [System.Windows.MessageBox]::Show("Le nom ne peut pas etre vide.", "Photo Viewer", 'OK', 'Warning') | Out-Null
            return
        }
        if ($newBaseName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0) {
            [System.Windows.MessageBox]::Show("Le nom contient des caracteres non autorises.", "Photo Viewer", 'OK', 'Warning') | Out-Null
            return
        }

        $newFileName = "$newBaseName$extension"
        if ($newFileName -ne $Photo.FileName) {
            $newFullPath = Join-Path (Split-Path -Path $Photo.FullPath -Parent) $newFileName
            if (Test-Path -LiteralPath $newFullPath) {
                [System.Windows.MessageBox]::Show("Un fichier '$newFileName' existe deja.", "Photo Viewer", 'OK', 'Warning') | Out-Null
                return
            }
            try {
                Rename-Item -LiteralPath $Photo.FullPath -NewName $newFileName -ErrorAction Stop
                $Photo.FullPath = $newFullPath
                $Photo.FileName = $newFileName
            } catch {
                [System.Windows.MessageBox]::Show("Erreur lors du renommage : $_", "Photo Viewer", 'OK', 'Error') | Out-Null
                return
            }
        }

        $popup.DialogResult = $true
    })

    $popup.ShowDialog() | Out-Null
}

# ------------------------------------------------------------------
# Mode plein ecran (etape 5) : navigation swipe/clavier, zoom/pan tactile, nom cliquable
# reutilisant la popup de renommage. $allPhotos represente toujours le repertoire affiche.
# ------------------------------------------------------------------
$script:fullScreenIndex = -1
$script:fullScreenGate = New-Object PhotoViewer.Services.GenerationGate

function Update-FullScreenImage {
    if ($script:fullScreenIndex -lt 0 -or $script:fullScreenIndex -ge $allPhotos.Count) { return }
    $photo = $allPhotos[$script:fullScreenIndex]
    $fullScreenFileNameButton.Content = $photo.FileName
    [PhotoViewer.Gestures.ZoomPanGesture]::Reset($fullScreenScale, $fullScreenTranslate)
    $fullScreenImage.Source = $null
    $generation = $script:fullScreenGate.Next()
    $onLoaded = { param($bmp) $fullScreenImage.Source = $bmp }
    [PhotoViewer.Services.ThumbnailLoader]::LoadFullImageAsync($photo.FullPath, 1920, $uiDispatcher, $script:fullScreenGate, $generation, [Action[System.Windows.Media.Imaging.BitmapImage]]$onLoaded)
}

function Show-FullScreenPhoto {
    param([PhotoViewer.Models.PhotoItem]$Photo)
    $index = $allPhotos.IndexOf($Photo)
    if ($index -lt 0) { return }
    $script:fullScreenIndex = $index
    $fullScreenOverlay.Visibility = 'Visible'
    Update-FullScreenImage
    $fullScreenOverlay.Focus() | Out-Null
}

function Hide-FullScreenPhoto {
    $fullScreenOverlay.Visibility = 'Collapsed'
    $fullScreenImage.Source = $null
    $script:fullScreenIndex = -1
}

function Show-NextFullScreenPhoto {
    if ($allPhotos.Count -eq 0) { return }
    $script:fullScreenIndex = ($script:fullScreenIndex + 1) % $allPhotos.Count
    Update-FullScreenImage
}

function Show-PreviousFullScreenPhoto {
    if ($allPhotos.Count -eq 0) { return }
    $script:fullScreenIndex = ($script:fullScreenIndex - 1 + $allPhotos.Count) % $allPhotos.Count
    Update-FullScreenImage
}

# ------------------------------------------------------------------
# Selection d'un dossier via l'explorateur Windows natif (aucune creation automatique)
# ------------------------------------------------------------------
function Select-Folder {
    param([string]$InitialPath, [string]$Description)
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Description
    $dialog.ShowNewFolderButton = $true
    if ($InitialPath -and (Test-Path -LiteralPath $InitialPath -PathType Container)) {
        $dialog.SelectedPath = $InitialPath
    }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.SelectedPath
    }
    return $null
}

# ------------------------------------------------------------------
# ViewModel + commande de tap
# ------------------------------------------------------------------
$viewModel = New-Object PhotoViewer.ViewModels.MainViewModel

$tapAction = {
    param($photo)
    # Sur le repertoire "existant" (lecture seule), le tap n'ouvre pas la popup de renommage :
    # l'appui long (mode plein ecran, etape 5) reste le moyen de consulter la photo en detail.
    if ($viewModel.IsReadOnly) { return }
    Show-RenamePopup -Photo $photo
}
$viewModel.TapCommand = New-Object PhotoViewer.Commands.RelayCommand ([Action[object]]$tapAction)

$toggleDirectoryAction = {
    param($unused)
    $script:currentDirectoryIndex = ($script:currentDirectoryIndex + 1) % $script:directoryConfigs.Count
    Load-CurrentDirectory
}
$viewModel.ToggleDirectoryCommand = New-Object PhotoViewer.Commands.RelayCommand ([Action[object]]$toggleDirectoryAction)

$chooseDirectoryAction = {
    param($unused)
    $config = $script:directoryConfigs[$script:currentDirectoryIndex]
    $description = if ($config.IsReadOnly) { "Choisir le repertoire 'existant' (lecture seule)" } else { "Choisir le repertoire 'fraiches' (lecture + ecriture)" }
    $picked = Select-Folder -InitialPath $config.Path -Description $description
    if ($picked) {
        $script:directoryConfigs[$script:currentDirectoryIndex].Path = $picked
        Save-DirectoryConfig
        Load-CurrentDirectory
        if ($script:currentDirectoryIndex -eq 0) { Start-FreshDirectoryWatcher }
    }
}
$viewModel.ChooseDirectoryCommand = New-Object PhotoViewer.Commands.RelayCommand ([Action[object]]$chooseDirectoryAction)

$window.DataContext = $viewModel

# ------------------------------------------------------------------
# Chargement de la liste des photos (repertoire actif, rechargeable via le switch)
# ------------------------------------------------------------------
$imageExtensions = @('.jpg', '.jpeg', '.png', '.bmp', '.gif')
$allPhotos = New-Object 'System.Collections.ObjectModel.ObservableCollection[PhotoViewer.Models.PhotoItem]'

# ------------------------------------------------------------------
# Regroupement en lignes (pour la virtualisation native de VirtualizingStackPanel)
# ------------------------------------------------------------------
$rows = New-Object 'System.Collections.ObjectModel.ObservableCollection[System.Collections.ObjectModel.ObservableCollection[PhotoViewer.Models.PhotoItem]]'
$rowsItemsControl.ItemsSource = $rows
$script:currentColumnCount = 0

function Update-PhotoRows {
    param([int]$Columns, [switch]$Force)
    if ($Columns -lt 1) { $Columns = 1 }
    if (-not $Force -and $Columns -eq $script:currentColumnCount) { return }
    $script:currentColumnCount = $Columns

    $rows.Clear()
    $currentRow = $null
    for ($i = 0; $i -lt $allPhotos.Count; $i++) {
        if ($i % $Columns -eq 0) {
            $currentRow = New-Object 'System.Collections.ObjectModel.ObservableCollection[PhotoViewer.Models.PhotoItem]'
            $rows.Add($currentRow)
        }
        $currentRow.Add($allPhotos[$i])
    }
}

$scrollViewer.Add_SizeChanged({
    $columns = [Math]::Max(1, [Math]::Floor($scrollViewer.ActualWidth / $ThumbnailCellWidth))
    Update-PhotoRows -Columns $columns
})

# ------------------------------------------------------------------
# Chargement asynchrone des miniatures : delegue entierement a du C# compile (voir
# PhotoViewer.Services.ThumbnailLoader) - un ScriptBlock PowerShell ne peut pas s'executer
# sur un thread brut du pool .NET (pas de Runspace attache), d'ou l'usage exclusif de C# ici.
# ------------------------------------------------------------------
$script:generationGate = New-Object PhotoViewer.Services.GenerationGate
[PhotoViewer.Services.ThumbnailLoader]::Configure($MaxConcurrentThumbnailLoads)
$script:errorLogPath = Join-Path $PSScriptRoot 'PhotoViewer.errors.log'

function Start-ThumbnailLoading {
    param([int]$Generation)
    foreach ($photo in $allPhotos) {
        [PhotoViewer.Services.ThumbnailLoader]::LoadAsync($photo, $ThumbnailDecodePixelWidth, $uiDispatcher, $script:generationGate, $Generation, $script:errorLogPath)
    }
}

# ------------------------------------------------------------------
# (Re)chargement du repertoire actif : vide la grille, relit le disque, relance les miniatures.
# Le compteur de generation invalide les chargements encore en vol depuis un repertoire quitte.
# ------------------------------------------------------------------
function Load-CurrentDirectory {
    $config = $script:directoryConfigs[$script:currentDirectoryIndex]
    $generation = $script:generationGate.Next()

    if ($fullScreenOverlay.Visibility -eq 'Visible') { Hide-FullScreenPhoto }

    $viewModel.DirectoryLabel = $config.Path
    $viewModel.IsReadOnly = $config.IsReadOnly

    $allPhotos.Clear()
    $rows.Clear()

    if (Test-Path -LiteralPath $config.Path -PathType Container) {
        $files = Get-ChildItem -LiteralPath $config.Path -File |
            Where-Object { $imageExtensions -contains $_.Extension.ToLowerInvariant() } |
            Sort-Object Name

        foreach ($file in $files) {
            $photo = New-Object PhotoViewer.Models.PhotoItem
            $photo.FullPath = $file.FullName
            $photo.FileName = $file.Name
            $allPhotos.Add($photo)
        }
    }

    $columns = [Math]::Max(1, [Math]::Floor($scrollViewer.ActualWidth / $ThumbnailCellWidth))
    Update-PhotoRows -Columns $columns -Force
    Start-ThumbnailLoading -Generation $generation
}

# ------------------------------------------------------------------
# Detection automatique des nouvelles photos dans le repertoire "fraiches" (voir
# PhotoViewer.Services.DirectoryWatcher - callback invoque sur le thread UI, donc du PowerShell y est sans risque).
# ------------------------------------------------------------------
$script:freshWatcher = $null

function Start-FreshDirectoryWatcher {
    if ($script:freshWatcher) {
        $script:freshWatcher.EnableRaisingEvents = $false
        $script:freshWatcher.Dispose()
        $script:freshWatcher = $null
    }

    $freshPath = $script:directoryConfigs[0].Path
    if (-not (Test-Path -LiteralPath $freshPath -PathType Container)) { return }

    $onFreshChanged = {
        if ($script:currentDirectoryIndex -eq 0) { Load-CurrentDirectory }
    }
    $script:freshWatcher = [PhotoViewer.Services.DirectoryWatcher]::Start($freshPath, $uiDispatcher, [Action]$onFreshChanged, 600)
}

Load-CurrentDirectory
Start-FreshDirectoryWatcher

$refreshButton.Add_Click({ param($s, $e) Load-CurrentDirectory })

# ------------------------------------------------------------------
# Mode plein ecran : appui long depuis la grille, boutons prev/suivant/fermer,
# nom cliquable (rename), Echap/fleches clavier, zoom/pan/swipe tactile.
# ------------------------------------------------------------------
$longPressAction = { param($photo) Show-FullScreenPhoto -Photo $photo }
[PhotoViewer.Gestures.LongPressGesture]::Attach($scrollViewer, [Action[PhotoViewer.Models.PhotoItem]]$longPressAction, $LongPressThresholdMs)

$fullScreenPrevButton.Add_Click({ param($s, $e) Show-PreviousFullScreenPhoto })
$fullScreenNextButton.Add_Click({ param($s, $e) Show-NextFullScreenPhoto })
$fullScreenCloseButton.Add_Click({ param($s, $e) Hide-FullScreenPhoto })

$fullScreenFileNameButton.Add_Click({
    param($s, $e)
    if ($viewModel.IsReadOnly) { return }
    if ($script:fullScreenIndex -lt 0 -or $script:fullScreenIndex -ge $allPhotos.Count) { return }
    $photo = $allPhotos[$script:fullScreenIndex]
    Show-RenamePopup -Photo $photo
    if ($allPhotos.Count -eq 0) { Hide-FullScreenPhoto; return }
    if ($script:fullScreenIndex -ge $allPhotos.Count) { $script:fullScreenIndex = $allPhotos.Count - 1 }
    Update-FullScreenImage
})

$fullScreenOverlay.Add_PreviewKeyDown({
    param($s, $e)
    switch ($e.Key) {
        'Escape' { Hide-FullScreenPhoto; $e.Handled = $true }
        'Left'   { Show-PreviousFullScreenPhoto; $e.Handled = $true }
        'Right'  { Show-NextFullScreenPhoto; $e.Handled = $true }
    }
})

# Swipe gauche -> photo suivante, swipe droite -> photo precedente (uniquement hors zoom).
[PhotoViewer.Gestures.ZoomPanGesture]::Attach(
    $fullScreenImageContainer, $fullScreenScale, $fullScreenTranslate,
    [Action]{ Show-NextFullScreenPhoto }, [Action]{ Show-PreviousFullScreenPhoto })

# ------------------------------------------------------------------
# Lancement de l'application
# ------------------------------------------------------------------
$app = New-Object System.Windows.Application
$app.Run($window) | Out-Null
