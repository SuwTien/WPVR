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
using System.Runtime.InteropServices;
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

            Action<DependencyObject, Point> startPress = (source, pos) =>
            {
                var element = FindPhotoItemElement(source);
                if (element == null) return;

                pressedItem = element.DataContext as PhotoItem;
                downPos = pos;

                timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(thresholdMs) };
                timer.Tick += (s, e) =>
                {
                    timer.Stop();
                    var item = pressedItem;
                    pressedItem = null;
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
            root.PreviewMouseLeftButtonUp += (s, e) => cancelPress();
            root.PreviewTouchDown += (s, e) => startPress(e.OriginalSource as DependencyObject, e.GetTouchPoint(root).Position);
            root.PreviewTouchUp += (s, e) => cancelPress();
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
        public static void Toggle(IntPtr ownerHandle)
        {
            try
            {
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

$uiDispatcher = $window.Dispatcher

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
    $toggleKeyboardButton = $popup.FindName('ToggleKeyboardButton')
    $numericKeypadPanel = $popup.FindName('NumericKeypadPanel')

    $extension = [System.IO.Path]::GetExtension($Photo.FileName)
    $nameTextBox.Text = [System.IO.Path]::GetFileNameWithoutExtension($Photo.FileName)
    $extensionLabel.Text = $extension
    $nameTextBox.IsReadOnly = $true
    $nameTextBox.Focusable = $false
    $nameTextBox.SelectAll()

    # HWND de la popup (force la creation du handle avant meme l'affichage) pour piloter le clavier tactile.
    $popupHandleHelper = New-Object System.Windows.Interop.WindowInteropHelper($popup)
    $popupHandleHelper.EnsureHandle() | Out-Null
    $popupHandle = $popupHandleHelper.Handle

    # Pave numerique custom : chaque bouton porte son chiffre (ou BACK) dans Tag.
    foreach ($child in $numericKeypadPanel.Children) {
        if ($child -isnot [System.Windows.Controls.Button]) { continue }
        if (-not $child.Tag) { continue }
        $child.Add_Click({
            param($s, $e)
            $key = $s.Tag.ToString()
            if ($key -eq 'BACK') {
                if ($nameTextBox.Text.Length -gt 0) {
                    $nameTextBox.Text = $nameTextBox.Text.Substring(0, $nameTextBox.Text.Length - 1)
                }
            } else {
                $nameTextBox.Text += $key
            }
            $nameTextBox.CaretIndex = $nameTextBox.Text.Length
        }.GetNewClosure())
    }

    # Bascule vers un clavier texte complet (systeme) quand des lettres sont necessaires.
    # L'etat (numerique/texte) est lu directement sur l'UI (Visibility du pave) plutot que sur
    # une variable a part : plus fiable qu'un booleen capture dans une closure de bouton.
    $toggleKeyboardButton.Add_Click({
        param($s, $e)
        $switchingToText = $numericKeypadPanel.Visibility -eq 'Visible'
        if ($switchingToText) {
            $numericKeypadPanel.Visibility = 'Collapsed'
            $nameTextBox.IsReadOnly = $false
            $nameTextBox.Focusable = $true
            $nameTextBox.Focus()
            $nameTextBox.SelectAll()
            $toggleKeyboardButton.Content = 'Clavier numerique (123)'
        } else {
            $numericKeypadPanel.Visibility = 'Visible'
            $nameTextBox.IsReadOnly = $true
            $nameTextBox.Focusable = $false
            $toggleKeyboardButton.Content = 'Clavier texte (ABC)'
        }
        [PhotoViewer.Interop.TouchKeyboard]::Toggle($popupHandle)
    }.GetNewClosure())

    # Bouton corbeille : placeholder, la suppression reelle sera branchee a l'etape 4.
    $trashButton.Add_Click({
        param($s, $e)
        [System.Windows.MessageBox]::Show(
            "Suppression (placeholder) : $($Photo.FileName)`n(la suppression reelle sera branchee a l'etape 4)",
            "Photo Viewer", 'OK', 'Information') | Out-Null
    }.GetNewClosure())

    $cancelButton.Add_Click({
        param($s, $e)
        $popup.DialogResult = $false
    }.GetNewClosure())

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
    }.GetNewClosure())

    $popup.ShowDialog() | Out-Null
}

function Show-LongPressPlaceholder {
    param([PhotoViewer.Models.PhotoItem]$Photo)
    # Validation du mecanisme d'appui long uniquement (le vrai plein ecran arrive a l'etape 5).
    [System.Windows.MessageBox]::Show(
        "Appui long detecte (test) : $($Photo.FileName)`n(le mode plein ecran sera branche a l'etape 5)",
        "Photo Viewer", 'OK', 'Information') | Out-Null
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
# Chargement asynchrone des miniatures (concurrence limitee via SemaphoreSlim)
# ------------------------------------------------------------------
$semaphore = New-Object System.Threading.SemaphoreSlim($MaxConcurrentThumbnailLoads, $MaxConcurrentThumbnailLoads)
$script:loadGeneration = 0

function Start-ThumbnailLoading {
    param([int]$Generation)
    foreach ($photo in $allPhotos) {
        $loadThumbnail = {
            if ($script:loadGeneration -ne $Generation) { return }
            $semaphore.Wait()
            try {
                if ($script:loadGeneration -ne $Generation) { return }
                $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
                $bmp.BeginInit()
                $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                $bmp.UriSource = New-Object System.Uri($photo.FullPath)
                $bmp.DecodePixelWidth = $ThumbnailDecodePixelWidth
                $bmp.EndInit()
                $bmp.Freeze()
                $uiDispatcher.Invoke([Action]{
                    if ($script:loadGeneration -eq $Generation) { $photo.Thumbnail = $bmp }
                })
            } catch {
                # Fichier illisible/corrompu : on laisse la vignette vide plutot que de planter le thread.
            } finally {
                $semaphore.Release() | Out-Null
            }
        }.GetNewClosure()

        [System.Threading.Tasks.Task]::Run([Action]$loadThumbnail) | Out-Null
    }
}

# ------------------------------------------------------------------
# (Re)chargement du repertoire actif : vide la grille, relit le disque, relance les miniatures.
# Le compteur de generation invalide les chargements encore en vol depuis un repertoire quitte.
# ------------------------------------------------------------------
function Load-CurrentDirectory {
    $config = $script:directoryConfigs[$script:currentDirectoryIndex]
    $script:loadGeneration++
    $generation = $script:loadGeneration

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

Load-CurrentDirectory

# ------------------------------------------------------------------
# Validation Add-Type : detection d'appui long sur la grille (voir Gestures.LongPressGesture)
# ------------------------------------------------------------------
$longPressAction = { param($photo) Show-LongPressPlaceholder -Photo $photo }
[PhotoViewer.Gestures.LongPressGesture]::Attach($scrollViewer, [Action[PhotoViewer.Models.PhotoItem]]$longPressAction, $LongPressThresholdMs)

# ------------------------------------------------------------------
# Lancement de l'application
# ------------------------------------------------------------------
$app = New-Object System.Windows.Application
$app.Run($window) | Out-Null
