#Requires -Version 5.1
<#
    Export-ContactSheet.ps1 - Genere une "planche contact" Excel a partir d'un dossier de photos,
    SANS macro VBA ni integration dans un classeur existant. Utilise l'automation COM d'Excel
    (Excel doit etre installe sur la machine - c'est deja le cas puisque le classeur existant
    l'utilise). Peut etre lance seul (double-clic / powershell.exe -File) ou depuis le bouton
    "Export planche contact" du viewer (PhotoViewer.ps1).

    IMPORTANT : comme PhotoViewer.ps1, a lancer via "powershell.exe" (STA), pas "pwsh.exe" -
    l'automation COM Excel exige un thread STA, tout comme WPF.
#>
param(
    # Fourni directement par le bouton du viewer (chemin du repertoire "fraiches" actif) pour
    # eviter une popup redondante avec le SaveFileDialog. En lancement autonome (sans ce parametre),
    # le script tente PhotoViewer.config.json, et ne propose le selecteur de dossier qu'en dernier recours.
    [string]$SourceDirectory
)

Add-Type -AssemblyName PresentationCore, WindowsBase, System.Windows.Forms -ErrorAction SilentlyContinue

$script:MaxImageWidthPoints = 220   # ~7.8 cm, lisible sans etre trop lourd
$script:ImageExtensions = @('.jpg', '.jpeg', '.png', '.bmp', '.gif')
$script:XlCenter = -4108   # xlCenter (horizontal et vertical) - constante Excel non exposee en late-binding COM

# Ajuste la largeur d'une colonne pour approcher une valeur cible en POINTS. ColumnWidth
# s'exprime en "caracteres" (unite dependante de la police), pas en points, contrairement a
# la largeur des images (Shape.Width, en points) : sans cette conversion, la colonne ne
# correspond pas a la largeur reelle des photos inserees. On affine par iteration car la
# relation caracteres/points n'est pas parfaitement lineaire.
function Set-ColumnWidthInPoints {
    param($Column, [double]$TargetPoints)
    $Column.ColumnWidth = $TargetPoints / 7
    for ($i = 0; $i -lt 4; $i++) {
        $diff = $TargetPoints - $Column.Width
        if ([Math]::Abs($diff) -lt 0.5) { break }
        $Column.ColumnWidth = $Column.ColumnWidth + ($diff / 7)
    }
}

# Date de prise de vue EXIF via les metadonnees WPF (BitmapMetadata.DateTaken) - pas besoin
# de decoder l'image entiere. Retourne $null si absente (l'appelant se rabat alors sur la
# date de modification du fichier).
function Get-PhotoTakenDate {
    param([Parameter(Mandatory)][string]$Path)
    try {
        $stream = [System.IO.File]::OpenRead($Path)
        try {
            $decoder = [System.Windows.Media.Imaging.BitmapDecoder]::Create(
                $stream,
                [System.Windows.Media.Imaging.BitmapCreateOptions]::DelayCreation,
                [System.Windows.Media.Imaging.BitmapCacheOption]::None)
            $metadata = $decoder.Frames[0].Metadata -as [System.Windows.Media.Imaging.BitmapMetadata]
            if ($metadata -and $metadata.DateTaken) {
                return [datetime]::Parse($metadata.DateTaken)
            }
        } finally {
            $stream.Close()
        }
    } catch {
        # Pas de metadonnee EXIF exploitable (format non supporte, image retouchee...).
    }
    return $null
}

# Genere le classeur Excel autonome. Ouvre/pilote Excel via COM, le ferme et libere
# proprement les objets COM dans tous les cas (evite un EXCEL.EXE fantome en arriere-plan).
function Export-ContactSheet {
    param(
        [Parameter(Mandatory)][string]$SourceDirectory,
        [Parameter(Mandatory)][string]$DestinationPath
    )

    $files = Get-ChildItem -LiteralPath $SourceDirectory -File |
        Where-Object { $script:ImageExtensions -contains $_.Extension.ToLowerInvariant() } |
        Sort-Object Name

    if ($files.Count -eq 0) {
        throw "Aucune photo trouvee dans '$SourceDirectory'."
    }

    $existingExcelPids = @((Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Id)

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $null
    $sheet = $null

    try {
        $workbook = $excel.Workbooks.Add()
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Name = 'Planche contact'
        $sheet.Cells.Item(1, 1) = 'Nom du fichier'
        $sheet.Cells.Item(1, 2) = 'Photo'
        $sheet.Cells.Item(1, 3) = 'Horodatage'
        $sheet.Range('A1:C1').Font.Bold = $true
        $sheet.Range('A1:C1').HorizontalAlignment = $script:XlCenter
        $sheet.Range('A1:C1').VerticalAlignment = $script:XlCenter
        Set-ColumnWidthInPoints -Column $sheet.Columns.Item(2) -TargetPoints ($script:MaxImageWidthPoints + 8)

        $row = 2
        foreach ($file in $files) {
            $sheet.Cells.Item($row, 1) = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

            $cellLeft = $sheet.Cells.Item($row, 2).Left
            $cellTop = $sheet.Cells.Item($row, 2).Top
            $pic = $sheet.Shapes.AddPicture($file.FullName, 0, 1, $cellLeft + 2, $cellTop + 2, -1, -1)
            if ($pic.Width -gt $script:MaxImageWidthPoints) {
                $pic.Height = $pic.Height * ($script:MaxImageWidthPoints / $pic.Width)
                $pic.Width = $script:MaxImageWidthPoints
            }
            $sheet.Rows.Item($row).RowHeight = $pic.Height + 6

            $takenDate = Get-PhotoTakenDate -Path $file.FullName
            if (-not $takenDate) { $takenDate = $file.LastWriteTime }
            $sheet.Cells.Item($row, 3) = $takenDate.ToString('dd/MM/yyyy HH:mm')

            $row++
        }

        $dataRange = $sheet.Range("A2:C$($row - 1)")
        $dataRange.HorizontalAlignment = $script:XlCenter
        $dataRange.VerticalAlignment = $script:XlCenter

        $sheet.Columns.Item(1).AutoFit() | Out-Null
        $sheet.Columns.Item(3).AutoFit() | Out-Null

        $workbook.SaveAs($DestinationPath)
        $workbook.Close($false)
    } finally {
        $excel.Quit()
        if ($sheet) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null }
        if ($workbook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        # Filet de securite : des references COM implicites (Cells.Item, Range, Rows.Item...) empechent
        # souvent Excel de vraiment quitter malgre Quit()/ReleaseComObject. On identifie le process EXCEL.EXE
        # qu'on vient de lancer (absent de la liste au demarrage) et on le ferme de force s'il traine encore.
        Start-Sleep -Milliseconds 500
        $newExcelPids = @((Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Id) | Where-Object { $existingExcelPids -notcontains $_ }
        foreach ($orphanPid in $newExcelPids) {
            Stop-Process -Id $orphanPid -Force -ErrorAction SilentlyContinue
        }
    }

    return $files.Count
}

# ------------------------------------------------------------------
# Flux interactif : popup dossier source, puis popup fichier de sortie.
# ------------------------------------------------------------------
# Determination du repertoire source, sans redemander si on le connait deja :
# 1) parametre -SourceDirectory (fourni par le bouton du viewer) ;
# 2) a defaut, PhotoViewer.config.json a cote de ce script (lancement autonome) ;
# 3) en tout dernier recours seulement, un selecteur de dossier manuel.
# ------------------------------------------------------------------
if (-not $SourceDirectory -or -not (Test-Path -LiteralPath $SourceDirectory -PathType Container)) {
    $configPath = Join-Path $PSScriptRoot 'PhotoViewer.config.json'
    if (Test-Path -LiteralPath $configPath -PathType Leaf) {
        try {
            $config = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
            if ($config.FreshPhotoDirectory -and (Test-Path -LiteralPath $config.FreshPhotoDirectory -PathType Container)) {
                $SourceDirectory = $config.FreshPhotoDirectory
            }
        } catch {
            # Config illisible/absente : on retombe sur la selection manuelle ci-dessous.
        }
    }
}

if (-not $SourceDirectory -or -not (Test-Path -LiteralPath $SourceDirectory -PathType Container)) {
    $sourceFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $sourceFolderDialog.Description = 'Choisir le dossier de photos a exporter'
    if ($sourceFolderDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $SourceDirectory = $sourceFolderDialog.SelectedPath
}

$saveDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveDialog.Title = 'Enregistrer la planche contact'
$saveDialog.Filter = 'Classeur Excel (*.xlsx)|*.xlsx'
$saveDialog.DefaultExt = 'xlsx'
$saveDialog.FileName = "PlancheContact_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').xlsx"
$saveDialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
if ($saveDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

try {
    $count = Export-ContactSheet -SourceDirectory $SourceDirectory -DestinationPath $saveDialog.FileName
    [System.Windows.Forms.MessageBox]::Show(
        "$count photo(s) exportee(s) dans :`n$($saveDialog.FileName)",
        'Export planche contact', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Erreur pendant l'export : $_",
        'Export planche contact', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
}
