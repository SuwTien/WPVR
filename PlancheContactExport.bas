Attribute VB_Name = "PlancheContactExport"
Option Explicit

' ============================================================================
' Export "planche contact" Excel - chantier INDEPENDANT du viewer WPF
' ============================================================================
'
' Objectif : scanner le repertoire "photos fraiches", et generer un NOUVEAU
' classeur Excel autonome (jamais le classeur macro existant) contenant,
' pour chaque photo : nom de fichier / image / horodatage (date de prise de
' vue EXIF si disponible, sinon date de modification du fichier).
'
' INSTALLATION
' ------------
'   1. Alt+F11 > Fichier > Importer un fichier... > selectionner ce fichier
'      PlancheContactExport.bas (ou creer un nouveau module et coller le
'      code a partir de "Option Explicit").
'   2. Modifier la constante FRESH_PHOTOS_DIRECTORY ci-dessous avec le
'      chemin reel du repertoire "photos fraiches" (le meme dossier que
'      celui configure dans le viewer WPF, a mettre a jour manuellement si
'      tu changes de dossier via le bouton "Dossier..." du viewer - ce
'      module n'a volontairement aucune dependance technique avec l'autre
'      chantier).
'   3. Rattacher la macro "ExporterPlancheContact" a un bouton du bandeau
'      personnalise (a cote de celui qui lance le viewer) : Fichier >
'      Options > Personnaliser le ruban > nouveau bouton > categorie
'      "Macros" > "ExporterPlancheContact".
'
' EN CAS DE PROBLEME
' -------------------
'   - Repertoire introuvable : verifier FRESH_PHOTOS_DIRECTORY.
'   - Date de prise de vue absente/incorrecte : la lecture EXIF passe par
'     l'API Windows Shell (Shell.Application), toujours presente sur Windows ;
'     si une photo n'a pas de metadonnee EXIF (capture d'ecran, image
'     retouchee...), la date de modification du fichier est utilisee a la
'     place, automatiquement.
'   - Fichier de sortie : nomme automatiquement avec un horodatage a la
'     seconde pres (PlancheContact_AAAA-MM-JJ_HHMMSS.xlsx, dans le dossier
'     Documents de l'utilisateur) - une collision est quasi impossible, mais
'     un suffixe numerique est ajoute par securite si le nom existe deja.
'   - Beaucoup de photos volumineuses : le redimensionnement (MAX_IMAGE_WIDTH_PT)
'     limite la taille des images inserees pour eviter un fichier trop lourd ;
'     augmenter cette constante si les vignettes sont trop petites a l'usage.
' ============================================================================

Private Const FRESH_PHOTOS_DIRECTORY As String = "C:\Chemin\Vers\PhotosFraiches"
Private Const MAX_IMAGE_WIDTH_PT As Double = 220   ' ~7.8 cm, lisible sans etre trop lourd

Public Sub ExporterPlancheContact()
    Dim sourceDir As String
    Dim fso As Object
    Dim f As Object
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim pic As Shape
    Dim destPath As String
    Dim rowIndex As Long
    Dim exportedCount As Long
    Dim captureDate As Date

    sourceDir = FRESH_PHOTOS_DIRECTORY
    If Right(sourceDir, 1) <> "\" Then sourceDir = sourceDir & "\"

    If Dir(sourceDir, vbDirectory) = vbNullString Then
        MsgBox "Repertoire introuvable :" & vbCrLf & sourceDir & vbCrLf & vbCrLf & _
               "Verifiez la constante FRESH_PHOTOS_DIRECTORY dans le module VBA " & _
               "PlancheContactExport.", vbExclamation, "Export planche contact"
        Exit Sub
    End If

    destPath = ChoisirNomFichierUnique()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo CleanFail

    Set wbNew = Workbooks.Add(xlWBATWorksheet)
    Set wsNew = wbNew.Worksheets(1)
    wsNew.Name = "Planche contact"

    wsNew.Cells(1, 1).Value = "Nom du fichier"
    wsNew.Cells(1, 2).Value = "Photo"
    wsNew.Cells(1, 3).Value = "Horodatage"
    wsNew.Range("A1:C1").Font.Bold = True
    wsNew.Columns(2).ColumnWidth = 32

    Set fso = CreateObject("Scripting.FileSystemObject")
    rowIndex = 2
    exportedCount = 0

    For Each f In fso.GetFolder(sourceDir).Files
        If EstUneImage(f.Name) Then
            wsNew.Cells(rowIndex, 1).Value = f.Name

            Set pic = wsNew.Shapes.AddPicture(f.Path, msoFalse, msoTrue, _
                wsNew.Cells(rowIndex, 2).Left + 2, wsNew.Cells(rowIndex, 2).Top + 2, -1, -1)

            If pic.Width > MAX_IMAGE_WIDTH_PT Then
                pic.Height = pic.Height * (MAX_IMAGE_WIDTH_PT / pic.Width)
                pic.Width = MAX_IMAGE_WIDTH_PT
            End If
            wsNew.Rows(rowIndex).RowHeight = pic.Height + 4

            captureDate = ObtenirDatePriseDeVue(f.Path)
            If captureDate = 0 Then captureDate = f.DateLastModified
            wsNew.Cells(rowIndex, 3).Value = captureDate
            wsNew.Cells(rowIndex, 3).NumberFormat = "dd/mm/yyyy hh:mm"

            rowIndex = rowIndex + 1
            exportedCount = exportedCount + 1
        End If
    Next f

    wsNew.Columns(1).AutoFit
    wsNew.Columns(3).AutoFit

    wbNew.SaveAs Filename:=destPath, FileFormat:=51 ' xlOpenXMLWorkbook (.xlsx)

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox exportedCount & " photo(s) exportee(s) dans :" & vbCrLf & destPath, _
           vbInformation, "Export planche contact"
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Erreur pendant l'export : " & Err.Description, vbCritical, "Export planche contact"
End Sub

' Nom de fichier horodate a la seconde pres ; ajoute un suffixe numerique si,
' par extraordinaire, ce nom existe deja (deux exports dans la meme seconde).
Private Function ChoisirNomFichierUnique() As String
    Dim baseName As String, candidate As String, suffix As Integer
    baseName = Environ$("USERPROFILE") & "\Documents\PlancheContact_" & Format(Now, "yyyy-mm-dd_hhnnss")
    candidate = baseName & ".xlsx"
    suffix = 1
    Do While Dir(candidate) <> vbNullString
        candidate = baseName & "_" & suffix & ".xlsx"
        suffix = suffix + 1
    Loop
    ChoisirNomFichierUnique = candidate
End Function

Private Function EstUneImage(ByVal fileName As String) As Boolean
    Dim ext As String
    Dim dotPos As Long
    dotPos = InStrRev(fileName, ".")
    If dotPos = 0 Then
        EstUneImage = False
        Exit Function
    End If
    ext = LCase(Mid(fileName, dotPos + 1))
    Select Case ext
        Case "jpg", "jpeg", "png", "bmp", "gif"
            EstUneImage = True
        Case Else
            EstUneImage = False
    End Select
End Function

' Date de prise de vue EXIF via l'API Windows Shell (Shell.Application). Retourne 0
' si indisponible (photo sans EXIF, format non supporte...) - l'appelant doit alors
' se rabattre sur la date de modification du fichier.
Private Function ObtenirDatePriseDeVue(ByVal filePath As String) As Date
    On Error GoTo Fallback
    Dim shellApp As Object, objFolder As Object, objFile As Object
    Dim folderPath As String, fileName As String
    Dim i As Integer, dateTakenCol As Integer
    Dim rawValue As String

    folderPath = Left(filePath, InStrRev(filePath, "\"))
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)

    Set shellApp = CreateObject("Shell.Application")
    Set objFolder = shellApp.Namespace(folderPath)
    If objFolder Is Nothing Then GoTo Fallback
    Set objFile = objFolder.ParseName(fileName)
    If objFile Is Nothing Then GoTo Fallback

    dateTakenCol = -1
    For i = 0 To 300
        Dim header As String
        header = objFolder.GetDetailsOf(Nothing, i)
        If header = "Date taken" Or header = "Date de prise de vue" Then
            dateTakenCol = i
            Exit For
        End If
    Next i
    If dateTakenCol = -1 Then GoTo Fallback

    rawValue = NettoyerChaineDateShell(objFolder.GetDetailsOf(objFile, dateTakenCol))
    If Trim(rawValue) = "" Then GoTo Fallback

    ObtenirDatePriseDeVue = CDate(rawValue)
    Exit Function

Fallback:
    ObtenirDatePriseDeVue = 0
End Function

' L'API Shell renvoie souvent des marqueurs Unicode invisibles (LRM, espaces
' insecables...) autour de la date, qui font echouer CDate() si on ne les retire pas.
Private Function NettoyerChaineDateShell(ByVal s As String) As String
    Dim i As Long, ch As String, result As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If (AscW(ch) >= 32 And AscW(ch) < 127) Then
            result = result & ch
        End If
    Next i
    NettoyerChaineDateShell = Trim(result)
End Function
