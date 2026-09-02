Attribute VB_Name = "PhotoViewerLauncher"
Option Explicit

' ============================================================================
' Photo Viewer - Lancement depuis Excel (bouton n4 du bandeau personnalise)
' ============================================================================
'
' INSTALLATION
' ------------
'   1. Dans Excel : Alt+F11 pour ouvrir l'editeur VBA.
'   2. Menu Fichier > Importer un fichier... > selectionner ce fichier
'      PhotoViewerLauncher.bas (ou : clic droit sur le projet VBA du classeur
'      > Insertion > Module, puis coller le code a partir de "Option Explicit").
'   3. Modifier la constante PHOTO_VIEWER_SCRIPT_PATH ci-dessous avec le
'      chemin reel du fichier PhotoViewer.ps1 sur la machine cible.
'   4. Enregistrer le classeur en .xlsm (macros activees) si ce n'est pas
'      deja le cas.
'   5. Rattacher la macro "LancerPhotoViewer" au 4e bouton du bandeau :
'        - Fichier > Options > Personnaliser le ruban
'          (ou clic droit sur le ruban > Personnaliser le ruban...).
'        - Retrouver l'onglet/groupe du bandeau personnalise deja utilise
'          pour les 3 autres boutons existants.
'        - Nouveau bouton > categorie "Macros" > choisir "LancerPhotoViewer".
'      Si le bandeau est genere via un XML personnalise (customUI.xml, cas
'      des rubans plus avances), ajouter un element bouton pointant vers
'      cette macro, par exemple :
'        <button id="btnPhotoViewer" label="Photo Viewer" imageMso="..."
'                onAction="LancerPhotoViewer" size="large"/>
'
' EN CAS D'ECHEC AU LANCEMENT - POINTS A VERIFIER
' ------------------------------------------------
'   - Chemin : PHOTO_VIEWER_SCRIPT_PATH doit pointer vers le vrai fichier
'     PhotoViewer.ps1 (chemin absolu recommande, pas de chemin relatif).
'   - Macros Excel desactivees : Fichier > Options > Centre de gestion de la
'     confidentialite > Parametres du Centre de gestion de la confidentialite
'     > Parametres des macros. Sur un poste d'entreprise, verifier aupres de
'     l'IT si le classeur doit etre place dans un emplacement de confiance.
'   - ExecutionPolicy PowerShell restrictive (Restricted / AllSigned) :
'     le parametre "-ExecutionPolicy Bypass" utilise ci-dessous contourne ceci
'     PROPREMENT - il ne s'applique qu'a ce lancement precis du process
'     PowerShell, ne modifie AUCUN reglage permanent ni global de la machine,
'     et ne necessite aucun droit administrateur.
'   - Verifier que "powershell.exe" (Windows PowerShell, present par defaut
'     sur toutes les versions de Windows depuis Windows 7) est bien accessible
'     depuis le PATH systeme.
'   - Une fenetre console apparait brievement puis disparait : normal au tout
'     premier lancement (compilation Add-Type), sans consequence.
' ============================================================================

Private Const PHOTO_VIEWER_SCRIPT_PATH As String = "C:\Chemin\Vers\PhotoViewer.ps1"

Public Sub LancerPhotoViewer()
    Dim scriptPath As String
    Dim commandLine As String

    scriptPath = PHOTO_VIEWER_SCRIPT_PATH

    If Dir(scriptPath) = vbNullString Then
        MsgBox "Photo Viewer introuvable :" & vbCrLf & scriptPath & vbCrLf & vbCrLf & _
               "Verifiez la constante PHOTO_VIEWER_SCRIPT_PATH dans le module VBA " & _
               "PhotoViewerLauncher.", vbExclamation, "Photo Viewer"
        Exit Sub
    End If

    ' -ExecutionPolicy Bypass : contourne une politique restrictive pour CE lancement
    '   uniquement (aucun changement permanent/global, aucun droit admin requis).
    ' -WindowStyle Hidden + vbHide (parametre Shell) : pas de fenetre console visible.
    ' -NoProfile : demarrage plus rapide, ignore les profils PowerShell de l'utilisateur.
    commandLine = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & scriptPath & """"

    Shell commandLine, vbHide
End Sub
