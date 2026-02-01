Option Explicit

' ========================================================================================================================
' DESCRIPTION DU CODE D'IMPORT DES DONNEES
' ========================================================================================================================
' Ce code importe automatiquement les plages des fichiers Excel vers leurs feuilles de destination respectives.
'
' Caractéristiques :
' - Optimisation maximale des performances (rafraichissement écran, événements et calculs désactivés lors de l'exécution)
' - Fonction d’import générique et factorisée
' - Facilité de maintenance grâce aux annotations dans le code et à la factorisation
' ========================================================================================================================
Public Sub Import_CA_Region1()

    Dim sourceFoldersCA() As String
    Dim templateRangesCA() As String
    Dim folderAR09 As String, folderVISIO As String, folderCORRESP As String
    Dim rangeAR09 As String, rangeVISIO As String, rangeCORRESP As String
    Dim i As Long

    ' ================= OPTIMISATION DES PERFORMANCES =================
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.StatusBar = "Démarrage de l'importation des données..."
    ' ================================================================

    ' ===================== DOSSIERS SOURCES =========================
    ReDim sourceFoldersCA(1)
    sourceFoldersCA(0) = "\\CHEMIN_SOURCE_ANONYME1"
    sourceFoldersCA(1) = "\\CHEMIN_SOURCE_ANONYME2"

    folderAR09 = "\\CHEMIN_SOURCE_ANONYME3"
    folderVISIO = "\\CHEMIN_SOURCE_ANONYME4"
    folderCORRESP = "\\CHEMIN_SOURCE_ANONYME5"

    ' ===================== PLAGES MODÈLES ===========================
    ' Ces plages servent de référence pour détecter :
    ' - la ligne de départ
    ' - les colonnes à importer
    ReDim templateRangesCA(1)
    templateRangesCA(0) = "A9:Z9"
    templateRangesCA(1) = "A9:Z9"

    rangeAR09 = "A3:X3"
    rangeVISIO = "A5:BK5"
    rangeCORRESP = "A3:D3"

    ' ================================================================
    ' ===================== IMPORT DES DONNÉES =======================
    ' ================================================================

    ' --- CA : deux dossiers, même feuille cible ---
    For i = LBound(sourceFoldersCA) To UBound(sourceFoldersCA)
        Application.StatusBar = "Import CA (" & i + 1 & "/" & UBound(sourceFoldersCA) + 1 & ")..."
        ImportLatestFileIntoSheet sourceFoldersCA(i), templateRangesCA(i), "CA", True
    Next i

    ' --- AR09 ---
    Application.StatusBar = "Import AR09..."
    ImportLatestFileIntoSheet folderAR09, rangeAR09, "AR09", False

    ' --- VISIO ---
    Application.StatusBar = "Import VISIO..."
    ImportLatestFileIntoSheet folderVISIO, rangeVISIO, "VISIO", False

    ' --- CORRESPONDANCE ---
    Application.StatusBar = "Import Correspondance..."
    ImportLatestFileIntoSheet folderCORRESP, rangeCORRESP, "Corresp bureau de l'expert", False

Cleanup:
    ' ================= RESTAURATION DE L’ENVIRONNEMENT ===============
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.StatusBar = "? Toutes les bases ont été importées avec succès."

End Sub


' =============================================================================
' FONCTION D’IMPORT
' =============================================================================
' Ouvre le fichier Excel le plus récent d’un dossier donné et copie les données vers la feuille cible.
'
' Paramètres :
' - sourceFolder  : Dossier contenant les fichiers Excel
' - templateRange : Plage modèle servant à identifier colonnes et ligne de départ
' - targetSheet   : Nom de la feuille de destination
' - appendData    : Si True, ajoute les données à la suite de l’existant
' =============================================================================
Public Sub ImportLatestFileIntoSheet( _
    ByVal sourceFolder As String, _
    ByVal templateRange As String, _
    ByVal targetSheet As String, _
    Optional ByVal appendData As Boolean = False)

    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim latestFile As String

    Dim startCell As String, endCell As String
    Dim startCol As String, endCol As String
    Dim startRow As Long
    Dim lastSourceRow As Long
    Dim targetRow As Long

    Set wsTarget = ThisWorkbook.Sheets(targetSheet)

    ' Récupération du fichier le plus récent
    latestFile = GetLatestExcelFile(sourceFolder)
    If latestFile = "" Then Exit Sub

    ' Ouverture du fichier source en lecture seule (performance)
    Set wbSource = Workbooks.Open(latestFile, ReadOnly:=True)
    Set wsSource = wbSource.Sheets(1)

    ' Analyse de la plage modèle
    startCell = Split(templateRange, ":")(0)
    endCell = Split(templateRange, ":")(1)

    startRow = wsSource.Range(startCell).Row
    startCol = Replace(startCell, startRow, "")
    endCol = Replace(endCell, startRow, "")

    ' Détection de la dernière ligne source à partir de la première colonne
    lastSourceRow = wsSource.Cells(wsSource.Rows.Count, startCol).End(xlUp).Row

    ' Détermination de la ligne de collage dans la feuille cible
    If appendData Then
        targetRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row + 1
    Else
        targetRow = 2
    End If

    ' Collage rapide : valeurs + formats numériques (dates conservées)
    wsSource.Range(startCol & startRow & ":" & endCol & lastSourceRow).Copy
    wsTarget.Range("A" & targetRow).PasteSpecial xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False

    wbSource.Close SaveChanges:=False

End Sub


' =============================================================================
' FONCTION DE RECUPERATION DES FICHIERS SOURCE
' =============================================================================
Public Function GetLatestExcelFile(ByVal folderPath As String) As String

    Dim fileName As String
    Dim latestFile As String
    Dim latestDate As Date

    fileName = Dir(folderPath & "\*.xlsx")
    latestDate = 0

    Do While fileName <> ""
        If FileDateTime(folderPath & "\" & fileName) > latestDate Then
            latestDate = FileDateTime(folderPath & "\" & fileName)
            latestFile = folderPath & "\" & fileName
        End If
        fileName = Dir
    Loop

    GetLatestExcelFile = latestFile

End Function
