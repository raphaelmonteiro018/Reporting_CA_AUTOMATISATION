Option Explicit

Sub Nettoyage_Région1()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim sheetsInfo As Variant
    Dim i As Long

    ' =====================================================================
    ' OPTIMISATION DES PERFORMANCES
    ' =====================================================================
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' =====================================================================
    ' PARAMÉTRAGE DES FEUILLES À NETTOYER
    ' =====================================================================
    ' Structure du tableau :
    '   - Nom de la feuille
    '   - Dernière colonne à vider
    '
    ' Le nettoyage s’effectue systématiquement à partir de la ligne 2
    ' afin de préserver les en-têtes.
    sheetsInfo = Array( _
        Array("AR09", "X"), _
        Array("VISIO", "BK"), _
        Array("Corresp bureau de l'expert", "D"), _
        Array("CA", "Z") _
    )

    ' =====================================================================
    ' BOUCLE DE NETTOYAGE DES FEUILLES
    ' =====================================================================
    For i = LBound(sheetsInfo) To UBound(sheetsInfo)

        ' Tentative de récupération de la feuille
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetsInfo(i)(0))
        On Error GoTo 0

        ' Si la feuille n’existe pas, on passe à la suivante
        If ws Is Nothing Then GoTo NextSheet

        ' Détection de la dernière ligne utilisée sur la colonne A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        ' Nettoyage uniquement si des données existent sous l’en-tête
        If lastRow >= 2 Then
            ' Nettoyage rapide :
            ' - suppression des valeurs uniquement
            ' - formats, largeurs et structures conservés
            ws.Range("A2:" & sheetsInfo(i)(1) & lastRow).Value = ""
        End If

NextSheet:
        ' Libération explicite de l’objet feuille avant itération suivante
        Set ws = Nothing
    Next i

    ' =====================================================================
    ' RESTAURATION DES PARAMÈTRES EXCEL
    ' =====================================================================
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' Message de confirmation utilisateur
    MsgBox "Nettoyage terminé pour IDFC !", vbInformation

End Sub
