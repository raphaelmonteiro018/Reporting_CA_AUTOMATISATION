Sub GenererBrouillonReporting()
    Dim outApp As Object
    Dim outMail As Object
    Dim strSignature As String
    Dim strDate As String
    Dim strBody As String
    Dim strFilePath As String

    ' Formatage de la date au format 01/01/2000
    strDate = Format(Date, "dd/mm/yyyy")
    
    ' CHEMIN DU FICHIER :
    strFilePath = "\\CHEMIN_PDF"

    On Error Resume Next
    Set outApp = GetObject(, "Outlook.Application")
    If outApp Is Nothing Then
        Set outApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0

    Set outMail = outApp.CreateItem(0)

    ' Affichage pour charger la signature
    outMail.Display 
    strSignature = outMail.HTMLBody

    ' Corps du mail pro optimisé pour le gain de temps
    strBody = "<span style='font-family:Calibri,sans-serif;font-size:11pt;'>" & _
              "Bonjour,<br><br>" & _
              "Veuillez trouver ci-joint le <b>reporting quotidien du chiffre d'affaires</b> pour la région, " & _
              "mis à jour au <b>" & strDate & "</b>.<br><br>" & _
              "Cordialement,</span>"

    With outMail
        .To = "destinataire.anonyme@entreprise.com"
        .CC = "copie.anonyme@entreprise.com"
        .Subject = "Reporting CA Région au " & strDate
        
        ' AJOUT DE LA PIÈCE JOINTE
        If strFilePath <> "" Then .Attachments.Add strFilePath
        
        .HTMLBody = strBody & strSignature 
    End With

    ' Nettoyage des objets
    Set outMail = Nothing
    Set outApp = Nothing
End Sub
