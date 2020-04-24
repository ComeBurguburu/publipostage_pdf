Sub publipostage()
Dim fusion As MailMerge
Dim x As Integer, nb As Integer
Dim chemin As String, nom As String
Set fusion = ActiveDocument.MailMerge
chemin = "D:\Mes documents\" 'mettre ici le chemin complet du dossier où stocker les fichiers sans oublier le \ à la fin
nb = fusion.DataSource.RecordCount
For x = 0 To nb - 1
With fusion
    .DataSource.FirstRecord = x + 1
    .DataSource.LastRecord = x + 1
    .Destination = wdSendToNewDocument
    .DataSource.ActiveRecord = x + 1
    nom = .DataSource.DataFields("Nom") 'Remplacer Nom" par le champ à utiliser
    .Execute
End With
ActiveDocument.ExportAsFixedFormat OutputFileName:=chemin & nom & ".pdf", ExportFormat:=wdExportFormatPDF, openafterexport:=False
ActiveDocument.Close savechanges:=False

Next
End Sub