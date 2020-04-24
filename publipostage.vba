Sub publipostage()
If ActiveDocument.MailMerge.DataSource.Name = "" Then
 MsgBox ("Veuillez ajouter une liste de diffusion")
 Exit Sub
End If

MsgBox ("Publipostage depuis " & ActiveDocument.MailMerge.DataSource.Name)
Dim fusion As MailMerge
Dim x As Integer, nb As Integer
Dim chemin As String, nom As String
Set fusion = ActiveDocument.MailMerge
ActiveDocument.MailMerge.DataSource.ActiveRecord = wdLastRecord

chemin = "C:\export\" 'mettre ici le chemin complet du dossier où stocker les fichiers sans oublier le \ à la fin

If Dir(chemin, vbDirectory) = "" Then
  MkDir (chemin)
End If

nb = ActiveDocument.MailMerge.DataSource.ActiveRecord

For x = 1 To nb - 1
With fusion
    .DataSource.FirstRecord = x + 1
    .DataSource.LastRecord = x + 1
    .Destination = wdSendToNewDocument
    .DataSource.ActiveRecord = x + 1
    nom = .DataSource.DataFields("Nom") 'Remplacer Nom" par le champ à utiliser
    prenom = .DataSource.DataFields("Prénom")
    montant = .DataSource.DataFields("Montant")
    .Execute
End With
If nom <> "" And prenom <> "" And montant <> "" Then
    ActiveDocument.ExportAsFixedFormat OutputFileName:=chemin & "fichier_" & nom & "_" & prenom & ".pdf", ExportFormat:=wdExportFormatPDF, openafterexport:=False
End If

ActiveDocument.Close savechanges:=False

Next
End Sub
