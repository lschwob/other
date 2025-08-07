Sub ExporterFeuillesEnFichiers()
    Dim feuille As Worksheet
    Dim chemin As String

    ' Dossier de sauvegarde (le même que le classeur actuel)
    chemin = ThisWorkbook.Path & "\"

    For Each feuille In ThisWorkbook.Worksheets
        feuille.Copy
        ActiveWorkbook.SaveAs Filename:=chemin & feuille.Name & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        ActiveWorkbook.Close SaveChanges:=False
    Next feuille

    MsgBox "Export terminé !"
End Sub
