Attribute VB_Name = "Module1"

Public Sub OuvrirFichier()

    Dim strFile As String
    Dim Myfile As FileDialog
    Set Myfile = Application.FileDialog(msoFileDialogFilePicker)

      With Myfile
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xls; *.xlsm", 1
        .Title = "Choisir le fichier du mois voulu"
        .AllowMultiSelect = False
        .Show
        strFile = .SelectedItems(1)
    End With
    
    Workbooks.Open (strFile)
    'Workbooks(Dir(strFile)).Activate
End Sub
Public Function getNbBeneVenus(wb As Workbook) As Integer
    wb.Activate
    Sheets(1).Select
    
    Dim nbTotalBene, nbBene As Integer
    nbTotalBene = getnbBene
    nbBene = 0
       
    For i = 1 To nbTotalBene
        If Cells(i + 1, 5) <> 0 Then
            nbBene = nbBene + 1
        End If
    Next i
    getNbBeneVenus = nbBene
End Function
Public Function getDemiJournees(wb As Workbook) As Integer
    wb.Activate
    Sheets(1).Select
    
    getDemiJournees = WorksheetFunction.Sum(Range("tabbenevoles[Aller/retour]").Value)
End Function
Function getnbBene() As Integer
    ' Sheets(1).Select
    getnbBene = Sheets(1).Range("A2").End(xlDown).Row - 1
    ' getnbBene Range("tabBenevoles[Nom]").Rows.Count
End Function
