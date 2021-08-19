Attribute VB_Name = "initTabBtn"
Public Sub initFeuille()
    Dim largeurBtn, hauteurBtn As Integer
    largeurBtn = 70
    hauteurBtn = 22
    
    
    columnWidthPoints (largeurBtn + 10)
    Call rowHeightPoints(hauteurBtn + 5, getnbBene + 1)
    
    ActiveSheet.Shapes("shpSupp").Width = largeurBtn
    ActiveSheet.Shapes("shpSupp").Height = hauteurBtn
    
End Sub
Function getnbBene() As Integer
    ' Sheets(1).Select
    getnbBene = Sheets(1).Range("A2").End(xlDown).Row - 1
    ' getnbBene Range("tabBenevoles[Nom]").Rows.Count
End Function
Sub columnWidthPoints(Optional largeur As Integer = 80)
 
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-column-width/
 
    Dim iCounter As Long
    
    ' selection cellule de la colonne voulue
    With ActiveSheet.Cells(2, colSupp)
        ' faire une loop 3 pour s'approcher
        ' au plus possible de la taille voulue
        For iCounter = 1 To 3
            .ColumnWidth = largeur * (.ColumnWidth / .Width)
        Next iCounter
    End With
 
End Sub

Sub rowHeightPoints(Optional longueur As Integer = 30, Optional nbBene As Integer = 100)
 
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/excel-vba-column-width/
 
    Dim iCounter As Long
    For i = 2 To nbBene
        With ActiveSheet.Cells(i, colSupp)
    
            For iCounter = 1 To 3
                .RowHeight = longueur * (.RowHeight / .Height)
            Next iCounter
            
        End With
    Next i
 
End Sub


