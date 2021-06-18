Attribute VB_Name = "Module1"

Public Sub test()
    Set mydocument = Worksheets(1)
    With mydocument.Shapes("truc1").Fill
    .ForeColor.RGB = RGB(128, 0, 0)
    .BackColor.RGB = RGB(170, 170, 170)
    .TwoColorGradient msoGradientHorizontal, 1
End With
End Sub


Public Sub modifBtn()
    Dim BtnSupp As Object
    Set BtnSupp = ActiveSheet.Shapes(1)
    
    With BtnSupp
        .Width = 70
        .Height = 24
    End With

    
End Sub

Public Sub duplicateBtn(nbBene As Integer)
    Dim BtnSupp, newBtn As Object
    Set BtnSupp = ActiveSheet.Shapes(1)
    
    BtnSupp.Copy
    Cells(nbBene + 1, 4).Select
    ActiveSheet.Paste
    Selection.Name = "Bouton" & CStr(nbBene)
    Set newBtn = ActiveSheet.Shapes("Bouton" & CStr(nbBene))
    
End Sub

Public Sub createBtn()
    Dim i As Integer
    i = 5
    ' For i = 1 To 5
        Call duplicateBtn(i)
        Call moveBtn(i)
    ' Next i
End Sub


Public Sub moveBtn(nbBene)
    Dim bBtn As Object
    Set bBtn = ActiveSheet.Shapes("Bouton" & CStr(nbBene))
    
    With bBtn
        ' centre bouton
        .Left = .Left + ((.TopLeftCell.Width - .Width) / 2)
        .Top = .Top + ((.TopLeftCell.Height - .Height) / 2)
    End With
    
    
End Sub
