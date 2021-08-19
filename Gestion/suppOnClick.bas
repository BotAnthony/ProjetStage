Attribute VB_Name = "suppOnClick"
Sub shpSupp_Cliquer()
    ' MsgBox ActiveSheet.Shapes(Application.Caller).TopLeftCell.Row
    Dim intAnswer, ligneBtn As Integer
    ligneBtn = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Row
    intAnswer = MsgBox("Voulez vraiment supprimer " & Cells(ligneBtn, 1).Value & " ?", vbOKCancel + vbExclamation, "Veuillez confirmez")
    If intAnswer = vbOK Then
        Application.DisplayAlerts = False
        Sheets(Cells(ligneBtn, 1).Value).Delete
        Application.DisplayAlerts = True
        Rows(ligneBtn).Delete

        
        
    End If
End Sub

Public Sub duplicateBtn(nbBene As Integer, Optional newbene = 1)
    Dim BtnSupp, newBtn As Object
    Set BtnSupp = ActiveSheet.Shapes("shpSupp")
    colSupp = 6
    BtnSupp.Copy
    Cells(nbBene + newbene, colSupp).Select
    ActiveSheet.Paste
    Selection.name = "Bouton" & CStr(nbBene)
    Set newBtn = ActiveSheet.Shapes("Bouton" & CStr(nbBene))
    
End Sub

Public Sub createBtn()
    Dim i As Integer
    i = getnbBene
    For i = 1 To getnbBene
        Call duplicateBtn(i)
        Call moveBtn(i)
    Next i
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
