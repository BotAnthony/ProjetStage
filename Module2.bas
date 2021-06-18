Attribute VB_Name = "Module2"
Sub shpSupp_Cliquer()
    ' MsgBox ActiveSheet.Shapes(Application.Caller).TopLeftCell.Row
    Dim intAnswer As Integer
    
    intAnswer = MsgBox("Voulez vraiment supprimer ce bénévole ?", vbOKCancel, "Veuillez confirmez")
    If intAnswer = vbOK Then
        Rows(ActiveSheet.Shapes(Application.Caller).TopLeftCell.Row).Delete
    End If
End Sub
