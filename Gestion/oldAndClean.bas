Attribute VB_Name = "oldAndClean"
Sub RemoveInfo()
Attribute RemoveInfo.VB_ProcData.VB_Invoke_Func = " \n14"
'
' RemoveInfo Macro
'

'
Dim x As Worksheet
For Each x In Worksheets
    Sheets(x.name).Select
    Range("D7").Select
    Selection.ClearContents
    Range("B21:F21").Select
    Selection.ClearContents
    Range("G4") = Year(Date)
    Range("B24:G37").Select
    Selection.ClearContents
    
Next x

    
    
End Sub
 Sub WorksheetLoop()

    Dim WS_Count As Integer
    Dim i As Integer
    ' Set WS_Count equal to the number of worksheets in the active
    ' workbook.
    WS_Count = ActiveWorkbook.Worksheets.Count
    ' Begin the loop.
    
    For i = 1 To WS_Count
        ' Insert your code here.
        ' The following line shows how to reference a sheet within
        ' the loop by displaying the worksheet name in a dialog box.
        ActiveWorkbook.Worksheets(i).Select
        Range("G22").Value = "semaine 6"
        Range("D38").Formula = "=SUM(B24:G37)"
    Next i
 End Sub

Sub TabsAscending()
 
For i = 1 To Application.Sheets.Count
    For j = 1 To Application.Sheets.Count - 1
        If UCase$(Application.Sheets(j).name) > UCase$(Application.Sheets(j + 1).name) Then
            Sheets(j).Move after:=Sheets(j + 1)
        End If
    Next
Next
MsgBox "Les feuilles ont été triées de A à Z"
Sheets(1).Select
End Sub
