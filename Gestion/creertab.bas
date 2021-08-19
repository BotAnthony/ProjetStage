Attribute VB_Name = "creertab"
Option Explicit
Global colSupp As Integer

Public Sub infosBene()
    Dim WS_Count, i, km As Integer
    Dim nomPrenom, nom, prenom, adresse, venue As String
    
    colSupp = 6
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    For i = 3 To WS_Count
        'If Sheets(i).Visible <> xlSheetHidden Then

            Sheets(i).Select
            
            nomPrenom = Range("C10").Value
            If nomPrenom <> "" Then
                nom = Sheets(i).name
                prenom = Split(nomPrenom)(UBound(Split(nomPrenom)))
            End If
            adresse = Range("C11").Value
            km = Range("F16").Value
            venue = Range("D38").Value
            Sheets(1).Select
            
            Cells(i - 1, 1).Select
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=nom & "!A1", TextToDisplay:=nom
            Cells(i - 1, 2).Value = prenom
            Cells(i - 1, 3).Value = adresse
            Cells(i - 1, 4).Value = km
            Cells(i - 1, 5).Value = venue
        
        'End If
        
    Next i
    With Columns("A:E")
        .Font.Size = 11.5
        .AutoFit
        .VerticalAlignment = xlVAlignCenter
    End With
 End Sub


Public Sub creationTableau()
    If CheckIfTableExists Then
        Sheets(1).ListObjects("tabBenevoles").Delete
        Call deleteBtn
    End If
    
    Call infosBene
    Call initFeuille
    
    Sheets(1).Select
    
    Range("A1").Value = "Nom"
    Range("B1").Value = "Prenom"
    Range("C1").Value = "Adresse"
    Range("D1").Value = "Km"
    Range("E1").Value = "Aller/retour"
    Cells(1, colSupp).Value = "Supprimer"
    
    Range("A1").Select
    Sheets(1).ListObjects.Add(xlSrcRange, Selection.CurrentRegion, xlYes).name = "tabBenevoles"
    ActiveSheet.ListObjects("tabbenevoles").TableStyle = "TableStyleLight15"
    
    Call createBtn
    
End Sub
Function CheckIfTableExists() As Boolean

'Create variables to hold the worksheet and the table
Dim ws As Worksheet
Dim tbl As ListObject
Dim tblName As String
Dim tblExists As Boolean

tblName = "tabBenevoles"

Set ws = Sheets(1)

    'Loop through each table in worksheet
    For Each tbl In ws.ListObjects

        If tbl.name = tblName Then

            tblExists = True

        End If

    Next tbl


If tblExists = True Then

    CheckIfTableExists = True

Else

    CheckIfTableExists = False

End If

End Function

Public Sub deleteBtn()
    Dim shp As Shape
    Dim sh As Worksheet
    Dim exRegBtn As Object
    
    Set exRegBtn = CreateObject("VBScript.RegExp")
    Set sh = ThisWorkbook.Worksheets(1)
    exRegBtn.Pattern = "^Bouton[0-9]{1,3}"

    'If there is any shape on the sheet
    If sh.Shapes.Count > 0 Then
        'Loop through all the shapes on the sheet
        For Each shp In sh.Shapes

            If exRegBtn.Test(shp.name) Then
                shp.Delete
            End If
        Next shp
    End If
End Sub
