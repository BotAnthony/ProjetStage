Attribute VB_Name = "Module1"
Option Explicit
Global nomBene As String
Global prenomBene As String
Global adresseBene As String
Global kmBene As String




Public Sub testFormulaire()
    newbene.Show
    
    Dim rowNewBene As Integer
    
    ' création par copy de nouvelle feuille Bene
    Sheets(".NOUVEAU").Visible = xlSheetVisible
    Sheets(".NOUVEAU").Copy after:=Sheets(Sheets.Count)
    ActiveSheet.name = UCase(nomBene)
    
    ' remplissage donnees sur feuille bene
    Range("C11").Value = adresseBene
    Range("F16").Value = kmBene
    Range("C10").Value = UCase(nomBene) & " " & prenomBene
    
    ' insertion dans le tableau
    ' Insertion nouvelle ligne
    rowNewBene = Sheets(1).Range("A2").End(xlDown).Row
    Sheets(1).Range("A2").End(xlDown).Rows().EntireRow.Insert
    
    'Insertion des infos dans tableau
    Sheets(1).Select
    Cells(rowNewBene, 1).Select
    ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=nomBene & "!A1", TextToDisplay:=UCase(nomBene)
    Cells(rowNewBene, 2).Value = prenomBene
    Cells(rowNewBene, 3).Value = adresseBene
    Cells(rowNewBene, 4).Value = kmBene
    
    ' Insertion bouton
    Call duplicateBtn(rowNewBene, 0)
    Call moveBtn(rowNewBene)
    
    ' rangement alphabetique
    Call RangerAlpha
    
    Sheets(".NOUVEAU").Visible = xlSheetVeryHidden
    Call creationTableau
End Sub
Sub RangerAlpha()
'
' RangerAlpha Macro
'

'
Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    
    Set ws = ActiveSheet
    Set tbl = ws.ListObjects("tabBenevoles")
    Set rng = Range("tabbenevoles[Nom]")
    
    With tbl.Sort
       .SortFields.Clear
       .SortFields.Add Key:=rng, SortOn:=xlSortOnValues, Order:=xlAscending
       .Header = xlYes
       .Apply
    End With

'    ActiveWorkbook.Worksheets(".Gestion").ListObjects("tabBenevoles").Sort. _
'        SortFields.Clear
'    ActiveWorkbook.Worksheets(".Gestion").ListObjects("tabBenevoles").Sort. _
'        SortFields.Add Key:=Range("tabBenevoles[[#All],[Nom]]"), SortOn:= _
'        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'    With ActiveWorkbook.Worksheets(".Gestion").ListObjects("tabBenevoles").Sort
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
End Sub


Public Sub testalpha()
Debug.Print Range("tabBenevoles[Nom]").Rows.Count
End Sub
