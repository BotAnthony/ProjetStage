Attribute VB_Name = "RemplirPrésence"

Public Sub OuvrirFichier()


'    strFile = Application.GetOpenFilename(FileFilter:= _
'        "Excel files (*.xlsm*), *.xlsm*", Title:="Choisir le fichier de la période correspondant au mois sélectionné")
'
    Dim strFile As String
    Dim Myfile As FileDialog
    Set Myfile = Application.FileDialog(msoFileDialogFilePicker)


      With Myfile
    
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xls; *.xlsm", 1
        .Title = "Choisir le fichier de la période correspondant au mois sélectionné"
        .AllowMultiSelect = False
        .Show
        strFile = .SelectedItems(1)
    End With

    
    
    
    Workbooks.Open (strFile)
    Workbooks(Dir(strFile)).Activate
End Sub
Public Sub boucleCopier(sem As Integer, jour As Integer, dateMois As Date)
    Call OuvrirFichier
    Sheets("S" & sem).Select

    Temp = 2 + jour * 2
    
    'selection des premiers jours du mois
    Range(Cells(5, Temp), Cells(94, 17)).Select
    Selection.Copy
    nbColCopiees = Range(Cells(5, Temp), Cells(94, 17)).Columns.Count - 1
    
    nbJoursMois = Day(Application.WorksheetFunction.EoMonth(dateMois, 0))
    
    ' creation nouvelle feuille temp
    Sheets.Add
    ActiveSheet.Name = "FeuilleTemp"
    Range("A1").Select
    ActiveSheet.Paste
    
    sem = sem + 1
    
    ' tant que le nombre copie est inferieur au nb jour du mois
    While nbColCopiees < nbJoursMois * 2
        ' copier
        Sheets("S" & sem).Select
        Range(Cells(5, 4), Cells(94, 17)).Select
        Selection.Copy
        
        ' coler
        Sheets("FeuilleTemp").Select
        Range("A1").Offset(, nbColCopiees + 1).Select
        ActiveSheet.Paste
        
        ' prochaine feuille
        nbColCopiees = nbColCopiees + 14
        sem = sem + 1
    Wend
    
    Call InsertionNomBene
    
    ' suppression des colonnes en trop
    Range(Columns(nbJoursMois * 2 + 2), Columns(nbColCopiees + 2)).Select
    Selection.Delete
    
End Sub
Function sheetExists(sheetToFind As String) As Boolean
    sheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function
Public Sub bouclePresence(dateMois As Date, Optional nbSemPassee As Integer = 0)
    
    ' declaration variable des WB
    Dim wbMoisInd, wbMain, wbSource As Workbook
    Set wbSource = ActiveWorkbook
    
    Set wbMain = Workbooks("Création2.xlsm")
    wbMain.Activate
    mois = Range("K3").Value
    annee = Range("L3").Value
    
    Set wbMoisInd = Workbooks("FEUILLE DE PRESENCE " & mois & " .xlsm")
    
    ' Determine nombre benevole
    wbSource.Activate
    Range("A1").Select
    nbBenevoles = Selection.End(xlDown).Row
    
    jourSem = Weekday(dateMois) - 2
    
    ' boucle sur tous les benevoles
    Dim i As Integer
    Dim nomBene As String
    For i = 1 To nbBenevoles
    
        ' copie nom benevole
        wbSource.Activate
        nomBene = Trim(Cells(i, 1).Value)
        
        'copie donnée du bénévoles
        Range(Cells(i, 2), Cells(i, 70)).Select
        Selection.Copy
        
        'ouverture feuille benevole
        wbMoisInd.Activate

        If sheetExists(nomBene) Then
        
            Sheets(nomBene).Select
        ' si la feuille du benevole n'est pas la -> creation
        Else
            Sheets(".NOUVEAU").Visible = xlSheetVisible
            Sheets(".NOUVEAU").Copy After:=Sheets(Sheets.Count)
            Sheets(".NOUVEAU").Visible = xlSheetVeryHidden
            ' remplissage infos bénévoles
            ActiveSheet.Name = nomBene
            Range("C10").Value = nomBene
            
            wbSource.Activate
            Range(Cells(i, 2), Cells(i, 70)).Select
            Selection.Copy
            
            wbMoisInd.Activate
            
        End If
           
           
        ' colle dans le fichier bénévole
        
        Range("A54").Select
        ActiveSheet.Paste

        ' remplissage info mois et annee
        Range("D7").Value = mois
        Range("G4").Value = annee

        
    
        Dim j, nbJoursMois As Integer
        nbJoursMois = Day(Application.WorksheetFunction.EoMonth(dateMois, 0))
        
        Dim cellInfo, cellJour As Range
        Set cellInfo = Range("A54")
        
        'selection premiere case en fonction du jour
        If nbSemPassee <> 0 Then
            Cells(24, 2 + nbSemPassee).Select
            nbJoursMois = nbJoursMois - (nbSemPassee * 7) + 1
        ElseIf jourSem < 0 Then
            Cells(36, 2).Select
        Else
            Cells(24 + jourSem * 2, 2).Select
        End If
        Set cellJour = ActiveCell
        
        ' boucle pour ranger les infos depuis la ligne 54 vers le tableau (b24:g37)
        For j = 1 To nbJoursMois * 2
            
            'si il y a quelque chose dans la cellule on met un 1
            If Not Trim(cellInfo.Value) = "" Then
                cellJour.Value = "1"
            End If
            
            Set cellInfo = cellInfo.Offset(0, 1)
            If cellJour.Row = 37 Then
                Set cellJour = Cells(24, cellJour.Column + 1)
            Else
                Set cellJour = cellJour.Offset(1, 0)
            End If
            
        Next j
        
        'del ligne 54
        Rows(54).Delete
        
        
    Next i
    
    ' suppression de la feuille temporaire
    wbSource.Activate
    Application.DisplayAlerts = False
    Sheets("FeuilleTemp").Delete
    Application.DisplayAlerts = True
    wbSource.Close savechanges:=False
    wbMoisInd.Activate
    Sheets(1).Select
End Sub

Public Sub InsertionNomBene()
    ' copie des noms bénévoles
    Sheets("S1").Select
    Range("B5").Select
    Range(Cells(5, 2), Cells(Selection.End(xlDown).Row, 2)).Select
    Selection.Copy
    
    ' Insertion des noms bénévoles
    Sheets("FeuilleTemp").Select
    Columns(1).Insert
    Range("A1").Select
    ActiveSheet.Paste
End Sub
