Attribute VB_Name = "MoisTransition"

Public Sub CopieDonneesTransition1(debutMois As Date, finPeriode As Date, nbTransi As Integer)
    
    Call SelectLastSheet
    Dim nbJoursFichier, nbColCopiees, nbJours As Integer
    Dim numSheet, numSheetTotal As Integer
    
    ' stocke num dernier sheet
    numSheet = Right(ActiveSheet.Name, 2)
    numSheetTotal = numSheet
    
    ' nombre de jours à copier dans ce fichier
    nbJoursFichier = finPeriode - debutMois + 1
    nbColCopiees = 0
    nbJours = 0
    
    
    
    ' ajout feuille temp
    Sheets.Add
    ActiveSheet.Name = "FeuilleTemp"
    
    
    ' tant que nb col copiees inferieur au nombre de jour à copier
    While nbColCopiees < nbJoursFichier * 2
        
        ' copie d'un jour complet (2 col)
        Sheets("S" & numSheet).Select
        Range(Cells(5, 17 - nbJours * 2), Cells(94, 16 - nbJours * 2)).Select
        Selection.Copy
        
        ' coller
        Sheets("FeuilleTemp").Select
        Cells(1, nbJoursFichier * 2 - nbColCopiees - 1).Select
        ActiveSheet.Paste
        
        nbColCopiees = nbColCopiees + 2
        nbJours = nbJours + 1
        
        ' si la semaine est finie -> change de page
        If nbColCopiees < nbJoursFichier * 2 And nbJours >= 7 Then
            numSheet = numSheet - 1
            nbJours = 0
        End If
    Wend
    
    Call InsertionNomBene
    ' Debug.Print (numSheetTotal - numSheet +1)
    
    Call bouclePresence(DateValue(debutMois))
    ' Call boucleCopier(1, 1, debutMois)
    
    
    ' If nbTransi = 1 Then
        Call CopieDonneesTransition2(finPeriode + 1, Application.WorksheetFunction.EoMonth(debutMois, 0))
    ' Else
        ' Call boucleCopier(1, 1, debutMois)
    ' End If
    
    Call bouclePresence(debutMois, numSheetTotal - numSheet + 1)
    
End Sub

Public Sub SelectLastSheet()
    Call OuvrirFichier
    Dim compSheet As Integer
    
    Sheets(Sheets.Count).Select
    While InStr(ActiveSheet.Name, "S") = 0
        Sheets(Sheets.Count - compSheet).Select
        compSheet = compSheet + 1
    Wend
    
End Sub



Public Sub CopieDonneesTransition2(debutPeriode As Date, finMois As Date)
    ' message utilisateur ouverture deuxieme fichier
    
    Call OuvrirFichier
      
    
    Dim nbJoursFichier, nbColCopiees, nbJours As Integer
    Dim numSheet, joursAvant As Integer
    
    ' stocke num premier sheet
    numSheet = 1
    
    ' nombre de jours à copier dans ce fichier
    nbJoursFichier = finMois - debutPeriode + 1
    nbColCopiees = 0
    nbJours = 0
    ' joursAvant = Day(debutPeriode)
    
    ' ajout feuille temp
    Sheets.Add
    ActiveSheet.Name = "FeuilleTemp"
    
    ' tant que nb col copiees inferieur au nombre de jour à copier
    While nbColCopiees < nbJoursFichier * 2
        
        ' copie d'un jour complet (2 col)
        Sheets("S" & numSheet).Select
        Range(Cells(5, 4 + nbJours * 2), Cells(94, 5 + nbJours * 2)).Select
        Selection.Copy
        
        ' coller
        Sheets("FeuilleTemp").Select
        ' Cells(1, nbJoursFichier * 2 + nbColCopiees).Select
        Range("A1").Offset(, nbColCopiees + 1).Select
        ActiveSheet.Paste
        
        nbColCopiees = nbColCopiees + 2
        nbJours = nbJours + 1
        
        ' si la semaine est finie -> change de page
        If nbJours >= 7 Then
            numSheet = numSheet + 1
            nbJours = 0
        End If
    Wend
    
    Columns(1).Delete
    Call InsertionNomBene ' sur Feuille temp en 1ere col
    
End Sub
