Attribute VB_Name = "CalculDate"

Public Sub TypeMois()
Dim mois As String
Dim annee As Integer

Dim debutMois, finMois, debutHiver, finHiver, debutEte As Date

Workbooks("Création2.xlsm").Activate

mois = Range("K3").Value
annee = Range("L3").Value
debutHiver = Range("D3").Value
finHiver = Range("E3").Value
debutEte = Range("B4").Value

debutMois = DateValue("01/" & mois & "/" & annee)
Range("M3").Value = debutMois
finMois = Range("N3").Value

'date mois :
' si avant debut été
If DateValue(debutMois) < DateValue(debutEte) Then
    MsgBox "1. Erreur le mois choisi est avant la date de début", vbCritical
    
' si après debut été
Else
    ' si avant début hiver
    
    If DateValue(debutMois) < DateValue(debutHiver) Then
        
        ' si dernier jour après début hiver
        If DateValue(finMois) > DateValue(debutHiver) Then
            ' MsgBox ("3. Mois de transition 1")
            Call CopieDonneesTransition1(DateValue(debutMois), DateValue(debutHiver) - 1, 1)
        Else
            ' MsgBox ("2. Mois été normal")
            Call calculSemJourEte
        End If
            
    ' si après début hiver
    Else
    
        ' si avant fin hiver
        If DateValue(debutMois) < DateValue(finHiver) Then
    
            ' dernier jour après fin hiver
            If DateValue(finMois) > DateValue(finHiver) Then
                ' MsgBox ("5. Mois de transition 2")
                Call CopieDonneesTransition1(DateValue(debutMois), DateValue(finHiver), 2)
                
            
            Else
                ' MsgBox ("4. Mois hiver normal")
                Call calculSemJourHiver
            End If
            
        ' après fin hiver
        Else
        MsgBox "6. Erreur les dates ne correspondent pas au mois choisi", vbCritical
        
        End If
        
    End If

End If


End Sub
Function MonthNumber(myMonthName As String)
       
    MonthNumber = Month(DateValue("1 " & myMonthName & " 2020"))
    MonthNumber = Format(MonthNumber, "00")
    
End Function



Public Sub calculSemJourEte()
    Dim debutMois, debutEte As Date
    debutMois = Range("M3").Value
    debutEte = Range("B4").Value
    
    Dim nbJoursDiff, sem, jour  As Integer
    nbJoursDiff = debutMois - debutEte
    
    
        'utilise fonction de excel roundDown
    sem = Application.WorksheetFunction.RoundDown(nbJoursDiff / 7, 0)
    
    ' calcul place du premier jour du mois dans le doc
    jour = nbJoursDiff - (sem * 7) + 1


   boucleCopier sem + 1, jour + 0, DateValue(debutMois)
    
    
    bouclePresence DateValue(debutMois)
End Sub

Public Sub calculSemJourHiver()
    Dim debutMois, debutHiver As Date
    debutMois = Range("M3").Value
    debutHiver = Range("D3").Value
    
    Dim nbJoursDiff, sem, jour  As Integer
    nbJoursDiff = debutMois - debutHiver
    
    
        'utilise fonction de excel roundDown
    sem = Application.WorksheetFunction.RoundDown(nbJoursDiff / 7, 0)
    
    ' calcul place du premier jour du mois dans le doc
    jour = nbJoursDiff - (sem * 7) + 1


   boucleCopier sem + 1, jour + 0, DateValue(debutMois)
    
    
    bouclePresence DateValue(debutMois)
End Sub

Public Sub calcSemJourTransi()
    Dim debutMois, debutEte As Date
    debutMois = Range("M3").Value
    debutEte = Range("B4").Value
    
    Dim nbJoursDiff, sem, jour  As Integer
    nbJoursDiff = -debutEte
    
    
        'utilise fonction de excel roundDown
    sem = Application.WorksheetFunction.RoundDown(nbJoursDiff / 7, 0)
    
    ' calcul place du premier jour du mois dans le doc
    jour = nbJoursDiff - (sem * 7) + 1


   boucleCopier sem + 1, jour + 0, DateValue(debutMois)
    
    
    bouclePresence DateValue(debutMois)
End Sub
