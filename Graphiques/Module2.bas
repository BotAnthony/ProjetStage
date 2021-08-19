Attribute VB_Name = "Module2"
Sub insertGraph()
Attribute insertGraph.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Call tezsdf(wb)
End Sub

Public Sub tezsdf(wb As Workbook)
    Debug.Print wb.Name
End Sub
Sub Charts_Example1()

  Dim MyChart As Chart
  Set MyChart = Charts.Add

  With MyChart
  .SetSourceData Sheets("Cr�ation").Range("A2").CurrentRegion
  .ChartType = xlColumnClustered
  .HasTitle = True
  .ChartTitle.Text = "Demi-journ�e et nombre de b�n�voles par mois"
  .SeriesCollection(1).AxisGroup = 2
  .SeriesCollection(1).ChartType = xlLine
  '.ApplyDataLabels
  .ApplyLayout (5)
  End With

With MyChart.SeriesCollection(1)
    .Name = .Name
    .Values = .Values
    .XValues = .XValues
End With
With MyChart.SeriesCollection(2)
    .Name = .Name
    .Values = .Values
    .XValues = .XValues
End With

End Sub


Public Sub CreerUnGraphique()
    
    Dim wbGraph, wbMois As Workbook
    Set wbGraph = ActiveWorkbook
    
    'nettoyage donn�es pr�c�dentes
    Range("A2").CurrentRegion.Offset(0, 1).Clear
    
    MsgBox "Ouvrez le fichier du mois voulu"
    Call OuvrirFichier
    Set wbMois = ActiveWorkbook
    
    'remplissage infos premier mois
    wbGraph.Sheets("Cr�ation").Range("B1").Value = Split(wbMois.Name)(3)
    wbGraph.Sheets("Cr�ation").Range("B2").Value = getNbBeneVenus(wbMois)
    wbGraph.Sheets("Cr�ation").Range("B3").Value = getDemiJournees(wbMois)
    
    Dim reponse, ncol As Integer
    
    Do
        'fermutre prec fichier
        wbMois.Close savechanges:=False
        
        'ouverture nouveau
        Call OuvrirFichier
        Set wbMois = ActiveWorkbook
        
        ' determine n� col premier mois
        ncol = wbGraph.Sheets("Cr�ation").Range("A2").CurrentRegion.Columns.Count + 1
        
        'remplis infos
        wbGraph.Sheets("Cr�ation").Cells(1, ncol).Value = Split(wbMois.Name)(3)
        wbGraph.Sheets("Cr�ation").Cells(2, ncol).Value = getNbBeneVenus(wbMois)
        wbGraph.Sheets("Cr�ation").Cells(3, ncol).Value = getDemiJournees(wbMois)
        
        reponse = MsgBox("Voulez-vous ajouter un mois suppl�mentaire ?", vbYesNo)
    Loop While reponse = vbYes
    wbMois.Close savechanges:=False
    Charts_Example1
End Sub


Public Sub infosMois(ByRef wbMois As Workbook, ByRef wbGraph As Workbook)
    Dim ncol As Integer
    ncol = Range("A2").CurrentRegion.Columns.Count + 1
    
    wbGraph.Activate
    Sheets("Cr�ation").Select
    
    wbGraph.Sheets("Cr�ation").Cells(1, ncol).Value = Split(wbMois.Name)(3)
    wbGraph.Sheets("Cr�ation").Cells(1, ncol).Value = getNbBeneVenus(wbMois)
    wbGraph.Sheets("Cr�ation").Cells(1, ncol).Value = getDemiJournees(wbMois)
End Sub

