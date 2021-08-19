Attribute VB_Name = "CréationMois"

Sub OpenWorkbook()

Dim source, path As String

path = ThisWorkbook.path & "\Gestion\"
file = "FEUILLE DE PRESENCE.xlsm"
Debug.Print (path & file)
If Dir(path & file) <> "" Then
    Workbooks.Open (path & file)
    Debug.Print (path & file)
    Workbooks(file).Activate
    Sheets("S11").Select
    Range("D5").CurrentRegion.Select
    Selection.Copy
    
    Workbooks("Création.xlsm").Activate
    Sheets("Feuil2").Select
    Range("A1").Select
    ActiveSheet.Paste
    
    Workbooks(file).Close
Else
     MsgBox "Le fichier n'existe pas"
End If



End Sub

Public Sub AfficherFormulaire()
    UserForm1.Show
    Call newFichePresence
    ' Call RemplirInfos
    Call TypeMois
    
End Sub

Public Sub newFichePresence()
    Dim oFSO As Object
    Dim path, mois As String
    
    mois = Range("K3").Value
    
    path = ThisWorkbook.path & "\Feuilles de Mois\FEUILLE DE PRESENCE " & mois & " .xlsm"
    
    ' si le fichier n'existe pas encore
    If Dir(path) = "" Then
        
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        
        ' copie fichier vierge

        Call oFSO.CopyFile(ThisWorkbook.path & "\Gestion\FEUILLE DE PRESENCE.xlsm", path)
        MsgBox ("Fichier créé   :D")
        
        'ouverture du fichier
        Workbooks.Open (path)
    Else
        ' fin du prog
        MsgBox ("Le fichier existe déjà   (¬_¬ )")
        End
    End If
End Sub

Public Sub RemplirInfos()

    Dim mois, annee As String
    ' inialise les variables avec les données rentrée par l'utilisateur
    mois = Range("K3").Value
    annee = Range("L3").Value
    
    path = ThisWorkbook.path & "\Gestion\FEUILLE DE PRESENCE " & mois & " .xlsm"
    
    Workbooks.Open (path)
    Workbooks("FEUILLE DE PRESENCE " & mois & " .xlsm").Activate
    
    Dim x As Worksheet
    For Each x In Worksheets
        Sheets(x.Name).Select
        Range("D7").Value = mois
        Range("G4").Value = annee
        Range("A1").Select
    Next x
    
End Sub
