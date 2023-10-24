Sub EnvoyerDonneesVersAccess()
    Dim db As Object
    Dim rs As Object
    Dim appAccess As Object
    Dim strSQL As String
    Dim ws As Worksheet
    Dim ligne_max As Integer
    Dim cheminAccess As String
    
    Range("B1").Select
    Selection.End(xlDown).Select
    ligne_max = ActiveCell.Row
        
    'Accès à la BDD
    cheminAccess = "I:\6_OpenClass Rooms\99_Référent technique\2022-10-15 - Cours VBA\analyser en vba\P3C1 - AVDJ\Base de données BRESCIA.accdb"
    
    'Nom de la feuille contenant les données
    Set ws = ThisWorkbook.Sheets("Data")
    
    'Ouverture de la base de données
    Set appAccess = CreateObject("Access.Application")
    appAccess.OpenCurrentDatabase cheminAccess
    Set db = appAccess.CurrentDb
    'nom de la table dans access
    Set rs = db.OpenRecordset("Data")
    
    'Boucle qui permet de remplir les lignes une à une en choisissant les données
    For i = 2 To ligne_max
        'Ajouter un nouvel enregistrement
        rs.AddNew
        rs.Fields("Date").Value = ws.Cells(i, 1).Value ' Remplacez "Colonne1" par le nom de votre colonne dans la table Access
        rs.Fields("Référence").Value = ws.Cells(i, 2).Value ' Remplacez "Colonne2" par le nom de votre colonne dans la table Access
        rs.Fields("Poids").Value = ws.Cells(i, 3).Value ' Remplacez "Colonne3" par le nom de votre colonne dans la table Access
        rs.Fields("Désignation").Value = ws.Cells(i, 4).Value ' Remplacez "Colonne3" par le nom de votre colonne dans la table Access
        rs.Fields("GLUTEN FREE").Value = ws.Cells(i, 5).Value ' Remplacez "Colonne3" par le nom de votre colonne dans la table Access
        rs.Fields("BIO").Value = ws.Cells(i, 6).Value ' Remplacez "Colonne3" par le nom de votre colonne dans la table Access
        rs.Fields("Code Marque").Value = ws.Cells(i, 7).Value ' Remplacez "Colonne3" par le nom de votre colonne dans la table Access
        rs.Fields("Marque").Value = ws.Cells(i, 8).Value ' Remplacez "Colonne3" par le nom de votre colonne dans la table Access
        rs.Fields("TVA").Value = ws.Cells(i, 9).Value ' Remplacez "Colonne3" par le nom de votre colonne dans la table Access
        rs.Fields("Prix de vente").Value = ws.Cells(i, 10).Value ' Remplacez "Colonne3" par le nom de votre colonne dans la table Access
        rs.Fields("Quantités mois").Value = ws.Cells(i, 11).Value ' Remplacez "Colonne3" par le nom de votre colonne dans la table Access
        rs.Fields("CA").Value = ws.Cells(i, 12).Value ' Remplacez "Colonne3" par le nom de votre colonne dans la table Access
        rs.Fields("Fournisseur").Value = ws.Cells(i, 13).Value ' Remplacez "Colonne3" par le nom de votre colonne dans la table Access
        'Mise à jour de la table
        rs.Update
    Next i
    
    'Fermeture de la BDD
    appAccess.DoCmd.Quit acQuitSaveAll
    
    'Libérer les objets
    Set rs = Nothing
    Set db = Nothing
    Set appAccess = Nothing
    
End Sub
