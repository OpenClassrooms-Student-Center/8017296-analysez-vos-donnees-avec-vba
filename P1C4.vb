Private Sub Workbook_Open()

Dim jour As Date
Dim statut As String
Dim etat As String
Dim num_ligne As Integer
Dim outlook As Object
Dim résultat As String
Dim nom_utilisateur As String

nom_utilisateur = Environ("USERNAME")

If nom_utilisteur = "Jeremy" Then

    Set outlook = CreateObject("Outlook.Application")
    
    jour = Date
        
    'Ouverture du fichier
    Workbooks.Open Filename:="d:\C1P3\reporting.xlsb"
        
    'Recherche du jour
    Sheets("Reporting").Select
    Range("D3:D40").Select
    Selection.Find(What:=jour, After:=ActiveCell, LookIn:=xlFormulas _
            , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
    
    'Récupération des paramétrages des cellules
    num_ligne = ActiveCell.Row
    statut = Range("E" & num_ligne).Value
    etat = Range("G" & num_ligne).Value
    
    'Teste si on doit lancer le sous-programme
    If (statut = "Oui" And etat = "") Then
        Liste_fichier
        Range("G" & num_ligne).Select
        ActiveCell.FormulaR1C1 = "Fait"
    End If
    
    'Créer une phrase à envoyer par email
    If Range("G" & num_ligne).Value = "Fait" Then
        Resultat = "la mise à jour est passée correctement en date du : " & jour
    Else
        Resultat = "Problème avec la mise à jour du : " & jour
    End If
    
    'Création et envoi de l'email automatiquement
    With outlook.Createitem(alMailitem)
        .Subject = "Reporting - " & jour
        .To = "jeremy.ollier@gmail.com"
        .Body = Resultat
        .send
    End With
End If

End Sub
