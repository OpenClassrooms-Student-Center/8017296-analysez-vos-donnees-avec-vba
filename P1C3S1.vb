Private Sub Workbook_Open()

Dim jour As Date
Dim statut As String
Dim etat As String
Dim num_ligne As Integer

jour = Date
    
'ouverture du fichier
'Workbooks.Open Filename:="D:\C1P3\reporting.xlsb"
    
'recherche du jour
Sheets("Reporting").Select
Range("D3:D40").Select
Selection.Find(What:=jour, After:=ActiveCell, LookIn:=xlFormulas _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate

'récupération des parametrages des cellules
num_ligne = ActiveCell.Row
statut = Range("E" & num_ligne).Value
etat = Range("G" & num_ligne).Value

'test de lancement du sous programme
If (statut = "Oui" And etat = "") Then
    Liste_fichiers
    Range("G" & num_ligne).Select
    ActiveCell.FormulaR1C1 = "Fait"
End If

End Sub
