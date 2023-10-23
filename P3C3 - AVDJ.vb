'A mettre dans Thisworkbook
Private Sub Workbook_AfterSave(ByVal Success As Boolean)

Dim outlook As Object
Dim date_j As String
Dim message As String
Dim total_CA_TTC As Double
Dim total_CA_HT As Double
Dim total_TVA As Double
Dim total_QTY As Long
Dim Prix_moyen As Double
Dim Date_max As Date

Date_max = Application.WorksheetFunction.Max(Range("A2:A1000"))

total_CA_TTC = Application.WorksheetFunction.SumIf(Range("A2:A1000"), Date_max, Range("K2:K1000"))
total_CA_HT = Application.WorksheetFunction.Round(total_CA_TTC / (1 + 0.055), 2)
total_TVA = total_CA_TTC - total_TVA
total_QTY = Application.WorksheetFunction.SumIf(Range("A2:A1000"), Date_max, Range("J2:J1000"))
Prix_moyen = Application.WorksheetFunction.Round(total_CA_TTC / total_QTY, 2)

If MsgBox("Voulez-vous envoyer l'email avec le reporting ?", vbYesNo, "Demande de confirmation") = vbYes Then
    Set outlook = CreateObject("Outlook.Application")
    
    date_j = Replace(Date, "/", "-")
    
    message = "Bonjour," & vbNewLine & vbNewLine & "Vous trouverez ci-joint le reporting en date du : " & date_j & vbNewLine & "Le CA TTC du mois dernier est de : " & total_CA_TTC & "€" & vbNewLine & "Le CA HT du mois dernier est de : " & total_CA_HT & "€" & vbNewLine & "Le prix moyen est de : " & Prix_moyen & "€" & vbNewLine & "les quantités vendus du mois dernier sont de : " & total_QTY & vbNewLine & vbNewLine & "Jeremy - Supplychain Analyst"
    
    'Création et envoie de l'email automatiquement
    With outlook.Createitem(alMailitem)
        .Subject = "Reporting - " & date_j
        .To = "jeremy.ollier@gmail.com"
        .Body = message
        .Attachments.Add ("I:\6_OpenClass Rooms\99_Référent technique\2022-10-15 - Cours VBA\analyser en vba\P1C3 - AVDJ\Reporting - " & date_j & ".xlsb")
        .send
    End With
End If

End Sub







'A mettre dans le module
Sub Listing_fichiers()

Dim Chemin As String
Dim Fichier As String
Dim i As Integer
Dim j As Integer
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Dim fichier_reporting As String
Dim fichier_a_rajouter As String
Dim cell_max As Integer
Dim date_j As String

fichier_reporting = ActiveWorkbook.Name
    
'/////////////////////////////////////////////////////////
'////////// BOUCLE POUR LE LISTING DES FICHIERS \\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

    'initialisation de la variable
    i = 1
    
    'Changement de feuille Excel
    ActiveSheet.Name = "Reporting"
    Sheets("Reporting").Select
    
    'choix du dossier à lister
    Chemin = "I:\6_OpenClass Rooms\99_Référent technique\2022-10-15 - Cours VBA\analyser en vba\P1C3 - AVDJ\Extraction\"
    Fichier = Dir(Chemin)
 
    'Boucle sur les fichiers xls du répertoire.
    Do While Len(Fichier) > 0
        Range("A" & i).Value = Chemin & Fichier
        i = i + 1
        Fichier = Dir()
    Loop
    
'Création du dossier
MkDir ("I:\6_OpenClass Rooms\99_Référent technique\2022-10-15 - Cours VBA\analyser en vba\P1C3 - AVDJ\Extraction\Fichiers_traités\")
    
'/////////////////////////////////////////////
'////////// traitement des fichiers \\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Data"
    
For j = 1 To (i - 1)
    'Changement de feuille Excel
    Sheets("Reporting").Select
    'recuperation du lien du fichier à compiler
    Fichier = Range("A" & j).Value
    'ouverture du fichier
    Workbooks.Open Filename:=Fichier
    'récupération du nom du fichier
    fichier_a_rajouter = ActiveWorkbook.Name
    'changement de fichier excel
    Windows(fichier_a_rajouter).Activate
    'Récupère les entetes des colonnes si c'est le premier copier/coller
    Range("A1").Select
    Selection.End(xlDown).Select
    cell_max = ActiveCell.Row
    
    If j > 1 Then
    Range("A2:" & "L" & cell_max).Select
    Selection.Copy
    Else
    Range("A1:" & "L" & cell_max).Select
    End If
    Selection.Copy

    'retour sur le fichier des données
    Windows(fichier_reporting).Activate
    'Changement de feuille Excel
    Sheets("Data").Select
    'si c'est le premier alors on garde la premiere cellule sinon on va chercher la premiere cellule vide
    If j > 1 Then
    Range("A2").Select
    Selection.End(xlDown).Select
    cell_max = ActiveCell.Row
    Range("A" & (cell_max + 1)).Select
    Else
    Range("A1").Select
    End If
    'collage des données
    ActiveSheet.Paste
    'retour sur le fichier
    Windows(fichier_a_rajouter).Activate
    'fermeture du fichier
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    'déplacement du fichier traité
    FSO.MoveFile Fichier, Chemin & "Fichiers_traités\"
    'Ajoute d'une indicatation dans le fichier pour valider que le traitement est OK
    Sheets("Reporting").Select
    Range("B" & j).Select
    ActiveCell.FormulaR1C1 = "Fichier OK"
Next j

    'Harmonisation des format
    Sheets("Data").Select
    Columns("H:H").Select
    Selection.NumberFormat = "0%"
    Columns("I:I").Select
    Selection.NumberFormat = "0.00"
    Columns("J:J").Select
    Selection.NumberFormat = "#,##0"
    Columns("K:K").Select
    Selection.NumberFormat = "#,##0.00 $"
    
    date_j = Replace(Date, "/", "-")
    'enregistrement
    ActiveWorkbook.SaveAs Filename:= _
    "I:\6_OpenClass Rooms\99_Référent technique\2022-10-15 - Cours VBA\analyser en vba\P1C3 - AVDJ\Reporting - " & date_j & ".xlsb" _
    , FileFormat:=xlExcel12, CreateBackup:=False
        
End Sub


