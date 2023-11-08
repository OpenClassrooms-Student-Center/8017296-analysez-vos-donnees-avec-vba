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

    'Initialisation de la variable
    i = 1
    
    'Changement de feuille Excel
    ActiveSheet.Name = "Reporting"
    Sheets("Reporting").Select
    
    'Choix du dossier à lister
    Chemin = "I:\Extraction\"
    Fichier = Dir(Chemin)
 
    'Boucle sur les fichiers xls du répertoire.
    Do While Len(Fichier) > 0
        Range("A" & i).Value = Chemin & Fichier
        i = i + 1
        Fichier = Dir()
    Loop
    
'Création du dossier
MkDir ("I:\Extraction\Fichiers_traités\")
    
'/////////////////////////////////////////////
'////////// traitement des fichiers \\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Data"
    
For j = 1 To (i - 1)
    'Changement de feuille Excel
    Sheets("Reporting").Select
    'Récupération du lien du fichier à compiler
    Fichier = Range("A" & j).Value
    'Ouverture du fichier
    Workbooks.Open Filename:=Fichier
    'Récupération du nom du fichier
    fichier_a_rajouter = ActiveWorkbook.Name
    'Changement de fichier excel
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

    'Retour sur le fichier des données
    Windows(fichier_reporting).Activate
    'Changement de feuille Excel
    Sheets("Data").Select
    'Si c'est le premier alors on garde la premiere cellule sinon on va chercher la premiere cellule vide
    If j > 1 Then
    Range("A2").Select
    Selection.End(xlDown).Select
    cell_max = ActiveCell.Row
    Range("A" & (cell_max + 1)).Select
    Else
    Range("A1").Select
    End If
    'Collage des données
    ActiveSheet.Paste
    'Retour sur le fichier
    Windows(fichier_a_rajouter).Activate
    'Fermeture du fichier
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    'Déplacement du fichier traité
    FSO.MoveFile Fichier, Chemin & "Fichiers_traités\"
    'Ajoute d'une indicatation dans le fichier pour valider que le traitement est OK
    Sheets("Reporting").Select
    Range("B" & j).Select
    ActiveCell.FormulaR1C1 = "Fichier OK"
Next j

'Harmonisation des formats
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
    "I:\Reporting - " & date_j & ".xlsb" _
    , FileFormat:=xlExcel12, CreateBackup:=False
        
End Sub



