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
    Chemin = "I:\6_OpenClass Rooms\99_Référent technique\2022-10-15 - Cours VBA\analyser en vba\P1C2 - ACDJ\Extraction\"
    Fichier = Dir(Chemin)
 
    'Boucle sur les fichiers xls du répertoire.
    Do While Len(Fichier) > 0
        Range("A" & i).Value = Chemin & Fichier
        i = i + 1
        Fichier = Dir()
    Loop
    
'Création du dossier
MkDir ("I:\6_OpenClass Rooms\99_Référent technique\2022-10-15 - Cours VBA\analyser en vba\P1C2 - ACDJ\Extraction\Fichiers_traités\")
    
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
    "I:\6_OpenClass Rooms\99_Référent technique\2022-10-15 - Cours VBA\analyser en vba\P1C2 - ACDJ\Reporting - " & date_j & ".xlsb" _
    , FileFormat:=xlExcel12, CreateBackup:=False
        
End Sub


Sub analyse_donnees()

Dim ligne_non_vide As Integer

'description du fichier
ligne_non_vide = WorksheetFunction.CountA(Sheets("Data").Range("A2:A10000"))
Sheets("Reporting").Range("G1").Value = ligne_non_vide

'Prix de vente
Sheets("Reporting").Range("G5").Value = WorksheetFunction.Min(Sheets("Data").Range("I2:I10000"))
Sheets("Reporting").Range("G6").Value = WorksheetFunction.Max(Sheets("Data").Range("I2:I10000"))
Sheets("Reporting").Range("G7").Value = WorksheetFunction.Average(Sheets("Data").Range("I2:I10000"))
Sheets("Reporting").Range("G8").Value = WorksheetFunction.CountBlank(Sheets("Data").Range("I2:I" & (ligne_non_vide + 1)))

'TVA
Sheets("Reporting").Range("H5").Value = WorksheetFunction.Min(Sheets("Data").Range("H2:H10000"))
Sheets("Reporting").Range("H6").Value = WorksheetFunction.Max(Sheets("Data").Range("H2:H10000"))
Sheets("Reporting").Range("H8").Value = WorksheetFunction.CountBlank(Sheets("Data").Range("H2:H" & (ligne_non_vide + 1)))

'Quantités
Sheets("Reporting").Range("I5").Value = WorksheetFunction.Min(Sheets("Data").Range("J2:J10000"))
Sheets("Reporting").Range("I6").Value = WorksheetFunction.Max(Sheets("Data").Range("J2:J10000"))
Sheets("Reporting").Range("I7").Value = WorksheetFunction.Average(Sheets("Data").Range("J2:J10000"))
Sheets("Reporting").Range("I8").Value = WorksheetFunction.CountBlank(Sheets("Data").Range("J2:J" & (ligne_non_vide + 1)))

'Poids
Sheets("Reporting").Range("J5").Value = WorksheetFunction.Min(Sheets("Data").Range("C2:C10000"))
Sheets("Reporting").Range("J6").Value = WorksheetFunction.Max(Sheets("Data").Range("C2:C10000"))
Sheets("Reporting").Range("J7").Value = WorksheetFunction.Average(Sheets("Data").Range("C2:C10000"))
Sheets("Reporting").Range("J8").Value = WorksheetFunction.CountBlank(Sheets("Data").Range("C2:C" & (ligne_non_vide + 1)))

'CA
Sheets("Reporting").Range("K5").Value = WorksheetFunction.Min(Sheets("Data").Range("K2:K10000"))
Sheets("Reporting").Range("K6").Value = WorksheetFunction.Max(Sheets("Data").Range("K2:K10000"))
Sheets("Reporting").Range("K7").Value = WorksheetFunction.Average(Sheets("Data").Range("K2:K10000"))
Sheets("Reporting").Range("K8").Value = WorksheetFunction.CountBlank(Sheets("Data").Range("K2:K" & (ligne_non_vide + 1)))


'Référence
Sheets("Reporting").Range("G14").Value = WorksheetFunction.CountBlank(Sheets("Data").Range("A2:A" & (ligne_non_vide + 1)))
SP_valeurunique ("B2:B" & (ligne_non_vide + 1)), "G15"


'glutenfree
Sheets("Reporting").Range("H14").Value = WorksheetFunction.CountBlank(Sheets("Data").Range("E2:E" & (ligne_non_vide + 1)))
SP_valeurunique ("E2:E" & (ligne_non_vide + 1)), "H15"

'Bio
Sheets("Reporting").Range("I14").Value = WorksheetFunction.CountBlank(Sheets("Data").Range("F2:F" & (ligne_non_vide + 1)))
SP_valeurunique ("F2:F" & (ligne_non_vide + 1)), "I15"

'code marque
Sheets("Reporting").Range("J14").Value = WorksheetFunction.CountBlank(Sheets("Data").Range("G2:G" & (ligne_non_vide + 1)))
SP_valeurunique ("G2:G" & (ligne_non_vide + 1)), "J15"

'fournisseur
Sheets("Reporting").Range("K14").Value = WorksheetFunction.CountBlank(Sheets("Data").Range("L2:L" & (ligne_non_vide + 1)))
SP_valeurunique ("L2:L" & (ligne_non_vide + 1)), "K15"

End Sub

Sub SP_valeurunique(ByVal zone_analyse As String, ByVal valeur_unique As String)

Dim C As Range
Dim Dico As Object

Set Dico = CreateObject("Scripting.Dictionary")

For Each C In Sheets("Data").Range(zone_analyse).SpecialCells(xlCellTypeVisible)
    If Not Dico.exists(C.Value) Then Dico.Add C.Value, C.Value
Next C

Sheets("Reporting").Range(valeur_unique).Resize(Dico.Count) = Application.Transpose(Dico.items)

End Sub


