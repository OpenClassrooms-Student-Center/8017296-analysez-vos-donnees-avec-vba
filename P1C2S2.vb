Sub Liste_fichiers()

Dim Chemin As String
Dim Fichier As String
Dim i As Integer
Dim j As Integer
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Dim fichier_reporting As String
Dim fichier_a_rajouter As String
Dim cell_max As Integer
    
fichier_reporting = ActiveWorkbook.Name

'/////////////////////////////////////////////////////////
'////////// BOUCLE POUR LE LISTING DES FICHIERS \\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'Initialisation de la variable
i = 1
    
'Changement de feuille Excel
Sheets("Reporting").Select

'Choix du dossier à lister
Chemin = "D:\Extraction\Données\"
Fichier = Dir(Chemin)

'Boucle sur les fichiers du répertoire.
Do While Len(Fichier) > 0
    Range("A" & i).Value = Chemin & Fichier
    i = i + 1
    Fichier = Dir()
Loop
    
'Création du dossier
MkDir ("D:\Extraction\Données\Fichiers_traités\")
    
'/////////////////////////////////////////////
'////////// traitement des fichiers \\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
For j = 1 To (i - 1)
    'Selection de la feuille
    Sheets("Reporting").Select
    'Récupération du lien du fichier à compiler
    Fichier = Range("A" & j).Value
    'Ouverture du fichier
    Workbooks.Open Filename:=Fichier
    'Récupération du nom du fichier
    fichier_a_rajouter = ActiveWorkbook.Name

    
    'Suppression des colonnes inutiles
    Columns("A:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("X:AF").Select
    Selection.Delete Shift:=xlToLeft
    'Changement du format
    Columns("W:W").Select
    Selection.NumberFormat = "#,##0.0"
    'Suppression des .0 dans le code postal
    Columns("N:N").Select
    Selection.Replace What:=".0", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    'Copie des données
    Range("A2:AU1000").Select
    Selection.Copy
    
    
    'Changement de fichier
    Windows(fichier_reporting).Activate
    'Selection de la feuille
    Sheets("Data").Select
    'Recherche de la dernière cellule vide pour coller nos données
    If j > 1 Then
    Range("A2").Select
    Else
    Range("A1").Select
    End If
    Selection.End(xlDown).Select
    cell_max = ActiveCell.Row
    Range("A" & (cell_max + 1)).Select
    'Collage des données
    ActiveSheet.Paste
    'Retour sur le fichier
    Workbooks(fichier_a_rajouter).Activate
    'Fermeture du fichier
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    'Déplacement du fichier traité
    FSO.MoveFile Fichier, Chemin & "Fichiers_traités\"
    'Ajout d'une indicatation dans le fichier pour valider que le traitement est OK
    Sheets("Reporting").Select
    Range("B" & j).Select
    ActiveCell.FormulaR1C1 = "Fichier OK"
Next j

End Sub
