Sub prevision()

Dim i As Integer
Dim mois_m1 As String

i = 32

'Recherche de la plage de données contenant les mois et les données
Do While Range("C" & i).Value <> "Total général"
i = i + 1
Loop

'On récupère le dernier mois et on l'incrémente d'un mois pour faire la prévision
mois_m1 = DateAdd("m", 1, WorksheetFunction.Max(Sheets("Analyse").Range("C32:C" & (i - 1))))

'Prévision pour BRESCIA
ActiveWorkbook.CreateForecastSheet Timeline:=Sheets("Analyse").Range( _
    "C32:C" & (i - 1)), Values:=Sheets("Analyse").Range("D32:D" & (i - 1)), ForecastEnd:= _
    mois_m1, ConfInt:=0.95, Seasonality:=1, ChartType:= _
    xlForecastChartTypeLine, Aggregation:=xlForecastAggregationAverage, _
    DataCompletion:=xlForecastDataCompletionInterpolate, ShowStatsTable:=False
ActiveSheet.Name = "Prévision BRESCIA"
    
'Prévision pour LECCE
ActiveWorkbook.CreateForecastSheet Timeline:=Sheets("Analyse").Range( _
    "C32:C" & (i - 1)), Values:=Sheets("Analyse").Range("E32:E" & (i - 1)), ForecastEnd:= _
    mois_m1, ConfInt:=0.95, Seasonality:=1, ChartType:= _
    xlForecastChartTypeLine, Aggregation:=xlForecastAggregationAverage, _
    DataCompletion:=xlForecastDataCompletionInterpolate, ShowStatsTable:=False
ActiveSheet.Name = "Prévision LECCE"

End Sub

Sub MAJ_Graphique()
    
ActiveWorkbook.RefreshAll

End Sub


