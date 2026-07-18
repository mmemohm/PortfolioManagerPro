Attribute VB_Name="modAnalyticsDashboardTests"
Option Explicit

Public Sub TestTitles()
    Debug.Assert TopPerformerTitle()="Top Performer"
    Debug.Assert WorstPerformerTitle()="Worst Performer"
End Sub
