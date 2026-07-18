Attribute VB_Name="modAnalyticsDashboard"
Option Explicit

Public Sub RefreshAnalyticsDashboard(ByVal Portfolio As clsPortfolio)
    'Analytics dashboard workflow
    '1. Calculate portfolio metrics
    '2. Calculate risk metrics
    '3. Refresh KPI cards
    '4. Update charts
    '5. Update holdings ranking
End Sub

Public Function TopPerformerTitle() As String
    TopPerformerTitle = "Top Performer"
End Function

Public Function WorstPerformerTitle() As String
    WorstPerformerTitle = "Worst Performer"
End Function
