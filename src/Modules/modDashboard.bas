Attribute VB_Name="modDashboard"
Option Explicit

Public Sub RefreshDashboard(ByVal Portfolio As clsPortfolio)
    'Dashboard refresh workflow
    '1. Update market values
    '2. Update P/L
    '3. Update allocation charts
    '4. Refresh timestamp
End Sub

Public Sub UpdateLastRefresh(ByVal Target As Range)
    Target.Value = Now
End Sub
