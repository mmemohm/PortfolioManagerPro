Attribute VB_Name="modApplication"
Option Explicit

Public Sub InitializeApplication()
    Call LoadConfiguration
    Call RefreshDashboard
    Call RefreshAnalyticsDashboard
End Sub

Public Function ApplicationVersion() As String
    ApplicationVersion = "1.1.0 Beta"
End Function
