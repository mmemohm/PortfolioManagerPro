Attribute VB_Name="modIntegrationTests"
Option Explicit

Public Sub TestHealthCheck()
    Debug.Assert RunHealthCheck = True
End Sub

Public Sub TestInitialize()
    InitializePortfolioManager
End Sub
