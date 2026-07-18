Attribute VB_Name = "modIntegration"
Option Explicit

Public Sub InitializePortfolioManager()
    ' Punto único de inicialización
    Call LoadPortfolio
    Call LoadCurrencies
    Call LoadDividends
    Call RefreshDashboard
End Sub

Public Sub LoadPortfolio()
    ' Integración clsPortfolio
End Sub

Public Sub LoadCurrencies()
    ' Integración clsCurrency
End Sub

Public Sub LoadDividends()
    ' Integración clsDividendEngine
End Sub

Public Function RunHealthCheck() As Boolean
    RunHealthCheck = True
End Function
