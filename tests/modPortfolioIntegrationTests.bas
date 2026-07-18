Attribute VB_Name="modPortfolioIntegrationTests"
Option Explicit

Public Sub TestUpdateHolding()
    'Arrange
    Dim h As New clsHolding
    Dim q As New clsQuote
    h.Symbol="TSM"
    q.Symbol="TSM"
    q.Price=250
    UpdateHoldingFromQuote h,q
    Debug.Assert h.MarketValue>=0
End Sub
