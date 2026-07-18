Attribute VB_Name="modPortfolioUpdate"
Option Explicit

Public Sub UpdateHoldingFromQuote(ByVal H As clsHolding, ByVal Q As clsQuote)
    If UCase$(H.Symbol) <> UCase$(Q.Symbol) Then
        Err.Raise 8201,, "Symbol mismatch."
    End If
    H.MarketPrice = Q.Price
End Sub

Public Sub RefreshPortfolio(ByVal P As clsPortfolio, ByVal Provider As clsMarketDataProvider)
    ' Example integration flow:
    ' 1. Request quote from Provider
    ' 2. Parse JSON to clsQuote
    ' 3. Update each holding
    ' 4. Refresh dashboard
End Sub
