Attribute VB_Name = "modFinancialEngine"
Option Explicit

Public Function AverageCost(ByVal TotalCost As Double, ByVal Shares As Double) As Double
    If Shares = 0 Then Exit Function
    AverageCost = TotalCost / Shares
End Function

Public Function UnrealizedPnL(ByVal MarketPrice As Double, ByVal AvgCost As Double, ByVal Shares As Double) As Double
    UnrealizedPnL = (MarketPrice - AvgCost) * Shares
End Function

Public Function ReturnPct(ByVal MarketPrice As Double, ByVal AvgCost As Double) As Double
    If AvgCost = 0 Then Exit Function
    ReturnPct = (MarketPrice - AvgCost) / AvgCost
End Function
