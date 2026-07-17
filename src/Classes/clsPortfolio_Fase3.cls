VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsPortfolio"

Option Explicit

'Financial aggregation methods for clsPortfolio.
'Requires each clsHolding to expose:
' MarketValue, CostBasis, UnrealizedGain,
' RealizedGain, DividendsReceived

Private mHoldings As Object

Private Sub Class_Initialize()
    Set mHoldings = CreateObject("Scripting.Dictionary")
End Sub

Public Function MarketValue() As Double
    Dim k
    For Each k In mHoldings.Keys
        MarketValue = MarketValue + CDbl(mHoldings(k).MarketValue)
    Next
End Function

Public Function CostBasis() As Double
    Dim k
    For Each k In mHoldings.Keys
        CostBasis = CostBasis + CDbl(mHoldings(k).CostBasis)
    Next
End Function

Public Function UnrealizedGain() As Double
    Dim k
    For Each k In mHoldings.Keys
        UnrealizedGain = UnrealizedGain + CDbl(mHoldings(k).UnrealizedGain)
    Next
End Function

Public Function RealizedGain() As Double
    Dim k
    For Each k In mHoldings.Keys
        RealizedGain = RealizedGain + CDbl(mHoldings(k).RealizedGain)
    Next
End Function

Public Function TotalDividends() As Double
    Dim k
    For Each k In mHoldings.Keys
        TotalDividends = TotalDividends + CDbl(mHoldings(k).DividendsReceived)
    Next
End Function

Public Function TotalProfit() As Double
    TotalProfit = UnrealizedGain + RealizedGain + TotalDividends
End Function

Public Function ROI() As Double
    If CostBasis = 0 Then
        ROI = 0
    Else
        ROI = TotalProfit / CostBasis
    End If
End Function

Public Function SummaryDictionary() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "MarketValue", MarketValue
    d.Add "CostBasis", CostBasis
    d.Add "UnrealizedGain", UnrealizedGain
    d.Add "RealizedGain", RealizedGain
    d.Add "Dividends", TotalDividends
    d.Add "ROI", ROI
    Set SummaryDictionary = d
End Function
