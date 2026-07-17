VERSION 1.0 CLASS
BEGIN
  MultiUse=-1
END
Attribute VB_Name="clsPortfolio"
Option Explicit
Private mHoldings As Object
Private Sub Class_Initialize()
 Set mHoldings=CreateObject("Scripting.Dictionary")
End Sub
Public Sub Initialize()
 Set mHoldings=CreateObject("Scripting.Dictionary")
End Sub
Public Sub Clear(): mHoldings.RemoveAll: End Sub
Public Function HoldingExists(Symbol As String) As Boolean
 HoldingExists=mHoldings.Exists(UCase$(Trim$(Symbol)))
End Function
Public Sub AddHolding(H As clsHolding)
 mHoldings(UCase$(H.Symbol))=H
End Sub
Public Sub RemoveHolding(Symbol As String)
 If HoldingExists(Symbol) Then mHoldings.Remove UCase$(Trim$(Symbol))
End Sub
Public Function HoldingCount() As Long
 HoldingCount=mHoldings.Count
End Function
Public Function MarketValue() As Double
 Dim k
 For Each k In mHoldings.Keys
  MarketValue=MarketValue+mHoldings(k).MarketValue
 Next
End Function
Public Function CostBasis() As Double
 Dim k
 For Each k In mHoldings.Keys
  CostBasis=CostBasis+mHoldings(k).CostBasis
 Next
End Function
Public Function UnrealizedGain() As Double
 UnrealizedGain=MarketValue-CostBasis
End Function
