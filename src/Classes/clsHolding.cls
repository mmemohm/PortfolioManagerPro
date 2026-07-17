VERSION 1.0 CLASS
BEGIN
 MultiUse=-1
END
Attribute VB_Name="clsHolding"
Option Explicit
Private mSymbol As String,mQuantity As Double,mAverageCost As Double,mMarketPrice As Double
Public Property Get Symbol() As String: Symbol=mSymbol: End Property
Public Property Let Symbol(ByVal v As String): mSymbol=UCase$(Trim$(v)): End Property
Public Property Get Quantity() As Double: Quantity=mQuantity: End Property
Public Property Let Quantity(ByVal v As Double): If v<0 Then Err.Raise 7101: mQuantity=v: End Property
Public Property Let AverageCost(ByVal v As Double): mAverageCost=v: End Property
Public Property Let MarketPrice(ByVal v As Double): mMarketPrice=v: End Property
Public Property Get CostBasis() As Double: CostBasis=mQuantity*mAverageCost: End Property
Public Property Get MarketValue() As Double: MarketValue=mQuantity*mMarketPrice: End Property
Public Property Get UnrealizedGain() As Double: UnrealizedGain=MarketValue-CostBasis: End Property
