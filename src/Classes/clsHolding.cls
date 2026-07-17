VERSION 1.0 CLASS
BEGIN
  MultiUse=-1
END
Attribute VB_Name="clsHolding"
Option Explicit

Private mSymbol As String
Private mQuantity As Double
Private mAverageCost As Double
Private mMarketPrice As Double
Private mRealizedGain As Double
Private mDividendsReceived As Double

Public Property Get Symbol() As String: Symbol=mSymbol: End Property
Public Property Let Symbol(ByVal v As String): mSymbol=UCase$(Trim$(v)): End Property
Public Property Get Quantity() As Double: Quantity=mQuantity: End Property
Public Property Let Quantity(ByVal v As Double)
 If v<0 Then Err.Raise 7101,, "Negative quantity"
 mQuantity=v
End Property
Public Property Let AverageCost(ByVal v As Double): mAverageCost=v: End Property
Public Property Let MarketPrice(ByVal v As Double): mMarketPrice=v: End Property
Public Property Get CostBasis() As Double: CostBasis=mQuantity*mAverageCost: End Property
Public Property Get MarketValue() As Double: MarketValue=mQuantity*mMarketPrice: End Property
Public Property Get UnrealizedGain() As Double: UnrealizedGain=MarketValue-CostBasis: End Property
Public Property Get RealizedGain() As Double: RealizedGain=mRealizedGain: End Property
Public Property Let RealizedGain(ByVal v As Double): mRealizedGain=v: End Property
Public Property Get DividendsReceived() As Double: DividendsReceived=mDividendsReceived: End Property
Public Property Let DividendsReceived(ByVal v As Double): mDividendsReceived=v: End Property
Public Function ROI() As Double
 If CostBasis=0 Then ROI=0 Else ROI=(UnrealizedGain+RealizedGain+DividendsReceived)/CostBasis
End Function
