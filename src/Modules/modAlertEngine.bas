Attribute VB_Name="modAlertEngine"
Option Explicit

Public Function PriceDropAlert(ByVal ChangePct As Double, ByVal Threshold As Double) As Boolean
    PriceDropAlert = (ChangePct <= -Abs(Threshold))
End Function

Public Function PriceTargetAlert(ByVal CurrentPrice As Double, ByVal TargetPrice As Double) As Boolean
    PriceTargetAlert = (CurrentPrice >= TargetPrice)
End Function

Public Function NewHighAlert(ByVal CurrentPrice As Double, ByVal PreviousHigh As Double) As Boolean
    NewHighAlert = (CurrentPrice > PreviousHigh)
End Function

Public Function EvaluatePortfolioAlerts(ByVal P As clsPortfolio) As Collection
    Dim Alerts As New Collection
    'TODO: recorrer holdings y generar clsAlert según reglas.
    Set EvaluatePortfolioAlerts = Alerts
End Function
