Attribute VB_Name = "modPortfolioMetricsTests"
Option Explicit

Public Sub TestReturn()
    Debug.Assert CalculateGainLoss(100,125)=25
    Debug.Assert Round(CalculateReturnPct(100,125),2)=25
    Debug.Assert Round(CalculateWeight(20,100),2)=20
End Sub
