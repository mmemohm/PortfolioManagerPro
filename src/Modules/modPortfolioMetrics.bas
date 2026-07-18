Attribute VB_Name = "modPortfolioMetrics"
Option Explicit

Public Function CalculateGainLoss(ByVal Cost As Double, ByVal MarketValue As Double) As Double
    CalculateGainLoss = MarketValue - Cost
End Function

Public Function CalculateReturnPct(ByVal Cost As Double, ByVal MarketValue As Double) As Double
    If Cost = 0 Then
        CalculateReturnPct = 0
    Else
        CalculateReturnPct = ((MarketValue - Cost) / Cost) * 100#
    End If
End Function

Public Function CalculateWeight(ByVal PositionValue As Double, ByVal PortfolioValue As Double) As Double
    If PortfolioValue = 0 Then
        CalculateWeight = 0
    Else
        CalculateWeight = (PositionValue / PortfolioValue) * 100#
    End If
End Function

Public Function BuildPortfolioMetrics(ByVal P As clsPortfolio) As clsPortfolioMetrics
    Dim M As New clsPortfolioMetrics
    'TODO: recorrer holdings y completar métricas.
    Set BuildPortfolioMetrics = M
End Function
