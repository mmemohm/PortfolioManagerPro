Attribute VB_Name="modRiskEngine"
Option Explicit

Public Function PositionRiskScore(ByVal WeightPct As Double) As Double
    PositionRiskScore = WeightPct
End Function

Public Function ConcentrationLevel(ByVal WeightPct As Double) As String
    Select Case WeightPct
        Case Is < 10: ConcentrationLevel = "LOW"
        Case Is < 25: ConcentrationLevel = "MEDIUM"
        Case Else: ConcentrationLevel = "HIGH"
    End Select
End Function

Public Function CalculatePortfolioRisk(ByVal P As clsPortfolio) As clsRiskMetrics
    Dim R As New clsRiskMetrics
    'TODO: Iterate holdings and compute aggregate risk.
    Set CalculatePortfolioRisk = R
End Function
