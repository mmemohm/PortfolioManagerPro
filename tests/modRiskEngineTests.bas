Attribute VB_Name="modRiskEngineTests"
Option Explicit

Public Sub TestRisk()
    Debug.Assert PositionRiskScore(15)=15
    Debug.Assert ConcentrationLevel(5)="LOW"
    Debug.Assert ConcentrationLevel(15)="MEDIUM"
    Debug.Assert ConcentrationLevel(30)="HIGH"
End Sub
