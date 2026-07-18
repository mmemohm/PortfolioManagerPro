Attribute VB_Name="modAlertEngineTests"
Option Explicit

Public Sub TestAlerts()
    Debug.Assert PriceDropAlert(-6,5)=True
    Debug.Assert PriceDropAlert(-2,5)=False
    Debug.Assert PriceTargetAlert(105,100)=True
    Debug.Assert NewHighAlert(210,200)=True
End Sub
