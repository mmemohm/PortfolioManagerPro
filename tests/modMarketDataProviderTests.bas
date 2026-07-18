Attribute VB_Name="modMarketDataProviderTests"
Option Explicit

Public Sub TestInitialization()
    Dim p As New clsMarketDataProvider
    p.Initialize
    Debug.Assert p.IsConnected = False
End Sub
