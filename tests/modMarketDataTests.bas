Attribute VB_Name="modMarketDataTests"
Option Explicit
Public Sub TestUrl()
 Dim p As New clsMarketDataProvider
 p.Initialize "TESTKEY"
 Debug.Assert InStr(p.BuildQuoteUrl("TSM"),"apikey=TESTKEY")>0
End Sub
