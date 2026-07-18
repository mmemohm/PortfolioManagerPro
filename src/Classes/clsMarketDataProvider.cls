VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsMarketDataProvider"
Option Explicit

Private mLastUpdate As Date

Public Sub Initialize()
    mLastUpdate = 0
End Sub

Public Function GetQuote(ByVal Symbol As String) As Double
    ' TODO: Implementar proveedor real (API)
    Err.Raise 8101, , "Market data provider not configured."
End Function

Public Function GetLastUpdate() As Date
    GetLastUpdate = mLastUpdate
End Function

Public Sub SetLastUpdate(ByVal Value As Date)
    mLastUpdate = Value
End Sub

Public Function IsConnected() As Boolean
    IsConnected = False
End Function
