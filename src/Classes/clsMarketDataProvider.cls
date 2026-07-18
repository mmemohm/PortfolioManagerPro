VERSION 1.0 CLASS
BEGIN
  MultiUse=-1
END
Attribute VB_Name="clsMarketDataProvider"
Option Explicit

Private Const BASE_URL As String = "https://api.twelvedata.com/quote"
Private mApiKey As String

Public Sub Initialize(ByVal ApiKey As String)
    mApiKey = Trim$(ApiKey)
End Sub

Public Function BuildQuoteUrl(ByVal Symbol As String) As String
    BuildQuoteUrl = BASE_URL & "?symbol=" & _
        Replace(UCase$(Trim$(Symbol)), " ", "%20") & _
        "&apikey=" & mApiKey
End Function

Public Function IsConfigured() As Boolean
    IsConfigured = (Len(mApiKey) > 0)
End Function

Public Function GetQuoteJson(ByVal Symbol As String) As String
    If Not IsConfigured Then Err.Raise 8101,, "API Key not configured."
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", BuildQuoteUrl(Symbol), False
    http.send
    If http.Status <> 200 Then
        Err.Raise 8102,, "HTTP Error: " & http.Status
    End If
    GetQuoteJson = http.responseText
End Function
