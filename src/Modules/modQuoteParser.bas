Attribute VB_Name="modQuoteParser"
Option Explicit

'Requires VBA-JSON (JsonConverter.bas)

Public Function ParseQuote(Json As Object) As clsQuote
    Dim q As New clsQuote
    q.Symbol = CStr(Json("symbol"))
    q.Price = CDbl(Json("close"))
    q.OpenPrice = CDbl(Json("open"))
    q.HighPrice = CDbl(Json("high"))
    q.LowPrice = CDbl(Json("low"))
    q.Volume = CDbl(Json("volume"))
    q.Currency = CStr(Json("currency"))
    q.ExchangeName = CStr(Json("exchange"))
    Set ParseQuote = q
End Function
