Attribute VB_Name="modPriceHistory"
Option Explicit

Public Function CreateHistoryRecord( _
    ByVal Symbol As String, _
    ByVal QuoteDate As Date, _
    ByVal OpenPrice As Double, _
    ByVal HighPrice As Double, _
    ByVal LowPrice As Double, _
    ByVal ClosePrice As Double, _
    ByVal Volume As Double) As clsPriceHistory

    Dim H As New clsPriceHistory

    H.Symbol = Symbol
    H.QuoteDate = QuoteDate
    H.OpenPrice = OpenPrice
    H.HighPrice = HighPrice
    H.LowPrice = LowPrice
    H.ClosePrice = ClosePrice
    H.Volume = Volume

    Set CreateHistoryRecord = H
End Function

Public Function HistoryFileName(ByVal Symbol As String) As String
    HistoryFileName = "history_" & UCase$(Symbol) & ".csv"
End Function

Public Sub AppendHistoryRecord(ByVal History As clsPriceHistory)
    'TODO:
    'Append record to CSV or worksheet history.
End Sub
