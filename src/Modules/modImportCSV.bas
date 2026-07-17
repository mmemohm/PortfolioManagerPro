Attribute VB_Name = "modImportCSV"
Option Explicit

Public Function ValidateCsvRow(ByVal Ticker As String, ByVal Qty As Double, ByVal Price As Double) As Boolean
    ValidateCsvRow = (Len(Trim$(Ticker)) > 0 And Qty <> 0 And Price >= 0)
End Function

Public Sub ImportTransactionsCsv(ByVal FilePath As String)
    'TODO: Parse CSV and create clsTransaction instances
End Sub
