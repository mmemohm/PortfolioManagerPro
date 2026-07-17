VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsDividendEngine"
Option Explicit

Private Type tDividend
    Symbol As String
    PayDate As Date
    GrossAmount As Double
    TaxWithheld As Double
    Currency As String
    Reinvested As Boolean
End Type

Private mDividends As Object

Private Sub Class_Initialize()
    Set mDividends = CreateObject("Scripting.Dictionary")
End Sub

Public Sub Initialize()
    Set mDividends = CreateObject("Scripting.Dictionary")
End Sub

Public Sub AddDividend(ByVal Id As String, ByVal Symbol As String, _
    ByVal PayDate As Date, ByVal Gross As Double, _
    Optional ByVal Tax As Double = 0, _
    Optional ByVal Currency As String = "USD", _
    Optional ByVal Reinvested As Boolean = False)

    Dim d As tDividend
    d.Symbol = UCase$(Trim$(Symbol))
    d.PayDate = PayDate
    d.GrossAmount = Gross
    d.TaxWithheld = Tax
    d.Currency = UCase$(Currency)
    d.Reinvested = Reinvested
    mDividends(Id)=d
End Sub

Public Function DividendCount() As Long
    DividendCount=mDividends.Count
End Function

Public Function TotalGross() As Double
    Dim k,v
    For Each k In mDividends.Keys
        LSet v = mDividends(k)
        TotalGross=TotalGross+v.GrossAmount
    Next
End Function

Public Function TotalTax() As Double
    Dim k,v
    For Each k In mDividends.Keys
        LSet v = mDividends(k)
        TotalTax=TotalTax+v.TaxWithheld
    Next
End Function

Public Function TotalNet() As Double
    TotalNet=TotalGross()-TotalTax()
End Function

Public Sub Clear()
    mDividends.RemoveAll
End Sub
