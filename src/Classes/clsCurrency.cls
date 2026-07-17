VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsCurrency"
Option Explicit

Private mBaseCurrency As String
Private mRates As Object

Private Sub Class_Initialize()
    Set mRates = CreateObject("Scripting.Dictionary")
    mBaseCurrency = "USD"
End Sub

Public Sub Initialize(Optional ByVal BaseCode As String = "USD")
    Set mRates = CreateObject("Scripting.Dictionary")
    SetBaseCurrency BaseCode
End Sub

Public Sub SetBaseCurrency(ByVal ISOCode As String)
    If Not ValidateISOCode(ISOCode) Then Err.Raise 7201,, "Invalid ISO currency code."
    mBaseCurrency = UCase$(ISOCode)
End Sub

Public Property Get BaseCurrency() As String
    BaseCurrency = mBaseCurrency
End Property

Public Function ValidateISOCode(ByVal ISOCode As String) As Boolean
    ISOCode = UCase$(Trim$(ISOCode))
    Select Case ISOCode
        Case "USD","MXN","EUR","GBP","JPY","CAD","CHF","CNY","AUD"
            ValidateISOCode = True
    End Select
End Function

Public Sub AddExchangeRate(ByVal FromCode As String, ByVal ToCode As String, ByVal Rate As Double)
    If Rate <= 0 Then Err.Raise 7202,, "Rate must be positive."
    mRates(UCase$(FromCode) & ">" & UCase$(ToCode)) = Rate
End Sub

Public Function ExchangeRate(ByVal FromCode As String, ByVal ToCode As String) As Double
    Dim k As String
    k = UCase$(FromCode) & ">" & UCase$(ToCode)
    If mRates.Exists(k) Then
        ExchangeRate = mRates(k)
    ElseIf UCase$(FromCode)=UCase$(ToCode) Then
        ExchangeRate = 1#
    Else
        Err.Raise 7203,, "Exchange rate not found."
    End If
End Function

Public Function Convert(ByVal Amount As Double, ByVal FromCode As String, ByVal ToCode As String) As Double
    Convert = Round(Amount * ExchangeRate(FromCode, ToCode), 2)
End Function

Public Function CurrencyExists(ByVal FromCode As String, ByVal ToCode As String) As Boolean
    CurrencyExists = mRates.Exists(UCase$(FromCode) & ">" & UCase$(ToCode))
End Function

Public Sub RemoveExchangeRate(ByVal FromCode As String, ByVal ToCode As String)
    Dim k As String
    k = UCase$(FromCode) & ">" & UCase$(ToCode)
    If mRates.Exists(k) Then mRates.Remove k
End Sub

Public Sub Clear()
    mRates.RemoveAll
End Sub

Public Function Clone() As clsCurrency
    Dim c As New clsCurrency, k
    c.Initialize mBaseCurrency
    For Each k In mRates.Keys
        c.AddExchangeRate Split(k, ">")(0), Split(k, ">")(1), mRates(k)
    Next
    Set Clone = c
End Function
