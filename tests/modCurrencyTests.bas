Attribute VB_Name="modCurrencyTests"
Option Explicit
Public Sub TestCurrency()
    Dim fx As New clsCurrency
    fx.Initialize "USD"
    fx.AddExchangeRate "USD","MXN",18.5
    Debug.Assert fx.Convert(100,"USD","MXN")=1850
End Sub
