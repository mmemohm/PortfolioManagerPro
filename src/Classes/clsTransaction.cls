VERSION 1.0 CLASS
BEGIN
 MultiUse=-1
END
Attribute VB_Name="clsTransaction"
Option Explicit
Public TradeDate As Date
Public Symbol As String
Public Quantity As Double
Public Price As Double
Public Fees As Double
Public TransactionType As String
Public Function GrossAmount() As Double: GrossAmount=Quantity*Price: End Function
Public Function NetAmount() As Double: NetAmount=GrossAmount+Fees: End Function
