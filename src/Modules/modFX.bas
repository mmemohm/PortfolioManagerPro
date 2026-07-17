Attribute VB_Name = "modFX"
Option Explicit

Public Function ConvertCurrency(ByVal Amount As Double, ByVal Rate As Double) As Double
    ConvertCurrency = Amount * Rate
End Function

Public Function TotalCost(ByVal Shares As Double, ByVal Price As Double, ByVal Commission As Double) As Double
    TotalCost = Shares * Price + Commission
End Function
