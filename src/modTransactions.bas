Attribute VB_Name="modTransactions"
Option Explicit

Public Enum TransactionType
    ttCompra=1
    ttVenta=2
    ttDividendo=3
End Enum

Public Function SaveTransaction() As Boolean
    'TODO: Integrar con historial y portafolio
    SaveTransaction=True
End Function
