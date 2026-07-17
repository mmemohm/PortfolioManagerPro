Attribute VB_Name="modDividendTests"
Option Explicit
Public Sub TestDividendEngine()
    Dim d As New clsDividendEngine
    d.Initialize
    d.AddDividend "DIV1","SPYL",Date,100,15,"USD",False
    Debug.Assert d.DividendCount=1
    Debug.Assert d.TotalGross=100
    Debug.Assert d.TotalTax=15
    Debug.Assert d.TotalNet=85
End Sub
