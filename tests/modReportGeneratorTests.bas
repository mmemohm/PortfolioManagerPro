Attribute VB_Name="modReportGeneratorTests"
Option Explicit
Public Sub TestReportName()
 Debug.Assert Left$(DefaultReportName(),17)="Portfolio_Report"
End Sub
