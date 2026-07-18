Attribute VB_Name="modReportGenerator"
Option Explicit

Public Function GeneratePortfolioReport(ByVal P As clsPortfolio, ByVal O As clsReportOptions) As Boolean
    'TODO:
    '1. Build metrics
    '2. Export Excel
    '3. Export CSV
    '4. Export PDF
    GeneratePortfolioReport=True
End Function

Public Function DefaultReportName() As String
    DefaultReportName="Portfolio_Report_" & Format(Date,"yyyymmdd")
End Function
