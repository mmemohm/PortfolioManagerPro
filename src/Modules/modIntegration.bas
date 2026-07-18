Attribute VB_Name="modIntegration"
Option Explicit

Public Function RunEndToEnd(ByVal Portfolio As clsPortfolio, _
                            ByVal Provider As clsMarketDataProvider) As Boolean
    On Error GoTo EH

    'Expected flow:
    '1. Load settings
    '2. Download market data
    '3. Parse JSON into clsQuote
    '4. Update holdings
    '5. Refresh dashboard
    '6. Save cache / settings

    RunEndToEnd = True
    Exit Function
EH:
    RunEndToEnd = False
End Function
