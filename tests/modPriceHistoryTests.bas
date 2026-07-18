Attribute VB_Name="modPriceHistoryTests"
Option Explicit

Public Sub TestHistoryFileName()
    Debug.Assert HistoryFileName("SPYL") = "history_SPYL.csv"
End Sub
