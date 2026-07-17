Attribute VB_Name="modErrorHandler"
Option Explicit
Public Sub HandleError(ByVal ProcName As String)
    LogError ProcName, Err.Number & " - " & Err.Description
End Sub
