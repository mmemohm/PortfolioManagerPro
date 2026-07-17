Attribute VB_Name = "modMain"
Option Explicit

Public Sub InitializeApplication()
    On Error GoTo EH

    Call ValidateEnvironment
    Call LoadConfiguration
    Call InitializeGlobals
    Call InitializeLogger
    Call ValidateWorkbook
    Exit Sub

EH:
    MsgBox "Initialization error: " & Err.Description, vbCritical
End Sub
