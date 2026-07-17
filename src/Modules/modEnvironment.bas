Attribute VB_Name = "modEnvironment"
Option Explicit

Public Sub ValidateEnvironment()
    If Application.Version < 16 Then
        Err.Raise vbObjectError + 1000, , "Excel 2019 or later required."
    End If
End Sub

Public Sub ValidateWorkbook()
    'TODO: Verify required worksheets
End Sub
