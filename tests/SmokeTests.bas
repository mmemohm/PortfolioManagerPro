Attribute VB_Name="SmokeTests"
Option Explicit
Public Sub RunSmokeTests()
    Debug.Assert ApplicationVersion = "1.1.0 Beta"
End Sub
