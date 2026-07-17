Attribute VB_Name="modPortfolioTests"
Option Explicit

Public Sub TestInitialize()
    Dim p As New clsPortfolio
    p.Initialize "Demo","Test"
    Debug.Assert p.Validate = True
End Sub

Public Sub TestClone()
    Dim p As New clsPortfolio
    Dim c As clsPortfolio
    p.Initialize "Original"
    Set c = p.Clone
    Debug.Assert c.Validate = True
End Sub
