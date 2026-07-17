VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name="clsPortfolio"
Option Explicit

Private mName As String
Private mDescription As String

Public Sub Initialize(ByVal PortfolioName As String, Optional ByVal Desc As String="")
    If Trim$(PortfolioName)="" Then Err.Raise 7001,"clsPortfolio","Portfolio name required."
    mName=Trim$(PortfolioName)
    mDescription=Desc
End Sub

Public Sub Clear()
    mName=""
    mDescription=""
End Sub

Public Function Validate() As Boolean
    Validate=(Len(mName)>0)
End Function

Public Function Clone() As clsPortfolio
    Dim p As New clsPortfolio
    p.Initialize mName,mDescription
    Set Clone=p
End Function

Public Function ExportDTO() As Object
    Dim d As Object
    Set d=CreateObject("Scripting.Dictionary")
    d("Name")=mName
    d("Description")=mDescription
    Set ExportDTO=d
End Function

Public Sub ImportDTO(ByVal dto As Object)
    If dto Is Nothing Then Err.Raise 7006,"clsPortfolio","DTO invalid."
    Initialize CStr(dto("Name")), CStr(dto("Description"))
End Sub
