VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsPortfolio"
Option Explicit

Private mPortfolioID As String
Private mName As String
Private mDescription As String
Private mCreatedDate As Date
Private mLastModified As Date
Private mIsActive As Boolean
Private mHoldings As Collection

Private Sub Class_Initialize()
    Set mHoldings = New Collection
    mCreatedDate = Now
    mLastModified = Now
    mIsActive = True
    mPortfolioID = CreateObject("Scriptlet.TypeLib").GUID
End Sub

Public Property Get PortfolioID() As String
    PortfolioID = mPortfolioID
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal Value As String)
    If Trim$(Value) = "" Then Err.Raise 7001, "clsPortfolio", "Portfolio name cannot be empty."
    mName = Trim$(Value)
    Touch
End Property

Public Property Get Description() As String
    Description = mDescription
End Property

Public Property Let Description(ByVal Value As String)
    mDescription = Value
    Touch
End Property

Public Property Get CreatedDate() As Date
    CreatedDate = mCreatedDate
End Property

Public Property Get LastModified() As Date
    LastModified = mLastModified
End Property

Public Property Get IsActive() As Boolean
    IsActive = mIsActive
End Property

Public Property Let IsActive(ByVal Value As Boolean)
    mIsActive = Value
    Touch
End Property

Public Sub AddHolding(ByVal Holding As Object)
    If Holding Is Nothing Then Err.Raise 7005, "clsPortfolio", "Holding is Nothing."
    On Error GoTo Duplicate
    mHoldings.Add Holding, CStr(Holding.Symbol)
    Touch
    Exit Sub
Duplicate:
    Err.Raise 7003, "clsPortfolio", "Duplicate holding."
End Sub

Public Function HoldingCount() As Long
    HoldingCount = mHoldings.Count
End Function

Public Function HoldingExists(ByVal Symbol As String) As Boolean
    On Error Resume Next
    Dim o As Object
    Set o = mHoldings(Symbol)
    HoldingExists = (Err.Number = 0)
    Err.Clear
End Function

Public Function RemoveHolding(ByVal Symbol As String) As Boolean
    On Error GoTo Fail
    mHoldings.Remove Symbol
    Touch
    RemoveHolding = True
    Exit Function
Fail:
    RemoveHolding = False
End Function

Public Function GetSummary() As String
    GetSummary = "Portfolio: " & mName & vbCrLf & _
                 "ID: " & mPortfolioID & vbCrLf & _
                 "Holdings: " & mHoldings.Count
End Function

Private Sub Touch()
    mLastModified = Now
End Sub
