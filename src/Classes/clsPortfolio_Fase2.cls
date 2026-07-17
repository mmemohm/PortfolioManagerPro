VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsPortfolio"

Option Explicit

Private mHoldings As Object 'Scripting.Dictionary

Private Sub Class_Initialize()
    Set mHoldings = CreateObject("Scripting.Dictionary")
End Sub

Public Sub AddHolding(ByVal Holding As clsHolding)
    If Holding Is Nothing Then Err.Raise 7005, "clsPortfolio", "Holding is Nothing."

    If Trim$(Holding.Symbol) = "" Then
        Err.Raise 7006, "clsPortfolio", "Empty symbol."
    End If

    If mHoldings.Exists(UCase$(Holding.Symbol)) Then
        Err.Raise 7003, "clsPortfolio", "Duplicate holding."
    End If

    mHoldings.Add UCase$(Holding.Symbol), Holding
End Sub

Public Function GetHolding(ByVal Symbol As String) As clsHolding
    If mHoldings.Exists(UCase$(Symbol)) Then
        Set GetHolding = mHoldings(UCase$(Symbol))
    Else
        Err.Raise 7004, "clsPortfolio", "Holding not found."
    End If
End Function

Public Function HoldingExists(ByVal Symbol As String) As Boolean
    HoldingExists = mHoldings.Exists(UCase$(Symbol))
End Function

Public Function HoldingCount() As Long
    HoldingCount = mHoldings.Count
End Function

Public Function RemoveHolding(ByVal Symbol As String) As Boolean
    If mHoldings.Exists(UCase$(Symbol)) Then
        mHoldings.Remove UCase$(Symbol)
        RemoveHolding = True
    Else
        RemoveHolding = False
    End If
End Function

Public Sub UpdateHolding(ByVal Holding As clsHolding)
    If Holding Is Nothing Then Err.Raise 7005
    If Not mHoldings.Exists(UCase$(Holding.Symbol)) Then
        Err.Raise 7004
    End If
    Set mHoldings(UCase$(Holding.Symbol)) = Holding
End Sub

Public Function GetAllHoldings() As Collection
    Dim c As New Collection
    Dim k
    For Each k In mHoldings.Keys
        c.Add mHoldings(k)
    Next
    Set GetAllHoldings = c
End Function
