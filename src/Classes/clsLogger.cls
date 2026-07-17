VERSION 1.0
Attribute VB_Name = "clsLogger"
Option Explicit

Private mSource As String

Public Property Let Source(ByVal Value As String)
    mSource = Value
End Property

Public Sub Info(ByVal Message As String)
    LogInfo mSource, Message
End Sub

Public Sub [Error](ByVal Message As String)
    LogError mSource, Message
End Sub
