Attribute VB_Name = "modLogger"
Option Explicit

Public Enum LogLevel
    LOG_INFO=1
    LOG_WARNING=2
    LOG_ERROR=3
    LOG_DEBUG=4
End Enum

Public Sub LogMessage(ByVal Level As LogLevel, ByVal Source As String, ByVal Message As String)
    Debug.Print Now & " | " & Level & " | " & Source & " | " & Message
    'TODO: Persist to Logs worksheet
End Sub

Public Sub LogInfo(ByVal Source As String, ByVal Message As String)
    LogMessage LOG_INFO, Source, Message
End Sub

Public Sub LogError(ByVal Source As String, ByVal Message As String)
    LogMessage LOG_ERROR, Source, Message
End Sub
