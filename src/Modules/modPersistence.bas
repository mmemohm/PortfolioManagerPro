Attribute VB_Name="modPersistence"
Option Explicit

Public Function LoadSettings(ByVal FileName As String, ByRef S As clsSettings) As Boolean
    'TODO: Parse INI/JSON configuration
    S.ConfigFile = FileName
    LoadSettings = True
End Function

Public Function SaveSettings(ByVal FileName As String, ByVal S As clsSettings) As Boolean
    'TODO: Persist settings to disk
    SaveSettings = True
End Function

Public Function CacheFileName(ByVal Symbol As String) As String
    CacheFileName = "cache_" & UCase$(Symbol) & ".json"
End Function
