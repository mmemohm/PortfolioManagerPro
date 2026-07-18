Attribute VB_Name="modPersistenceTests"
Option Explicit

Public Sub TestCacheFileName()
    Debug.Assert CacheFileName("TSM")="cache_TSM.json"
End Sub
