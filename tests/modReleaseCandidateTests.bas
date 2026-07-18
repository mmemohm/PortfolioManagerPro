Attribute VB_Name="modReleaseCandidateTests"
Option Explicit

Public Sub SmokeRelease()
    Debug.Assert ValidateReleaseCandidate() = True
End Sub
