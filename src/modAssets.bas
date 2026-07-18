Attribute VB_Name="modAssets"
Option Explicit

Public Function ValidateAssetInput() As Boolean
    ValidateAssetInput = True
End Function

Public Sub SaveAsset()
    If ValidateAssetInput() Then
        'TODO: Integrar con clsPortfolio
    End If
End Sub
