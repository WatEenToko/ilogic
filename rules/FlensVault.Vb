Imports Inventor

Public Module FlensVault
    'vault spulle los
    Public Function CheckOutFile(value As String) As Boolean
        Return VoerVaultCommandoUit(value, "VaultCheckoutTop", "Check Out")
    End Function

Public Function CheckInFile(value As String) As Boolean
        Return VoerVaultCommandoUit(value, "VaultCheckinTop", "Check In")
    End Function

    'uitvoeren commando
    Private Function VoerVaultCommandoUit(value As String, commandId As String, actie As String) As Boolean
        Try
            Dim oDoc As Inventor.Document = FlensGlobals.ThisApp.Documents.Open(value, True)
            oDoc.Activate()
            Dim oControlDef As Inventor.ControlDefinition =
                FlensGlobals.ThisApp.CommandManager.ControlDefinitions.Item(commandId)
            If Not oControlDef.Enabled Then
                FlensGlobals.LogError("Vault " & actie & " is not enabled for: " & value)
                Return False
            End If
            oControlDef.Execute2(True)
            Return True
        Catch ex As Exception
            FlensGlobals.LogError("Vault " & actie & " failed: " & value & " - " & ex.Message)
            Return False
        End Try
    End Function
End Module
