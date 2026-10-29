Imports Inventor

Public Module FlensVault
    'vault item handling
    Public Function CheckOutFile(value As String) As Boolean
        Return ExecuteVaultCommand(value, "VaultCheckoutTop", "Check Out")
    End Function

    Public Function CheckInFile(value As String) As Boolean
        Return ExecuteVaultCommand(value, "VaultCheckinTop", "Check In")
    End Function

    'execute the Vault command
    Private Function ExecuteVaultCommand(value As String, commandId As String, action As String) As Boolean
        Try
            Dim document As Inventor.Document = FlensGlobals.ThisApp.Documents.Open(value, True)
            document.Activate()
            Dim controlDefinition As Inventor.ControlDefinition =
                FlensGlobals.ThisApp.CommandManager.ControlDefinitions.Item(commandId)
            If Not controlDefinition.Enabled Then
                FlensGlobals.LogError("Vault " & action & " is not enabled for: " & value)
                Return False
            End If
            controlDefinition.Execute2(True)
            Return True
        Catch ex As Exception
            FlensGlobals.LogError("Vault " & action & " failed: " & value & " - " & ex.Message)
            Return False
        End Try
    End Function
End Module
