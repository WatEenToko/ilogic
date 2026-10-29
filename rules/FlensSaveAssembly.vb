Imports Inventor
Imports System.Globalization

Public Module FlensSaveAssembly
    Public Sub SaveAssemblyCopy(asm As AssemblyDocument, outputFolder As String, model As Input)
        If asm Is Nothing Then
            Throw New ArgumentNullException(NameOf(asm))
        End If

        If String.IsNullOrWhiteSpace(outputFolder) Then
            Throw New ArgumentException("Output folder is required.", NameOf(outputFolder))
        End If

        If model Is Nothing Then
            Throw New ArgumentNullException(NameOf(model))
        End If

        Dim lengthText As String = FormatDimension(model.InputLength)
        Dim widthText As String = FormatDimension(model.InputWidth)
        Dim fileBaseName As String = "Samenstelling_" & lengthText & "x" & widthText

        Dim targetPath As String = System.IO.Path.Combine(outputFolder, fileBaseName & ".iam")

        If System.IO.File.Exists(targetPath) Then
            System.IO.File.Delete(targetPath)
        End If

        asm.SaveAs(targetPath, False)

        Dim copiedAsm As AssemblyDocument = Nothing
        Try
            copiedAsm = CType(FlensGlobals.ThisApp.Documents.Open(targetPath, False), AssemblyDocument)
            UpdateAssemblyDimensions(copiedAsm, model.InputLength, model.InputWidth)
            copiedAsm.Update()
            copiedAsm.Save()
            FlensGlobals.LogInfo("Saved copied assembly: " & targetPath)
        Finally
            If copiedAsm IsNot Nothing Then
                copiedAsm.Close(True)
            End If
        End Try
    End Sub

    Private Sub UpdateAssemblyDimensions(asmDoc As AssemblyDocument, newLength As Double, newWidth As Double)
        If asmDoc Is Nothing Then Return

        Dim parameters As Parameters = asmDoc.ComponentDefinition.Parameters

        SetParameterIfExists(parameters, "Lengte", newLength)
        SetParameterIfExists(parameters, "Breedte", newWidth)
        SetParameterIfExists(parameters, "Length", newLength)
        SetParameterIfExists(parameters, "Width", newWidth)
        SetParameterIfExists(parameters, "L", newLength)
        SetParameterIfExists(parameters, "B", newWidth)
    End Sub

    Private Sub SetParameterIfExists(parameters As Parameters, name As String, value As Double)
        Try
            Dim p As Parameter = parameters.Item(name)
            If p IsNot Nothing Then
                p.Expression = MmText(value)
            End If
        Catch
            ' ignore if parameter does not exist
        End Try
    End Sub

    Private Function FormatDimension(value As Double) As String
        Dim rounded As Double = Math.Round(value, 2)
        If rounded = Math.Truncate(rounded) Then
            Return CInt(rounded).ToString(CultureInfo.InvariantCulture)
        End If
        Return rounded.ToString("0.##", CultureInfo.InvariantCulture)
    End Function

    Private Function MmText(value As Double) As String
        Return value.ToString(CultureInfo.CurrentCulture) & " mm"
    End Function
End Module