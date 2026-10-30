Imports Inventor
Imports System.Globalization

Public Module FlensSaveAssembly
    Public Sub SaveAssemblyCopy(asm As AssemblyDocument, outputFolder As String, model As Input)
        If asm Is Nothing Then Throw New ArgumentNullException(NameOf(asm))
        If model Is Nothing Then Throw New ArgumentNullException(NameOf(model))
        If String.IsNullOrWhiteSpace(outputFolder) Then
            Throw New ArgumentException("Output folder is required.", NameOf(outputFolder))
        End If

        Dim targetPath As String = System.IO.Path.Combine(outputFolder,
            "Samenstelling_" & FormatDimension(model.InputLength) & "x" &
            FormatDimension(model.InputWidth) & ".iam")

        If System.IO.File.Exists(targetPath) Then
            System.IO.File.Delete(targetPath)
        End If

        asm.SaveAs(targetPath, False)

        Dim copiedAsm As AssemblyDocument = Nothing
        Try
            copiedAsm = CType(FlensGlobals.ThisApp.Documents.Open(targetPath, False), AssemblyDocument)

            ' Use the real user parameters, not guessed names
            SetParameter(copiedAsm, "InputLengte", model.InputLength)
            SetParameter(copiedAsm, "InputBreedte", model.InputWidth)
            SetParameter(copiedAsm, "InputFlensHoogte", model.InputFlangeHeight)
            SetParameter(copiedAsm, "InputFlensBreedte", model.InputFlangeWidth)
            SetParameter(copiedAsm, "InputGatDia", model.InputHoleDiameter)

            copiedAsm.Update()
            copiedAsm.Save()
            FlensGlobals.LogInfo("Saved copied assembly: " & targetPath)
        Finally
            If copiedAsm IsNot Nothing Then
                copiedAsm.Close(True)
            End If
        End Try
    End Sub

    Private Sub SetParameter(asmDoc As AssemblyDocument, parameterName As String, value As Double)
        If asmDoc Is Nothing Then Return

        Try
            Dim param As Parameter = asmDoc.ComponentDefinition.Parameters.Item(parameterName)
            If param IsNot Nothing Then
                param.Expression = value.ToString(CultureInfo.CurrentCulture) & " mm"
            End If
        Catch
            ' Ignore if parameter not found
        End Try
    End Sub

    Private Function FormatDimension(value As Double) As String
        Dim rounded As Double = Math.Round(value, 2)
        If rounded = Math.Truncate(rounded) Then
            Return CInt(rounded).ToString(CultureInfo.InvariantCulture)
        Else
            Return rounded.ToString("0.##", CultureInfo.InvariantCulture)
        End If
    End Function
End Module