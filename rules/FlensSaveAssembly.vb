Imports Inventor
Imports System.Globalization

Public Module FlensSaveAssembly
    'copieren en opnieuw opslaan
    Public Sub SaveAssemblyCopy(asm As AssemblyDocument, outputFolder As String, model As Input)
        'kijken of de gebruiker de folders heeft
        If asm Is Nothing Then Throw New ArgumentNullException(NameOf(asm))
        If model Is Nothing Then Throw New ArgumentNullException(NameOf(model))
        If String.IsNullOrWhiteSpace(outputFolder) Then
            Throw New ArgumentException("Output folder is required.", NameOf(outputFolder))
        End If
        If Not System.IO.Directory.Exists(outputFolder) Then
            System.IO.Directory.CreateDirectory(outputFolder)
        End If
        
        'benaming van het nieuwe deel
        Dim targetPath As String = System.IO.Path.Combine(outputFolder,
            "Samenstelling_" & FormatDimension(model.InputLength) & "x" &
            FormatDimension(model.InputWidth) & ".iam")
        
        'het copieren zelf
        If System.IO.File.Exists(targetPath) Then
            Dim existingAsm As AssemblyDocument = Nothing
            Try
                existingAsm = CType(FlensGlobals.ThisApp.Documents.Open(targetPath, False), AssemblyDocument)
                ApplyModelToAssembly(existingAsm, model)
                existingAsm.Update()
                existingAsm.Save()
                FlensGlobals.LogInfo("Used existing assembly copy: " & targetPath)
            Finally
                If existingAsm IsNot Nothing Then
                    existingAsm.Close(True)
                End If
            End Try
            Return
        End If

        asm.SaveAs(targetPath, False)
        
        'de gecopieerde aanpassen
        Dim copiedAsm As AssemblyDocument = Nothing
        Try
            copiedAsm = CType(FlensGlobals.ThisApp.Documents.Open(targetPath, False), AssemblyDocument)
            'naam + model>sub plaatsen
            ApplyModelToAssembly(copiedAsm, model)
            copiedAsm.Update()
            copiedAsm.Save()
            FlensGlobals.LogInfo("Created and saved new assembly copy: " & targetPath)
        Finally
            If copiedAsm IsNot Nothing Then
                copiedAsm.Close(True)
            End If
        End Try
    End Sub

    'aanpassing in de file
    Private Sub ApplyModelToAssembly(asmDoc As AssemblyDocument, model As Input)
        If asmDoc Is Nothing Then Return
        SetParameter(asmDoc, "InputLengte", model.InputLength)
        SetParameter(asmDoc, "InputBreedte", model.InputWidth)
        SetParameter(asmDoc, "InputFlensHoogte", model.InputFlangeHeight)
        SetParameter(asmDoc, "InputFlensBreedte", model.InputFlangeWidth)
        SetParameter(asmDoc, "InputGatDia", model.InputHoleDiameter)
    End Sub

    'Oplossing glitch kan bestaande niet vinden
    Private Sub SetParameter(asmDoc As AssemblyDocument, parameterName As String, value As Double)
        If asmDoc Is Nothing Then Return
        Try
            Dim param As Parameter = asmDoc.ComponentDefinition.Parameters.Item(parameterName)
            If param IsNot Nothing Then
                param.Expression = value.ToString(CultureInfo.CurrentCulture) & " mm"
            End If
        Catch
            ' ignore if not found
        End Try
    End Sub

    'naam opmaken van bestandsnaam
    Private Function FormatDimension(value As Double) As String
        Dim rounded As Double = Math.Round(value, 2)
        If rounded = Math.Truncate(rounded) Then
            Return CInt(rounded).ToString(CultureInfo.InvariantCulture)
        End If
        Return rounded.ToString("0.##", CultureInfo.InvariantCulture)
    End Function
End Module