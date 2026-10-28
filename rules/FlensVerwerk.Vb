Imports Inventor

Public Module FlensVerwerk
    't gene wat rondgestuurt word en de van buite berijkbaar
    Public Function ProcessSeed(oAsm As AssemblyDocument, targetName As String,
        templatePath As String, outputDir As String, model As Input,
        isSquare As Boolean, ByRef occName As String, ByRef outputPath As String) As Boolean
        'lengtes/gatdia aanpassen
        Dim partLength As Double =
            If(targetName = "HoekInY", model.InputLengte, model.InputBreedte)
        Dim holeParameter As String =
            If(isSquare OrElse targetName = "HoekInX", "d134", "d87")
        'Naam output bestand + locatie
        Dim configKey As String =
            targetName & "_" & FileNumber(partLength) & "_" &
            FileNumber(model.InputFlensHoogte) & "_" &
            FileNumber(model.InputFlensBreedte) & "_" &
            FileNumber(model.InputGatDia)
        outputPath = System.IO.Path.Combine(outputDir, configKey & ".ipt")

        'zoekt in de tekening binnen de pattern naar de onderdelen
        Dim occ As ComponentOccurrence = VindOccurrence(oAsm, targetName)
        If occ Is Nothing Then
            FlensGlobals.Show("Occurrence niet gevonden: " & targetName)
            ConfigurePart(outputPath, partLength, holeParameter, model)
            Return False
        End If
        occName = occ.Name
        Try
            If Not System.IO.File.Exists(outputPath) Then
                System.IO.File.Copy(templatePath, outputPath, False)
                FlensGlobals.LogInfo("Nieuw IPT aangemaakt: " & outputPath)
            Else
                If Not CheckOutFile(outputPath) Then
                    FlensGlobals.Show("Checkout mislukt: " & outputPath)
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            FlensGlobals.LogError("Fout bij " & targetName & ": " & ex.ToString())
            FlensGlobals.Show("Fout bij " & targetName & ":" & vbCrLf & ex.Message)
            Return False
        End Try
    End Function
    'de sub functie van het daadwerkelijk zoeken
    Private Function VindOccurrence(oAsm As AssemblyDocument, targetName As String) As ComponentOccurrence
        For Each occ As ComponentOccurrence In oAsm.ComponentDefinition.Occurrences
            If occ.Name = targetName OrElse occ.Name.StartsWith(targetName & ":") Then Return occ
        Next
        Return Nothing
    End Function

    'subfunctie van het verwerken van het onderdeel
    Private Sub ConfigurePart(outputPath As String, partLength As Double,
        holeParameter As String, model As Input)
        Dim partDoc As PartDocument = Nothing
        Try
            'instellen van de part
            partDoc = CType(FlensGlobals.ThisApp.Documents.Open(outputPath, False), PartDocument)
            Dim partParams As Parameters = partDoc.ComponentDefinition.Parameters
            SetMm(partParams, "LengteInw", partLength)
            SetMm(partParams, holeParameter, model.InputGatDia)
            SetMm(partParams, "d23", model.InputFlensHoogte)
            SetMm(partParams, "d24", model.InputFlensBreedte)
            ControleerReadOnly(outputPath)
            partDoc.Update()
            partDoc.Save()
            FlensGlobals.LogInfo("Opgeslagen: " & outputPath)
            partDoc.Close(True)
        Catch
            If partDoc IsNot Nothing Then
                Try : partDoc.Close(False) : Catch : End Try
            End If
            Throw
        End Try
    End Sub
    'TODO: PUNT COMMA CONVERT ERROR? EU/US? 1,2 1.2?
    Private Sub SetMm(partParams As Parameters, parameterName As String, value As Double)
        partParams.Item(parameterName).Expression = FileNumber(value) & " mm"
    End Sub
    
    'readonly errors verhelpen
    Private Sub ControleerReadOnly(outputPath As String)
        If Not System.IO.File.Exists(outputPath) Then Return
        Dim attrs As System.IO.FileAttributes = System.IO.File.GetAttributes(outputPath)
        If (attrs And System.IO.FileAttributes.ReadOnly) <> 0 Then
            Throw New System.IO.IOException(
                "Het IPT-bestand is read-only. Controleer de Vault-status: " & outputPath)
        End If
    End Sub
End Module
