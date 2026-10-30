Imports Inventor

Public Module FlensVerwerk
    Public Function ProcessTask(task As FlangeTask) As Boolean
        DetermineDimensions(task)
        DetermineOutputPath(task)

        If Not FindOccurrence(task) Then
            FlensGlobals.Show("Occurrence not found: " & task.PartName)
            Return False
        End If

        Try
            If Not CreateOrCheckOut(task) Then Return False
            ConfigurePart(task)
            Return True
        Catch ex As Exception
            FlensGlobals.LogError("Error for " & task.PartName & ": " & ex.ToString())
            FlensGlobals.Show("Error for " & task.PartName & ":" & vbCrLf & ex.Message)
            Return False
        End Try
    End Function

    Private Sub DetermineDimensions(task As FlangeTask)
        task.PartLength = If(task.PartName = "HoekInY",
            task.Model.InputLength, task.Model.InputWidth)

        task.HoleParameter = If(task.IsSquare OrElse task.PartName = "HoekInX",
            "d134", "d87")
    End Sub

    Private Sub DetermineOutputPath(task As FlangeTask)
        Dim key As String =
            task.PartName & "_" &
            FormatDimension(task.Model.InputLength) & "x" &
            FormatDimension(task.Model.InputWidth) & "_" &
            FormatDimension(task.Model.InputFlangeHeight) & "_" &
            FormatDimension(task.Model.InputFlangeWidth) & "_" &
            FormatDimension(task.Model.InputHoleDiameter)

        task.OutputPath = System.IO.Path.Combine(task.OutputFolder, key & ".ipt")
    End Sub

    Private Function FormatDimension(value As Double) As String
        Return value.ToString(System.Globalization.CultureInfo.InvariantCulture)
    End Function

    Private Function FindOccurrence(task As FlangeTask) As Boolean
        For Each occ As ComponentOccurrence In task.Asm.ComponentDefinition.Occurrences
            If occ.Name = task.PartName OrElse occ.Name.StartsWith(task.PartName & ":") Then
                task.OccurrenceName = occ.Name
                Return True
            End If
        Next

        Return False
    End Function

    Private Function CreateOrCheckOut(task As FlangeTask) As Boolean
        If Not System.IO.File.Exists(task.OutputPath) Then
            System.IO.File.Copy(task.TemplatePath, task.OutputPath, False)
            RemoveReadOnlyFlag(task.OutputPath)
            FlensGlobals.LogInfo("New IPT created: " & task.OutputPath)
            Return True
        End If

        If Not CheckOutFile(task.OutputPath) Then
            FlensGlobals.Show("Checkout failed: " & task.OutputPath)
            Return False
        End If

        Return True
    End Function

    Private Sub ConfigurePart(task As FlangeTask)
        Dim partDoc As PartDocument = Nothing

        Try
            partDoc = CType(FlensGlobals.ThisApp.Documents.Open(task.OutputPath, False), PartDocument)
            SetParametersIn(partDoc, task)
            CheckReadOnly(task.OutputPath)
            partDoc.Update()
            partDoc.Save()
            FlensGlobals.LogInfo("Saved: " & task.OutputPath)
            partDoc.Close(True)
        Catch
            If partDoc IsNot Nothing Then
                Try
                    partDoc.Close(False)
                Catch
                End Try
            End If
            Throw
        End Try
    End Sub

    Private Sub SetParametersIn(partDoc As PartDocument, task As FlangeTask)
        Dim parameters As Parameters = partDoc.ComponentDefinition.Parameters

        SetMm(parameters, New ParameterValue("LengteInw", task.PartLength))
        SetMm(parameters, New ParameterValue(task.HoleParameter, task.Model.InputHoleDiameter))
        SetMm(parameters, New ParameterValue("d23", task.Model.InputFlangeHeight))
        SetMm(parameters, New ParameterValue("d24", task.Model.InputFlangeWidth))
    End Sub

    Private Sub SetMm(partParams As Parameters, pv As ParameterValue)
        partParams.Item(pv.Name).Expression = MmText(pv.Value)
    End Sub

    Private Sub CheckReadOnly(filePath As String)
        If Not System.IO.File.Exists(filePath) Then Return
        Dim attributes As System.IO.FileAttributes = System.IO.File.GetAttributes(filePath)
        If (attributes And System.IO.FileAttributes.ReadOnly) <> 0 Then
            Throw New System.IO.IOException("The IPT file is read-only: " & filePath)
        End If
    End Sub

    Private Sub RemoveReadOnlyFlag(filePath As String)
        If Not System.IO.File.Exists(filePath) Then Return
        Dim attributes As System.IO.FileAttributes = System.IO.File.GetAttributes(filePath)
        If (attributes And System.IO.FileAttributes.ReadOnly) <> 0 Then
            System.IO.File.SetAttributes(filePath, attributes And Not System.IO.FileAttributes.ReadOnly)
        End If
    End Sub
End Module