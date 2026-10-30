AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensModel.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensGlobals.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensPaden.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensVerwerk.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensVault.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensSaveAssembly.vb"

Sub Main()
    FlensGlobals.ThisApp = ThisApplication
    FlensGlobals.InitLogger(AddressOf Logger.Info, AddressOf Logger.Error)

    Dim model As New Input With {
        .InputLength = Parameter("InputLengte"),
        .InputWidth = Parameter("InputBreedte"),
        .InputFlangeHeight = Parameter("InputFlensHoogte"),
        .InputFlangeWidth = Parameter("InputFlensBreedte"),
        .InputHoleDiameter = Parameter("InputGatDia")
    }

    If Not model.IsValid Then
        FlensGlobals.Show("Please enter all dimensions correctly.")
        Return
    End If

    Dim paths As New FlensPaden()
    If Not paths.Validate(model) Then
        Return
    End If

    Dim assemblyDoc As AssemblyDocument = CType(ThisDoc.Document, AssemblyDocument)
    Dim task As FlangeTask

    If model.IsSquare Then
        task = CreateTask(assemblyDoc, model, paths.OutputDir, "HoekInY", paths.TemplateX)
        If Not ProcessSeed(task) Then Return
        ReplaceComponent(task)

        task = CreateTask(assemblyDoc, model, paths.OutputDir, "HoekInX", paths.TemplateX)
        If Not ProcessSeed(task) Then Return
        ReplaceComponent(task)
    Else
        task = CreateTask(assemblyDoc, model, paths.OutputDir, "HoekInY", paths.TemplateY)
        If Not ProcessSeed(task) Then Return
        ReplaceComponent(task)

        task = CreateTask(assemblyDoc, model, paths.OutputDir, "HoekInX", paths.TemplateX)
        If Not ProcessSeed(task) Then Return
        ReplaceComponent(task)
    End If

    FlensSaveAssembly.SaveAssemblyCopy(assemblyDoc, paths.OutputDir, model)

    FlensGlobals.Show("The components have been processed.")
End Sub

Private Function CreateTask(asm As AssemblyDocument,
                            model As Input,
                            outputFolder As String,
                            partName As String,
                            templatePath As String) As FlangeTask
    Return New FlangeTask With {
        .Asm = asm,
        .Model = model,
        .OutputFolder = outputFolder,
        .IsSquare = model.IsSquare,
        .PartName = partName,
        .TemplatePath = templatePath
    }
End Function

Sub ReplaceComponent(task As FlangeTask)
    Component.Replace(task.OccurrenceName, task.OutputPath, False)
End Sub