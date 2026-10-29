AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensModel.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensGlobals.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensPaden.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensVerwerk.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensVault.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensSaveAssembly.vb"
'AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\FlensIsah.vb"

Sub Main()
    'Application and logger are passed to the modules.
    FlensGlobals.ThisApp = ThisApplication
    FlensGlobals.InitLogger(AddressOf Logger.Info, AddressOf Logger.Error)

    'Read input dimensions from the form or document.
    Dim model As New Input With {
        .InputLength = Parameter("InputLengte"),
        .InputWidth = Parameter("InputBreedte"),
        .InputFlangeHeight = Parameter("InputFlensHoogte"),
        .InputFlangeWidth = Parameter("InputFlensBreedte"),
        .InputHoleDiameter = Parameter("InputGatDia")}
    If Not model.IsValid Then
        FlensGlobals.Show("Please enter all dimensions correctly.")
        Return
    End If

    'Check the required folders and templates.
    Dim paths As New FlensPaden()
    If Not paths.Validate(model) Then Return

    'Generate the new part first, then fix its dimensions, then place it in the assembly.
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

    'Save the assembly as a separate generated copy after the replacement is finished.
    FlensSaveAssembly.SaveAssemblyCopy(assemblyDoc, paths.OutputDir)

    InventorVb.DocumentUpdate()
    iLogicVb.UpdateWhenDone = True
    FlensGlobals.Show("The components have been processed.")
End Sub

Private Function CreateTask(asm As AssemblyDocument, model As Input, outputFolder As String, partName As String, templatePath As String) As FlangeTask
    Return New FlangeTask With {
        .Asm = asm,
        .Model = model,
        .OutputFolder = outputFolder,
        .IsSquare = model.IsSquare,
        .PartName = partName,
        .TemplatePath = templatePath}
End Function

Sub ReplaceComponent(task As FlangeTask)
    'The generated file is created first, then dimensioned, then inserted into the assembly.
    Component.Replace(task.OccurrenceName, task.OutputPath, False)
    InventorVb.DocumentUpdate()
End Sub
