AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensModel.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensGlobals.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensPaden.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensVerwerk.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensVault.vb"
'AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\FlensIsah.vb"

Sub Main()
    'Pass the application and logging to the modules
    FlensGlobals.ThisApp = ThisApplication
    FlensGlobals.InitLogger(AddressOf Logger.Info, AddressOf Logger.Error)

    'Read input parameters from the form or document
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

    'Validate the paths
    Dim paths As New FlensPaden()
    If Not paths.Validate(model) Then Return

    'Build and process the task
    Dim task As New FlangeTask With {
        .Asm = CType(ThisDoc.Document, AssemblyDocument),
        .Model = model,
        .OutputFolder = paths.OutputDir,
        .IsSquare = model.IsSquare}
    If model.IsSquare Then
        task.Start("HoekInY", paths.TemplateX)
        If Not ProcessSeed(task) Then Return
        ReplaceComponent(task)
        task.Start("HoekInX", paths.TemplateX)
        If Not ProcessSeed(task) Then Return
        ReplaceComponent(task)
    Else
        task.Start("HoekInY", paths.TemplateY)
        If Not ProcessSeed(task) Then Return
        ReplaceComponent(task)
        task.Start("HoekInX", paths.TemplateX)
        If Not ProcessSeed(task) Then Return
        ReplaceComponent(task)
    End If

    'If Not SynchroniseerIsah(ThisDoc.Document) Then
    '    FlensGlobals.Show("IsahPrimKey could not be updated.") : Return
    'End If
    InventorVb.DocumentUpdate()
    iLogicVb.UpdateWhenDone = True
    FlensGlobals.Show("The components have been processed.")
End Sub

Sub ReplaceComponent(task As FlangeTask)
    Component.Replace(task.OccurrenceName, task.OutputPath, False)
    InventorVb.DocumentUpdate()
End Sub
