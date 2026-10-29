AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensModel.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensGlobals.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensPaden.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensVerwerk.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensVault.vb"
'AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\FlensIsah.vb"

Sub Main()
    'App en logging doorgeven aan de modules
    FlensGlobals.ThisApp = ThisApplication
    FlensGlobals.InitLogger(AddressOf Logger.Info, AddressOf Logger.Error)

    'Inputparameters uitlezen (form/document)
    Dim model As New Input With {
        .InputLengte = Parameter("InputLengte"),
        .InputBreedte = Parameter("InputBreedte"),
        .InputFlensHoogte = Parameter("InputFlensHoogte"),
        .InputFlensBreedte = Parameter("InputFlensBreedte"),
        .InputGatDia = Parameter("InputGatDia")}
    If Not model.IsValid Then
        FlensGlobals.Show("Vul alle afmetingen correct in.")
        Return
    End If

    'Paden controleren
    Dim paden As New FlensPaden()
    If Not paden.Validate(model) Then Return

    'Taak opbouwen en verwerken
    Dim taak As New FlensTaak With {
        .Asm = CType(ThisDoc.Document, AssemblyDocument),
        .Model = model,
        .UitvoerMap = paden.OutputDir,
        .IsVierkant = model.IsVierkant}
    If model.IsVierkant Then
        taak.Start("HoekInY", paden.TemplateX)
        If Not ProcessSeed(taak) Then Return
        VervangComponent(taak)
        taak.Start("HoekInX", paden.TemplateX)
        If Not ProcessSeed(taak) Then Return
        VervangComponent(taak)
    Else
        taak.Start("HoekInY", paden.TemplateY)
        If Not ProcessSeed(taak) Then Return
        VervangComponent(taak)
        taak.Start("HoekInX", paden.TemplateX)
        If Not ProcessSeed(taak) Then Return
        VervangComponent(taak)
    End If

    'If Not SynchroniseerIsah(ThisDoc.Document) Then
    '    FlensGlobals.Show("IsahPrimKey kon niet worden bijgewerkt.") : Return
    'End If
    InventorVb.DocumentUpdate()
    iLogicVb.UpdateWhenDone = True
    FlensGlobals.Show("De componenten zijn verwerkt.")
End Sub

Sub VervangComponent(taak As FlensTaak)
    Component.Replace(taak.OccNaam, taak.OutputPad, False)
    InventorVb.DocumentUpdate()
End Sub