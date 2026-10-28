Imports Inventor

'AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\FlensModel.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\FlensGlobals.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\FlensModel.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\FlensPaden.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\FlensVerwerk.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\FlensVault.vb"
'AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\FlensIsah.vb"

Sub Main()
    'paden fixen voor OOP progammering
    FlensGlobals.Init(ThisApplication, AddressOf Logger.Info, AddressOf Logger.Error)
    'model checken
    Dim model As New Input With {
        .InputLengte = Parameter("InputLengte"), 
        .InputBreedte = Parameter("InputBreedte"),
        .InputFlensHoogte = Parameter("InputFlensHoogte"), 
        .InputFlensBreedte = Parameter("InputFlensBreedte"),
        .InputGatDia = Parameter("InputGatDia")}
    If Not model.IsValid Then
        Logger.Info("breedte = " & model.InputBreedte & " Lengte = " & model.InputLengte & " Gat dia = " & model.InputGatDia)
        FlensGlobals.Show("Vul alle afmetingen correct in.") : Return
    End If
    'paden doorgeven
    Dim paden As New FlensPaden()
    If Not paden.Validate(model) Then Return

    'iam aanpassen
    Dim oAsm As AssemblyDocument = CType(ThisDoc.Document, AssemblyDocument)
    Dim occName As String, outputPath As String
    'uitwisselingen
    If model.IsVierkant Then
        If Not ProcessSeed(oAsm, "HoekInY", paden.TemplateX, paden.OutputDir, model, True, occName, outputPath) Then Return
        Component.Replace(occName, outputPath, False) : InventorVb.DocumentUpdate()
        If Not ProcessSeed(oAsm, "HoekInX", paden.TemplateX, paden.OutputDir, model, True, occName, outputPath) Then Return
        Component.Replace(occName, outputPath, False) : InventorVb.DocumentUpdate()
    Else
        If Not ProcessSeed(oAsm, "HoekInY", paden.TemplateY, paden.OutputDir, model, False, occName, outputPath) Then Return
        Component.Replace(occName, outputPath, False) : InventorVb.DocumentUpdate()
        If Not ProcessSeed(oAsm, "HoekInX", paden.TemplateX, paden.OutputDir, model, False, occName, outputPath) Then Return
        Component.Replace(occName, outputPath, False) : InventorVb.DocumentUpdate()
    End If

    'If Not SynchroniseerIsah(ThisDoc.Document) Then
    '    FlensGlobals.Show("IsahPrimKey kon niet worden bijgewerkt.") : Return
    'End If
    InventorVb.DocumentUpdate()
    iLogicVb.UpdateWhenDone = True
    FlensGlobals.Show("De componenten zijn verwerkt.")
End Sub
