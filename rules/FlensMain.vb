AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensGlobals.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensPaden.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensVerwerk.vb"
AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\V2\ilogic\rules\FlensVault.vb"
'AddVbFile "L:\CAD_Conf2025\Inventor 2025\Design Data\iLogic\Flenzen\FlensIsah.vb"

Sub Main()
    FlensGlobals.ThisApp = ThisApplication
    FlensGlobals.InitLogger(AddressOf Logger.Info, AddressOf Logger.Error)

    Dim model As New Input With {
    If Not model.IsValid Then
        Return
    End If


        .Asm = CType(ThisDoc.Document, AssemblyDocument),
        .Model = model,
        .OutputFolder = paths.OutputDir,
    Else
    End If

    'If Not SynchroniseerIsah(ThisDoc.Document) Then
    'End If
    InventorVb.DocumentUpdate()
    iLogicVb.UpdateWhenDone = True
End Sub

    InventorVb.DocumentUpdate()
