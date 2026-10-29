Imports Inventor

Public Module FlensVerwerk
    Public Function ProcessSeed(taak As FlensTaak) As Boolean
        BepaalAfmetingen(taak)
        BepaalOutputPad(taak)
        If Not VindOccurrence(taak) Then
            FlensGlobals.Show("Occurrence niet gevonden: " & taak.Onderdeelnaam)
            Return False
        End If
        Try
            If Not MaakOfCheckOut(taak) Then Return False
            ConfigurePart(taak)
            Return True
        Catch ex As Exception
            FlensGlobals.LogError("Fout bij " & taak.Onderdeelnaam & ": " & ex.ToString())
            FlensGlobals.Show("Fout bij " & taak.Onderdeelnaam & ":" & vbCrLf & ex.Message)
            Return False
        End Try
    End Function

    Private Sub BepaalAfmetingen(taak As FlensTaak)
        taak.PartLengte = If(taak.Onderdeelnaam = "HoekInY",
            taak.Model.InputLengte, taak.Model.InputBreedte)
        taak.GatParameter = If(taak.IsVierkant OrElse
            taak.Onderdeelnaam = "HoekInX", "d134", "d87")
    End Sub

    Private Sub BepaalOutputPad(taak As FlensTaak)
        Dim sleutel As String =
            taak.Onderdeelnaam & "_" & Bestandsnummer(taak.PartLengte) & "_" &
            Bestandsnummer(taak.Model.InputFlensHoogte) & "_" &
            Bestandsnummer(taak.Model.InputFlensBreedte) & "_" &
            Bestandsnummer(taak.Model.InputGatDia)
        taak.OutputPad = System.IO.Path.Combine(taak.UitvoerMap, sleutel & ".ipt")
    End Sub

    Private Function VindOccurrence(taak As FlensTaak) As Boolean
        For Each occ As ComponentOccurrence In taak.Asm.ComponentDefinition.Occurrences
            If occ.Name = taak.Onderdeelnaam OrElse
               occ.Name.StartsWith(taak.Onderdeelnaam & ":") Then
                taak.OccNaam = occ.Name
                Return True
            End If
        Next
        Return False
    End Function

    'Nieuw bestand: alleen kopiëren, GEEN checkout (staat nog niet in Vault).
    'Bestaand bestand: eerst checkout.
    Private Function MaakOfCheckOut(taak As FlensTaak) As Boolean
        If Not System.IO.File.Exists(taak.OutputPad) Then
            System.IO.File.Copy(taak.TemplatePad, taak.OutputPad, False)
            FlensGlobals.LogInfo("Nieuw IPT aangemaakt: " & taak.OutputPad)
            Return True
        End If
        If Not CheckOutFile(taak.OutputPad) Then
            FlensGlobals.Show("Checkout mislukt: " & taak.OutputPad)
            Return False
        End If
        Return True
    End Function

    Private Sub ConfigurePart(taak As FlensTaak)
        Dim partDoc As PartDocument = Nothing
        Try
            partDoc = CType(
                FlensGlobals.ThisApp.Documents.Open(taak.OutputPad, False), PartDocument)
            StelParametersIn(partDoc, taak)
            ControleerReadOnly(taak.OutputPad)
            partDoc.Update()
            partDoc.Save()
            FlensGlobals.LogInfo("Opgeslagen: " & taak.OutputPad)
            partDoc.Close(True)
        Catch
            If partDoc IsNot Nothing Then
                Try : partDoc.Close(False) : Catch : End Try
            End If
            Throw
        End Try
    End Sub

    Private Sub StelParametersIn(partDoc As PartDocument, taak As FlensTaak)
        Dim p As Parameters = partDoc.ComponentDefinition.Parameters
        Dim m As Input = taak.Model
        SetMm(p, New ParamWaarde("LengteInw", taak.PartLengte))
        SetMm(p, New ParamWaarde(taak.GatParameter, m.InputGatDia))
        SetMm(p, New ParamWaarde("d23", m.InputFlensHoogte))
        SetMm(p, New ParamWaarde("d24", m.InputFlensBreedte))
    End Sub

    'Gebruikt MmTekst (huidige cultuur): 12,5 mm op NL-systemen
    Private Sub SetMm(partParams As Parameters, pv As ParamWaarde)
        partParams.Item(pv.Naam).Expression = MmTekst(pv.Waarde)
    End Sub

    Private Sub ControleerReadOnly(bestand As String)
        If Not System.IO.File.Exists(bestand) Then Return
        Dim attrs As System.IO.FileAttributes = System.IO.File.GetAttributes(bestand)
        If (attrs And System.IO.FileAttributes.ReadOnly) <> 0 Then
            Throw New System.IO.IOException("Het IPT-bestand is read-only: " & bestand)
        End If
    End Sub
End Module