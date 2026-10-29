Imports Inventor

Public Class Input
    Public Property InputLengte As Double
    Public Property InputBreedte As Double
    Public Property InputFlensHoogte As Double
    Public Property InputFlensBreedte As Double
    Public Property InputGatDia As Double

    Public ReadOnly Property IsValid As Boolean
        Get
            Return InputLengte > 0 AndAlso InputBreedte > 0 AndAlso
                   InputFlensHoogte > 0 AndAlso InputFlensBreedte > 0 AndAlso
                   InputGatDia > 0
        End Get
    End Property

    Public ReadOnly Property IsVierkant As Boolean
        Get
            Return InputLengte = InputBreedte
        End Get
    End Property
End Class

Public Class FlensTaak
    Public Asm As AssemblyDocument
    Public Model As Input
    Public IsVierkant As Boolean
    Public UitvoerMap As String
    Public Onderdeelnaam As String
    Public TemplatePad As String
    Public OccNaam As String
    Public OutputPad As String
    Public PartLengte As Double
    Public GatParameter As String

    Public Sub Start(naam As String, template As String)
        Onderdeelnaam = naam
        TemplatePad = template
    End Sub
End Class

Public Class ParamWaarde
    Public Naam As String
    Public Waarde As Double

    Public Sub New(naam As String, waarde As Double)
        Me.Naam = naam
        Me.Waarde = waarde
    End Sub
End Class

Public Module NummerHelper
    'Invariante cultuur: voor bestandsnamen (punt als decimaal)
    Public Function Bestandsnummer(value As Double) As String
        Return value.ToString(System.Globalization.CultureInfo.InvariantCulture)
    End Function

    'Huidige cultuur: voor Inventor-parameterexpressies (NL = komma)
    Public Function MmTekst(value As Double) As String
        Return value.ToString(System.Globalization.CultureInfo.CurrentCulture) & " mm"
    End Function
End Module