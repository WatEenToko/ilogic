Imports Inventor

Public Class Input
    Public Property InputLength As Double
    Public Property InputWidth As Double
    Public Property InputFlangeHeight As Double
    Public Property InputFlangeWidth As Double
    Public Property InputHoleDiameter As Double

    Public ReadOnly Property IsValid As Boolean
        Get
            Return InputLength > 0 AndAlso InputWidth > 0 AndAlso
                   InputFlangeHeight > 0 AndAlso InputFlangeWidth > 0 AndAlso
                   InputHoleDiameter > 0
        End Get
    End Property

    Public ReadOnly Property IsSquare As Boolean
        Get
            Return InputLength = InputWidth
        End Get
    End Property
End Class

Public Class FlangeTask
    Public Asm As AssemblyDocument
    Public Model As Input
    Public IsSquare As Boolean
    Public OutputFolder As String
    Public PartName As String
    Public TemplatePath As String
    Public OccurrenceName As String
    Public OutputPath As String
    Public PartLength As Double
    Public HoleParameter As String

    Public Sub Start(name As String, template As String)
        PartName = name
        TemplatePath = template
    End Sub
End Class

Public Class ParameterValue
    Public Name As String
    Public Value As Double

    Public Sub New(name As String, value As Double)
        Me.Name = name
        Me.Value = value
    End Sub
End Class

Public Module NumberHelper
    'Invariant culture: used for file names (dot as decimal separator)
    Public Function FileNumber(value As Double) As String
        Return value.ToString(System.Globalization.CultureInfo.InvariantCulture)
    End Function

    'Current culture: used for Inventor parameter expressions (NL = comma)
    Public Function MmText(value As Double) As String
        Return value.ToString(System.Globalization.CultureInfo.CurrentCulture) & " mm"
    End Function
End Module
