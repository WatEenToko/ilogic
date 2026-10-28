Imports Inventor
'inputs verwerken met een aantal standaard aantallen omzetten naar model
Public Class Input
    Public Property InputLengte As Double
    Public Property InputBreedte As Double
    Public Property InputFlensHoogte As Double = 50
    Public Property InputFlensBreedte As Double =50 
    Public Property InputGatDia As Double
    Public Property OutputPath as String

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

Public Module NummerHelper
    Public Function FileNumber(value As Double) As String
        Return value.ToString(
            System.Globalization.CultureInfo.InvariantCulture)
    End Function
End Module
