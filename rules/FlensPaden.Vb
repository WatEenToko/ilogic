'paden van diverse bestanden 
Public Class FlensPaden
    'x/y en output locaties
    Public ReadOnly TemplateX As String =
        "C:\Vault workspace\Inventor\CONCEPT\Flenzen\4losse\Kant2.ipt"
    Public ReadOnly TemplateY As String =
        "C:\Vault workspace\Inventor\CONCEPT\Flenzen\4losse\Kant1.ipt"
    Public ReadOnly OutputDir As String =
        "C:\Vault Workspace\Inventor\CONCEPT\Flenzen\tests"

    'input checken en templates doorsturen
    Public Function Validate(model As Input) As Boolean
        If model.IsVierkant Then
            If Not System.IO.File.Exists(TemplateX) Then
                FlensGlobals.Show("X-template ontbreekt: " & TemplateX)
                Return False
            End If
        Else
            If Not System.IO.File.Exists(TemplateX) OrElse
               Not System.IO.File.Exists(TemplateY) Then
                FlensGlobals.Show("Een of beide IPT-templatebestanden ontbreken.")
                Return False
            End If
        End If
        If Not System.IO.Directory.Exists(OutputDir) Then
            FlensGlobals.Show("Uitvoermap bestaat niet: " & OutputDir)
            Return False
        End If
        Return True
    End Function
End Class
