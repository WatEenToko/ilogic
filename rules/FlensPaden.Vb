'Paths for the various templates and output files
Public Class FlensPaden
    'X/Y and output locations
    Public ReadOnly TemplateX As String =
        "C:\Vault workspace\Inventor\CONCEPT\Flenzen\4losse\Kant2.ipt"
    Public ReadOnly TemplateY As String =
        "C:\Vault workspace\Inventor\CONCEPT\Flenzen\4losse\Kant1.ipt"
    Public ReadOnly OutputDir As String =
        "C:\Vault Workspace\Inventor\CONCEPT\Flenzen\tests"

    'Validate the input and pass the templates through
    Public Function Validate(model As Input) As Boolean
        If model.IsSquare Then
            If Not System.IO.File.Exists(TemplateX) Then
                FlensGlobals.Show("X-template is missing: " & TemplateX)
                Return False
            End If
        Else
            If Not System.IO.File.Exists(TemplateX) OrElse
               Not System.IO.File.Exists(TemplateY) Then
                FlensGlobals.Show("One or both IPT template files are missing.")
                Return False
            End If
        End If
        If Not System.IO.Directory.Exists(OutputDir) Then
            FlensGlobals.Show("Output folder does not exist: " & OutputDir)
            Return False
        End If
        Return True
    End Function
End Class
