Imports Inventor

'Vaak gebuikte berichten
Public Module FlensGlobals
    Public ThisApp As Inventor.Application
    Public LogInfo As Action(Of String)
    Public LogError As Action(Of String)

    Public Sub Init(app As Inventor.Application, info As Action(Of String), err As Action(Of String))
        ThisApp = app : LogInfo = info : LogError = err
    End Sub

    Public Sub Show(bericht As String)
        System.Windows.Forms.MessageBox.Show(bericht)
    End Sub
End Module
