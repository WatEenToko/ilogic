Imports Inventor

'veel toegepasde functies
Public Module FlensGlobals
    Public ThisApp As Inventor.Application
    Public LogInfo As Action(Of String)
    Public LogError As Action(Of String)

    Public Sub InitLogger(info As Action(Of String), errorHandler As Action(Of String))
        LogInfo = info
        LogError = errorHandler
    End Sub

    Public Sub Show(message As String)
        System.Windows.Forms.MessageBox.Show(message)
    End Sub
End Module
