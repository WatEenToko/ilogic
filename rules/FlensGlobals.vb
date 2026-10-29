Imports Inventor

Public Module FlensGlobals
    Public ThisApp As Inventor.Application
    Public LogInfo As Action(Of String)
    Public LogError As Action(Of String)

    Public Sub InitLogger(info As Action(Of String), fout As Action(Of String))
        LogInfo = info
        LogError = fout
    End Sub

    Public Sub Show(bericht As String)
        System.Windows.Forms.MessageBox.Show(bericht)
    End Sub
End Module