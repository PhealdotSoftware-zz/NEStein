Public Class FrmRender
    Private Sub FrmRender_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        With Me
            .Width = Screen.PrimaryScreen.Bounds.Width
            .Height = Screen.PrimaryScreen.Bounds.Height
            .Top = 0
            .Left = 0
        End With

        With NesScreen
            .Width = (Me.Width \ 256) * 256
            .Height = (Me.Height \ 240) * 240
            .Top = (Me.Height / 2) - (.Height / 2)
            .Left = (Me.Width / 2) - (.Width / 2)
        End With
    End Sub
    Private Sub FrmRender_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        KeyCodes(e.KeyCode) = &H41
        If e.KeyCode = Keys.Escape Then
            FullScreen = False
            Me.Close()
            FrmMain.Show()
        End If
    End Sub
    Private Sub FrmRender_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        KeyCodes(e.KeyCode) = &H40
    End Sub
End Class