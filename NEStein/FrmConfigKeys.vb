Public Class FrmConfigKeys
    Dim CurController, CurKey As Integer
    Private Sub FrmConfigKeys_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        UpdateLabel()
    End Sub
    Private Sub FrmConfigKeys_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        lblPressedKey.Text = e.KeyCode.ToString
        If e.KeyCode = Keys.Escape Then Me.Close()
    End Sub
    Private Sub FrmConfigKeys_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        lblPressedKey.Text = "..."

        InputKeys(CurController, CurKey) = e.KeyCode

        If CurKey < 7 Then
            CurKey += 1
        Else
            If CurController = 0 Then
                CurKey = 0
                CurController = 1
            Else
                Me.Close()
            End If
        End If

        UpdateLabel()
    End Sub
    Private Sub UpdateLabel()
        Dim Str As String = Nothing

        Select Case CurKey
            Case 0 : Str = "A"
            Case 1 : Str = "B"
            Case 2 : Str = "Select"
            Case 3 : Str = "Start"
            Case 4 : Str = "Up"
            Case 5 : Str = "Down"
            Case 6 : Str = "Left"
            Case 7 : Str = "Right"
        End Select

        lblKeyToPress.Text = "Press key ""Controller " & CurController + 1 & " " & Str & """"
    End Sub
End Class