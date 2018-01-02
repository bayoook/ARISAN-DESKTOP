Public Class FormPinAdmin
    Public sukses As Integer
    Private Sub FormPinAdmin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox_Pin.Clear()
        TextBox_Pin.PasswordChar = "*"
        TextBox_Pin.Focus()
    End Sub
    Private Sub TextBox_Pin_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox_Pin.KeyPress
        If e.KeyChar = Chr(13) Then Button_Ok.PerformClick()
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub
    Private Sub Button_Ok_Click(sender As Object, e As EventArgs) Handles Button_Ok.Click
        If TextBox_Pin.Text = "326584" Then
            sukses = 1
        Else
            sukses = 0
        End If
        FormRegister.sukses = sukses
        Me.Close()
    End Sub
    Private Sub Button_X_Click(sender As Object, e As EventArgs) Handles Button_X.Click
        Me.Close()
    End Sub
    Private Sub Button___Click(sender As Object, e As EventArgs) Handles Button__.Click
        Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
    End Sub
    'agar form bisa di drag
    Public drag As Boolean = False
    Public xM As Integer
    Public yM As Integer
    Private Sub Form_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        If e.Button = MouseButtons.Left Then
            drag = True
            xM = e.X
            yM = e.Y
        End If
    End Sub
    Private Sub Form_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
        If e.Button = MouseButtons.Left Then
            drag = False
        End If
    End Sub
    Private Sub Form_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        If drag = True Then
            Dim tmp As Point = New Point()
            tmp.X = Me.Location.X + (e.X - xM)
            tmp.Y = Me.Location.Y + (e.Y - yM)
            Me.Location = tmp
            tmp = Nothing
        End If
    End Sub
End Class