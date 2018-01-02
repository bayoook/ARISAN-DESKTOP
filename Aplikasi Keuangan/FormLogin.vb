Imports System.Data.SqlClient
Public Class FormLogin

    'Enkripsi password
    Public Function AES_Encrypt(ByVal input As String, ByVal pass As String) As String
        Dim AES As New System.Security.Cryptography.RijndaelManaged
        Dim Hash_AES As New System.Security.Cryptography.MD5CryptoServiceProvider
        Dim encrypted As String = ""
        Try
            Dim hash(31) As Byte
            Dim temp As Byte() = Hash_AES.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(pass))
            Array.Copy(temp, 0, hash, 0, 16)
            Array.Copy(temp, 0, hash, 15, 16)
            AES.Key = hash
            AES.Mode = Security.Cryptography.CipherMode.ECB
            Dim DESEncrypter As System.Security.Cryptography.ICryptoTransform = AES.CreateEncryptor
            Dim Buffer As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(input)
            encrypted = Convert.ToBase64String(DESEncrypter.TransformFinalBlock(Buffer, 0, Buffer.Length))
            Return encrypted
        Catch ex As Exception
        End Try
    End Function

    'ketika form diaktifkan
    Private Sub FormLogin_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        TextBox_Username.Focus()
        TextBox_Password.Text = ""
        FormRegister.Close()
    End Sub

    'button X click
    Private Sub Button_X_Click(sender As Object, e As EventArgs) Handles Button_X.Click
        Application.Exit()
    End Sub

    'button _ click
    Private Sub Button___Click(sender As Object, e As EventArgs) Handles Button__.Click
        Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
    End Sub

    'button login click
    Private Sub Button_Login_Click(sender As Object, e As EventArgs) Handles Button_Login.Click
        If TextBox_Username.Text = "" Or TextBox_Password.Text = "" Then
            MsgBox("Data Login Belum Lengkap", MsgBoxStyle.Information, "Login Failed")
            TextBox_Username.Focus()
            Exit Sub
        Else
            Call Koneksi()
            'memilih database ketika username=textbox1 dan password=textbox2
            CMD = New SqlCommand("select * from TBL_ACC where Username='" & TextBox_Username.Text & "' and Password='" & AES_Encrypt(TextBox_Password.Text, TextBox_Password.Text) & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                Dim username As String = RD.Item("Username")
                'ketika data tersebut tidak ada di database / tidak cocok
                If username = TextBox_Username.Text And RD.Item("Kode") = "ADMIN" Then
                    FormAdmin.id = RD.Item("Id")
                    FormAdmin.user = RD.Item("Username")
                    FormAdmin.path = RD.Item("Path")
                    FormAdmin.mail = RD.Item("Email")
                    Me.Hide()
                    FormAdmin.ShowDialog()
                    Me.Show()
                ElseIf username = TextBox_Username.Text And RD.Item("Kode") = "USER" Then 'ketika data tersebut cocok dan sebagai user
                    FormUser.id = RD.Item("Id")
                    ShowPict.id = RD.Item("Id")
                    FormUser.user = RD.Item("Username")
                    FormUser.path = RD.Item("Path")
                    FormUser.mail = RD.Item("Email")
                    Me.Hide()
                    FormUser.ShowDialog()
                    Me.Show()
                Else
                    MsgBox("Username atau Password yang anda masukkan salah", MsgBoxStyle.Critical, "Login Failed")
                    TextBox_Username.Focus()
                End If
            Else
                MsgBox("Username atau Password yang anda masukkan salah", MsgBoxStyle.Critical, "Login Failed")
                TextBox_Username.Focus()
            End If
        End If
    End Sub

    'button LupaPassword click
    Private Sub Button_LupaPassword_Click(sender As Object, e As EventArgs) Handles Button_LupaPassword.Click
        Me.Hide()
        FormLupaPass.ShowDialog()
        Me.Show()
    End Sub

    'button register click
    Private Sub Button_Register_Click(sender As Object, e As EventArgs) Handles Button_Register.Click
        Me.Hide()
        FormRegister.ShowDialog()
        Me.Show()
    End Sub

    'enter setiap textbox
    Private Sub TextBox_Username_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox_Username.KeyPress
        If e.KeyChar = Chr(13) Then TextBox_Password.Focus()
    End Sub
    Private Sub TextBox_Password_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox_Password.KeyPress
        If e.KeyChar = Chr(13) Then Button_Login.PerformClick()
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