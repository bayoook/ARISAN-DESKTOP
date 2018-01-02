Imports System.Data.SqlClient
Public Class FormLupaPass

    'Enkripsi untuk password
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

    'ketika form dipanggil
    Private Sub FormLupaPass_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FormRegister.Close()
        TextBox_Username.Focus()
        TextBox_NewPassword.PasswordChar = "*"
        TextBox_Username.Text = ""
        TextBox_Email.Text = ""
        TextBox_NewPassword.Text = ""
    End Sub
    Private Sub FormLupaPass_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        FormRegister.Close()
        TextBox_Username.Focus()
    End Sub

    'Enter setiap textbox
    Private Sub TextBox_Username_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox_Username.KeyPress
        If e.KeyChar = Chr(13) Then TextBox_Email.Focus()
    End Sub
    Private Sub TextBox_Email_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox_Email.KeyPress
        If e.KeyChar = Chr(13) Then TextBox_NewPassword.Focus()
    End Sub
    Private Sub TextBox_NewPassword_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox_NewPassword.KeyPress
        If e.KeyChar = Chr(13) Then Button_Ganti.Focus()
    End Sub

    'button ganti click
    Private Sub Button_Ganti_Click(sender As Object, e As EventArgs) Handles Button_Ganti.Click
        If TextBox_Username.Text = "" Or TextBox_Email.Text = "" Or TextBox_NewPassword.Text = "" Then
            MsgBox("Data belum lengkap")
            TextBox_Username.Focus()
        ElseIf TextBox_NewPassword.TextLength < 6 Then
            MsgBox("Password minimal 6 karakter")
            TextBox_NewPassword.Focus()
        Else
            Call Koneksi()
            'memeriksa data di database ketika username=textbox1 dan email=textbox2
            CMD = New SqlCommand("select * from TBL_ACC where username ='" & TextBox_Username.Text & "' and email='" & TextBox_Email.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            'ketika data tidak ada atau tidak cocok
            If Not RD.HasRows Then
                MsgBox("Username dan Email yang anda masukkan tidak cocok")
            Else 'ketika data cocok
                Dim Id As String = RD.Item("Id")
                RD.Close()
                'string edit diisikan dengan update untuk database
                Dim edit As String = "update TBL_ACC set Password='" & AES_Encrypt(TextBox_NewPassword.Text, TextBox_NewPassword.Text) & "' where Id='" & Id & "'"
                'mengeksekusi string edit di database
                CMD = New SqlCommand(edit, CONN)
                'mengeksekusi database
                CMD.ExecuteNonQuery()
                MsgBox("Anda berhasil mengganti password")
                Me.Close()
            End If
        End If
    End Sub

    'button kembali click
    Private Sub Button_Kembali_Click(sender As Object, e As EventArgs) Handles Button_Kembali.Click
        Me.Close()
    End Sub

    'button X click
    Private Sub Button_X_Click(sender As Object, e As EventArgs) Handles Button_X.Click
        Me.Close()
    End Sub

    'button _ click
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