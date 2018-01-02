Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Public Class FormRegister

    Public sukses As Integer
    'Cek validasi string email
    Public Function validateEmail(emailAddress) As Boolean
        ' Dim email As New Regex("^(?<user>[^@]+)@(?<host>.+)$")
        Dim email As New Regex("([\w-+]+(?:\.[\w-+]+)*@(?:[\w-]+\.)+[a-zA-Z]{2,7})")
        If email.IsMatch(emailAddress) Then
            Return True
        Else
            Return False
        End If
    End Function

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

    'sub untuk menampilkan combobox1
    Sub TampilLevel()
        ComboBox_Kode.Items.Clear()
        ComboBox_Kode.Items.Add("USER")
        ComboBox_Kode.Items.Add("ADMIN")
        ComboBox_Kode.SelectedItem = "USER"
    End Sub

    'Ketika form aktif
    Private Sub FormRegister_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        FormLupaPass.Close()
        TextBox_Id.Focus()
    End Sub
    Private Sub FormRegister_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TampilLevel()
        FormLupaPass.Close()
        TextBox_Id.Focus()
    End Sub

    'Enter setiap textbox
    Private Sub TextBox_Id_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox_Id.KeyPress
        If e.KeyChar = Chr(13) Then TextBox_Username.Focus()
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub
    Private Sub TextBox_Username_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox_Username.KeyPress
        If e.KeyChar = Chr(13) Then TextBox_Email.Focus()
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) = 32 Then
                MsgBox("Username tidak boleh memakai spasi")
                e.KeyChar = Chr(0)
            End If
        End If
    End Sub
    Private Sub TextBox_Email_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox_Email.KeyPress
        If e.KeyChar = Chr(13) Then TextBox_Password.Focus()
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) = 32 Then
                MsgBox("Email tidak boleh memakai spasi")
                e.KeyChar = Chr(0)
            End If
        End If
    End Sub
    Private Sub TextBox_Password_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox_Password.KeyPress
        If e.KeyChar = Chr(13) Then Button_Daftar.PerformClick()
    End Sub

    'button daftar klik
    Private Sub Button_Daftar_Click(sender As Object, e As EventArgs) Handles Button_Daftar.Click
        If TextBox_Id.Text = "" Or TextBox_Username.Text = "" Or TextBox_Email.Text = "" Or TextBox_Password.Text = "" Then
            MsgBox("Data pendaftaran belum lengkap")
            TextBox_Id.Focus()
            Exit Sub
        ElseIf validateEmail(TextBox_Email.Text) = False Then
            MsgBox("Isi email dengan format yang benar")
            TextBox_Email.Focus()
        ElseIf TextBox_Username.TextLength < 6 Then
            MsgBox("Username minimal 6 karakter")
            TextBox_Username.Focus()
        ElseIf TextBox_Password.Textlength < 6 Then
            MsgBox("Password minimal 6 karakter")
        Else
            'kalau combobox1 terisi sebagai admin harus mengisi pin
            If ComboBox_Kode.Text = "ADMIN" Then
                FormPinAdmin.ShowDialog()
                'jika pin yang dimasukkan salah
                If sukses = 0 Then
                    MsgBox("Pin yang anda masukkan salah")
                ElseIf sukses = 1 Then 'jika pin yang dimasukkan salah
                    Call Koneksi()
                    'memeriksa data di database ketika username=textbox2 atau email=textbox3 atau id=textbox1
                    CMD = New SqlCommand("select * from TBL_ACC where Username='" & TextBox_Username.Text & "' or Email='" & TextBox_Email.Text & "' or Id='" & TextBox_Id.Text & "'", CONN)
                    RD = CMD.ExecuteReader
                    Dim simpan As String = "insert into TBL_ACC values ('" &
                        TextBox_Id.Text & "','" &
                        ComboBox_Kode.Text & "','" &
                        TextBox_Username.Text & "','" &
                        AES_Encrypt(TextBox_Password.Text, TextBox_Password.Text) & "','" &
                        TextBox_Email.Text & "','" &
                        "0','0','0','0','0','0','0','0','0','0','0','0','0','')"
                    RD.Read()
                    'jika data tersebut sudah ada
                    If RD.HasRows Then
                        MsgBox("Username atau Email atau Id sudah terdaftar")
                    Else 'jika data tersebut belum ada
                        RD.Close()
                        CMD = New SqlCommand(simpan, CONN)
                        CMD.ExecuteNonQuery()
                        MsgBox("Anda Berhasil Register", MsgBoxStyle.Information, "Information")
                        Me.Close()
                    End If
                End If
            Else 'jika yang dipilih adalah user
                Call Koneksi()
                'memeriksa data di database ketika username=textbox2 atau email=textbox4 atau id=textbox1
                CMD = New SqlCommand("select * from TBL_ACC where Username='" & TextBox_Username.Text & "' or Email='" & TextBox_Email.Text & "' or Id='" & TextBox_Id.Text & "'", CONN)
                RD = CMD.ExecuteReader
                Dim simpan As String = "insert into TBL_ACC values ('" &
                        TextBox_Id.Text & "','" &
                        ComboBox_Kode.Text & "','" &
                        TextBox_Username.Text & "','" &
                        AES_Encrypt(TextBox_Password.Text, TextBox_Password.Text) & "','" &
                        TextBox_Email.Text & "','" &
                        "0','0','0','0','0','0','0','0','0','0','0','0','0','')"
                RD.Read()
                'jika data tersebut sudah ada
                If RD.HasRows Then
                    MsgBox("Username atau Email atau Id sudah terdaftar")
                Else 'jika data tersebut belum ada
                    RD.Close()
                    CMD = New SqlCommand(simpan, CONN)
                    CMD.ExecuteNonQuery()
                    MsgBox("Anda Berhasil Register", MsgBoxStyle.Information, "Information")
                    Me.Close()
                End If
            End If
        End If
    End Sub

    'button kembali klik
    Private Sub Button_Kembali_Click(sender As Object, e As EventArgs) Handles Button_Kembali.Click
        FormLogin.Show()
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