Imports System.Data.SqlClient
Public Class ShowPict
    Public id As String
    Public path As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.ShowDialog()
        If Not OpenFileDialog1.FileName = "" Then
            Dim edit As String = "update TBL_ACC set Path='" & OpenFileDialog1.FileName & "' where Id='" & id & "'"
            CMD = New SqlCommand(edit, CONN)
            CMD.ExecuteNonQuery()
            PictureBox1.ImageLocation = OpenFileDialog1.FileName
            FormAdmin.path = OpenFileDialog1.FileName
            FormUser.path = OpenFileDialog1.FileName
            Button2.Text = "Ok"
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub ShowPict_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        PictureBox1.ImageLocation = path
        Button2.Text = "Exit"
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim edit As String = "update TBL_ACC set Path='' where Id='" & id & "'"
        CMD = New SqlCommand(edit, CONN)
        CMD.ExecuteNonQuery()
        FormAdmin.path = ""
        FormUser.path = ""
        PictureBox1.ImageLocation = ""
        PictureBox1.Image = Aplikasi_Keuangan.My.Resources.Resources.blank_user
        Button2.Text = "Ok"
    End Sub
End Class