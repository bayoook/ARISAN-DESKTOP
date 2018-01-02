Imports System.Data.SqlClient
Module ModuleArisan
    Public CONN As SqlConnection
    Public CMD As SqlCommand
    Public DS As New DataSet
    Public DA As SqlDataAdapter
    Public RD As SqlDataReader
    Public LokasiData As String
    Public Sub Koneksi()
        LokasiData = "data source=RIZQIMUBAYOOOK;initial catalog=DBARISAN; integrated security=true"
        CONN = New SqlConnection(LokasiData)
        If CONN.State = ConnectionState.Closed Then
            CONN.Open()
        Else
            MsgBox("Gagal Connect")
        End If
    End Sub
End Module
