Imports System.Data.SqlClient
Imports System.Runtime.InteropServices
Public Class FormAdmin
#Region "Variabel"
    Shared random As New Random()
    Public id As String
    Public user As String
    Public mail As String
    Public path As String
    Public drag As Boolean = False
    Public xM As Integer
    Public yM As Integer
    Private Const BorderWidth As Integer = 6
    Private _resizeDir As ResizeDirection = ResizeDirection.None
#End Region
#Region "Sub Tampil"
    Sub TampilBulan() 'Sub yang berfungsi untuk mengisikan combobox1 dengan nama-nama bulan
        ComboBox_Bulan.Items.Clear()
        ComboBox_Bulan.Items.Add("Januari")
        ComboBox_Bulan.Items.Add("Februari")
        ComboBox_Bulan.Items.Add("Maret")
        ComboBox_Bulan.Items.Add("April")
        ComboBox_Bulan.Items.Add("Mei")
        ComboBox_Bulan.Items.Add("Juni")
        ComboBox_Bulan.Items.Add("Juli")
        ComboBox_Bulan.Items.Add("Agustus")
        ComboBox_Bulan.Items.Add("September")
        ComboBox_Bulan.Items.Add("Oktober")
        ComboBox_Bulan.Items.Add("November")
        ComboBox_Bulan.Items.Add("Desember")
        For i As Integer = 0 To 12
            If i = DateTimePicker1.Value.Month - 1 Then
                ComboBox_Bulan.SelectedIndex = i
            Else
                Continue For
            End If
        Next
        TextBox_JumlahUang.Focus()
    End Sub
    Sub TampilGrid() 'sub yang digunakan untuk menampilkan isi grid
        Call Koneksi()
        ComboBox_Bulan.Items.Clear()
        ComboBox_Nama.Items.Clear()
        'Mengambil data dari database
        DA = New SqlDataAdapter("select Username as Username,
                    Januari as Januari, Februari as Februari, 
                    Maret as Maret, 
                    April as April, 
                    Mei as Mei, 
                    Juni as Juni, 
                    Juli as Juli, 
                    Agustus as Agustus, 
                    September as September, 
                    Oktober as Oktober, 
                    November as November, 
                    Desember as Desember,
                    Menang as Saldo from TBL_ACC", CONN)
        DS = New DataSet
        DA.Fill(DS, "TBL_ACC")
        'Mengisi source dari data grid view dengan data set dari tabel TBL_ACC
        DataGridView_Admin.DataSource = DS.Tables("TBL_ACC")
        DataGridView_Admin.ReadOnly = True
        Dim i As Integer
        'Mengambil cumlah baris yang ada
        Dim kolom As Integer = DS.Tables("TBL_ACC").Rows.Count
        For i = 1 To 13
            'Mengecilkan jarak antar kolom
            DataGridView_Admin.Columns(i).Width = 60
        Next
        For i = 0 To kolom - 1
            'Menambahkan combobox2 dengan data pada tabel ke 0,i
            ComboBox_Nama.Items.Add(DataGridView_Admin.Item(0, i).Value)
        Next
        ComboBox_Nama.SelectedItem = DataGridView_Admin.Item(0, 0).Value
        TextBox_JumlahUang.Focus()
    End Sub
    Sub TampilList()
        Call Koneksi()
        If Not ComboBox_NamaTagihan.Text = "" Then
            CMD = New SqlCommand("select * from TBL_ACC where username='" & ComboBox_NamaTagihan.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            Dim kurang As Integer
            Dim lebih As Integer
            ListBox_Tagihan.Items.Clear()
            If Convert.ToInt32(RD.Item("Januari")) = 0 Then
                ListBox_Tagihan.Items.Add("Januari - Belum Dibayar")
            ElseIf Convert.ToInt32(RD.Item("Januari")) < 1000000 Then
                kurang = 1000000 - Convert.ToInt32(RD.Item("Januari"))
                ListBox_Tagihan.Items.Add("Januari - Belum Lunas, Kurang Rp " + Convert.ToString(kurang))
            ElseIf Convert.ToInt32(RD.Item("Januari")) > 1000000 Then
                lebih = Convert.ToInt32(RD.Item("Januari")) - 1000000
                ListBox_Tagihan.Items.Add("Januari - Lunas, Lebih Rp " + Convert.ToString(lebih))
            Else
                ListBox_Tagihan.Items.Add("Januari - Lunas")
            End If
            If Convert.ToInt32(RD.Item("Februari")) = 0 Then
                ListBox_Tagihan.Items.Add("Februari - Belum Dibayar")
            ElseIf Convert.ToInt32(RD.Item("Februari")) < 1000000 Then
                kurang = 1000000 - Convert.ToInt32(RD.Item("Februari"))
                ListBox_Tagihan.Items.Add("Februari - Belum Lunas, Kurang Rp " + Convert.ToString(kurang))
            ElseIf Convert.ToInt32(RD.Item("Februari")) > 1000000 Then
                lebih = Convert.ToInt32(RD.Item("Februari")) - 1000000
                ListBox_Tagihan.Items.Add("Februari - Lunas, Lebih Rp " + Convert.ToString(lebih))
            Else
                ListBox_Tagihan.Items.Add("Februari - Lunas")
            End If
            If Convert.ToInt32(RD.Item("Maret")) = 0 Then
                ListBox_Tagihan.Items.Add("Maret - Belum Dibayar")
            ElseIf Convert.ToInt32(RD.Item("Maret")) < 1000000 Then
                kurang = 1000000 - Convert.ToInt32(RD.Item("Maret"))
                ListBox_Tagihan.Items.Add("Maret - Belum Lunas, Kurang Rp " + Convert.ToString(kurang))
            ElseIf Convert.ToInt32(RD.Item("Maret")) > 1000000 Then
                lebih = Convert.ToInt32(RD.Item("Maret")) - 1000000
                ListBox_Tagihan.Items.Add("Maret - Lunas, Lebih Rp " + Convert.ToString(lebih))
            Else
                ListBox_Tagihan.Items.Add("Maret - Lunas")
            End If
            If Convert.ToInt32(RD.Item("April")) = 0 Then
                ListBox_Tagihan.Items.Add("April - Belum Dibayar")
            ElseIf Convert.ToInt32(RD.Item("April")) < 1000000 Then
                kurang = 1000000 - Convert.ToInt32(RD.Item("April"))
                ListBox_Tagihan.Items.Add("April - Belum Lunas, Kurang Rp " + Convert.ToString(kurang))
            ElseIf Convert.ToInt32(RD.Item("April")) > 1000000 Then
                lebih = Convert.ToInt32(RD.Item("April")) - 1000000
                ListBox_Tagihan.Items.Add("April - Lunas, Lebih Rp " + Convert.ToString(lebih))
            Else
                ListBox_Tagihan.Items.Add("April - Lunas")
            End If
            If Convert.ToInt32(RD.Item("Mei")) = 0 Then
                ListBox_Tagihan.Items.Add("Mei - Belum Dibayar")
            ElseIf Convert.ToInt32(RD.Item("Mei")) < 1000000 Then
                kurang = 1000000 - Convert.ToInt32(RD.Item("Mei"))
                ListBox_Tagihan.Items.Add("Mei - Belum Lunas, Kurang Rp " + Convert.ToString(kurang))
            ElseIf Convert.ToInt32(RD.Item("Mei")) > 1000000 Then
                lebih = Convert.ToInt32(RD.Item("Mei")) - 1000000
                ListBox_Tagihan.Items.Add("Mei - Lunas, Lebih Rp " + Convert.ToString(lebih))
            Else
                ListBox_Tagihan.Items.Add("Mei - Lunas")
            End If
            If Convert.ToInt32(RD.Item("Juni")) = 0 Then
                ListBox_Tagihan.Items.Add("Juni - Belum Dibayar")
            ElseIf Convert.ToInt32(RD.Item("Juni")) < 1000000 Then
                kurang = 1000000 - Convert.ToInt32(RD.Item("Juni"))
                ListBox_Tagihan.Items.Add("Juni - Belum Lunas, Kurang Rp " + Convert.ToString(kurang))
            ElseIf Convert.ToInt32(RD.Item("Juni")) > 1000000 Then
                lebih = Convert.ToInt32(RD.Item("Juni")) - 1000000
                ListBox_Tagihan.Items.Add("Juni - Lunas, Lebih Rp " + Convert.ToString(lebih))
            Else
                ListBox_Tagihan.Items.Add("Juni - Lunas")
            End If
            If Convert.ToInt32(RD.Item("Juli")) = 0 Then
                ListBox_Tagihan.Items.Add("Juli - Belum Dibayar")
            ElseIf Convert.ToInt32(RD.Item("Juli")) < 1000000 Then
                kurang = 1000000 - Convert.ToInt32(RD.Item("Juli"))
                ListBox_Tagihan.Items.Add("Juli - Belum Lunas, Kurang Rp " + Convert.ToString(kurang))
            ElseIf Convert.ToInt32(RD.Item("Juli")) > 1000000 Then
                lebih = Convert.ToInt32(RD.Item("Juli")) - 1000000
                ListBox_Tagihan.Items.Add("Juli - Lunas, Lebih Rp " + Convert.ToString(lebih))
            Else
                ListBox_Tagihan.Items.Add("Juli - Lunas")
            End If
            If Convert.ToInt32(RD.Item("Agustus")) = 0 Then
                ListBox_Tagihan.Items.Add("Agustus - Belum Dibayar")
            ElseIf Convert.ToInt32(RD.Item("Agustus")) < 1000000 Then
                kurang = 1000000 - Convert.ToInt32(RD.Item("Agustus"))
                ListBox_Tagihan.Items.Add("Agustus - Belum Lunas, Kurang Rp " + Convert.ToString(kurang))
            ElseIf Convert.ToInt32(RD.Item("Agustus")) > 1000000 Then
                lebih = Convert.ToInt32(RD.Item("Agustus")) - 1000000
                ListBox_Tagihan.Items.Add("Agustus - Lunas, Lebih Rp " + Convert.ToString(lebih))
            Else
                ListBox_Tagihan.Items.Add("Agustus - Lunas")
            End If
            If Convert.ToInt32(RD.Item("September")) = 0 Then
                ListBox_Tagihan.Items.Add("September - Belum Dibayar")
            ElseIf Convert.ToInt32(RD.Item("September")) < 1000000 Then
                kurang = 1000000 - Convert.ToInt32(RD.Item("September"))
                ListBox_Tagihan.Items.Add("September - Belum Lunas, Kurang Rp " + Convert.ToString(kurang))
            ElseIf Convert.ToInt32(RD.Item("September")) > 1000000 Then
                lebih = Convert.ToInt32(RD.Item("September")) - 1000000
                ListBox_Tagihan.Items.Add("September - Lunas, Lebih Rp " + Convert.ToString(lebih))
            Else
                ListBox_Tagihan.Items.Add("September - Lunas")
            End If
            If Convert.ToInt32(RD.Item("Oktober")) = 0 Then
                ListBox_Tagihan.Items.Add("Oktober - Belum Dibayar")
            ElseIf Convert.ToInt32(RD.Item("Oktober")) < 1000000 Then
                kurang = 1000000 - Convert.ToInt32(RD.Item("Oktober"))
                ListBox_Tagihan.Items.Add("Oktober - Belum Lunas, Kurang Rp " + Convert.ToString(kurang))
            ElseIf Convert.ToInt32(RD.Item("Oktober")) > 1000000 Then
                lebih = Convert.ToInt32(RD.Item("Oktober")) - 1000000
                ListBox_Tagihan.Items.Add("Oktober - Lunas, Lebih Rp " + Convert.ToString(lebih))
            Else
                ListBox_Tagihan.Items.Add("Oktober - Lunas")
            End If
            If Convert.ToInt32(RD.Item("November")) = 0 Then
                ListBox_Tagihan.Items.Add("November - Belum Dibayar")
            ElseIf Convert.ToInt32(RD.Item("November")) < 1000000 Then
                kurang = 1000000 - Convert.ToInt32(RD.Item("November"))
                ListBox_Tagihan.Items.Add("November - Belum Lunas, Kurang Rp " + Convert.ToString(kurang))
            ElseIf Convert.ToInt32(RD.Item("November")) > 1000000 Then
                lebih = Convert.ToInt32(RD.Item("November")) - 1000000
                ListBox_Tagihan.Items.Add("November - Lunas, Lebih Rp " + Convert.ToString(lebih))
            Else
                ListBox_Tagihan.Items.Add("November - Lunas")
            End If
            If Convert.ToInt32(RD.Item("Desember")) = 0 Then
                ListBox_Tagihan.Items.Add("Desember - Belum Dibayar")
            ElseIf Convert.ToInt32(RD.Item("Desember")) < 1000000 Then
                kurang = 1000000 - Convert.ToInt32(RD.Item("Desember"))
                ListBox_Tagihan.Items.Add("Desember - Belum Lunas, Kurang Rp " + Convert.ToString(kurang))
            ElseIf Convert.ToInt32(RD.Item("Desember")) > 1000000 Then
                lebih = Convert.ToInt32(RD.Item("Desember")) - 1000000
                ListBox_Tagihan.Items.Add("Desember - Lunas, Lebih Rp " + Convert.ToString(lebih))
            Else
                ListBox_Tagihan.Items.Add("Desember - Lunas")
            End If
            RD.Close()
        End If

    End Sub
    Sub tampilbulantagihan()
        ComboBox_BulanTagihan.Items.Clear()
        ComboBox_BulanTagihan.Items.Add("Januari")
        ComboBox_BulanTagihan.Items.Add("Februari")
        ComboBox_BulanTagihan.Items.Add("Maret")
        ComboBox_BulanTagihan.Items.Add("April")
        ComboBox_BulanTagihan.Items.Add("Mei")
        ComboBox_BulanTagihan.Items.Add("Juni")
        ComboBox_BulanTagihan.Items.Add("Juli")
        ComboBox_BulanTagihan.Items.Add("Agustus")
        ComboBox_BulanTagihan.Items.Add("September")
        ComboBox_BulanTagihan.Items.Add("Oktober")
        ComboBox_BulanTagihan.Items.Add("November")
        ComboBox_BulanTagihan.Items.Add("Desember")
    End Sub
    Sub tampilnamatagihan()
        ComboBox_NamaTagihan.Items.Clear()
        Dim kolom As Integer = DS.Tables("TBL_ACC").Rows.Count
        Dim i As Integer
        For i = 0 To kolom - 1
            ComboBox_NamaTagihan.Items.Add(DataGridView_Admin.Item(0, i).Value)
        Next
    End Sub
    Sub TampilBulanNew()
        Call Koneksi()
        If Not ComboBox_BulanTagihan.Text = "" Then
            Dim kolom As Integer = DS.Tables("TBL_ACC").Rows.Count
            Dim kurang As Integer
            Dim lebih As Integer
            ListBox_Tagihan.Items.Clear()
            For i = 0 To kolom - 1
                CMD = New SqlCommand("select * from TBL_ACC where username='" & DataGridView_Admin.Item(0, i).Value & "'", CONN)
                RD = CMD.ExecuteReader
                RD.Read()
                If Convert.ToInt32(RD.Item(ComboBox_BulanTagihan.Text)) = 0 Then
                    ListBox_Tagihan.Items.Add(DataGridView_Admin.Item(0, i).Value + " - Belum Dibayar")
                ElseIf Convert.ToInt32(RD.Item(ComboBox_BulanTagihan.Text)) < 1000000 Then
                    kurang = 1000000 - Convert.ToInt32(RD.Item(ComboBox_BulanTagihan.Text))
                    ListBox_Tagihan.Items.Add(DataGridView_Admin.Item(0, i).Value + " - Belum Lunas, Kurang Rp " + Convert.ToString(kurang))
                ElseIf Convert.ToInt32(RD.Item(ComboBox_BulanTagihan.Text)) > 1000000 Then
                    lebih = Convert.ToInt32(RD.Item(ComboBox_BulanTagihan.Text)) - 1000000
                    ListBox_Tagihan.Items.Add(DataGridView_Admin.Item(0, i).Value + "- Lunas, Lebih Rp " + Convert.ToString(lebih))
                Else
                    ListBox_Tagihan.Items.Add(DataGridView_Admin.Item(0, i).Value + "- Lunas")
                End If
                RD.Close()
            Next
        End If
    End Sub
#End Region
#Region "Button"
    'button input click
    Private Sub Button_Input_Click(sender As Object, e As EventArgs) Handles Button_Input.Click
        Dim jumlah As Integer
        If TextBox_JumlahUang.Text = "" Or ComboBox_Bulan.Text = "" Or ComboBox_Nama.Text = "" Then
            MsgBox("Data belum lengkap, mohon diisi semua ya bosq!!")
            Exit Sub
        Else
            Call Koneksi()
            CMD = New SqlCommand("select * from TBL_ACC where username='" & ComboBox_Nama.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            jumlah = RD.Item(ComboBox_Bulan.Text) + TextBox_JumlahUang.Text
            RD.Close()
            'string edit diisi dengan kata2 untuk menupdate database
            Dim edit As String = "update TBL_ACC set " & ComboBox_Bulan.Text & "='" & jumlah & "' where Username='" & ComboBox_Nama.Text & "'"
            'memasukkan string edit tadi untuk diproses oleh database
            CMD = New SqlCommand(edit, CONN)
            'mengexekusi database
            CMD.ExecuteNonQuery()
            Call TampilGrid()
            Call TampilBulan()
        End If
    End Sub
    'button delete click
    Private Sub Button_Delete_Click(sender As Object, e As EventArgs) Handles Button_Delete.Click
        'i sebagai indeks kolom yg dipilih
        Dim i As Integer = DataGridView_Admin.CurrentCell.ColumnIndex
        'j sebagai indeks baris yg dipilih
        Dim j As Integer = DataGridView_Admin.CurrentRow.Index
        'ketika i=0 (pada kolom username)
        If i = 0 Then
            'mengeluarkan message box bertipe yesno
            Select Case MsgBox("Apakah anda ingin menghapus akun '" & DataGridView_Admin.Item(i, j).Value & "' ?", MsgBoxStyle.YesNo)
                'ketika pilihannya yes
                Case MsgBoxResult.Yes
                    'string hapus diisi dengan perintah delete database
                    Dim hapus As String = "delete from TBL_ACC where username='" & DataGridView_Admin.Item(i, j).Value & "'"
                    'mengeksekusi string hapus
                    CMD = New SqlCommand(hapus, CONN)
                    'mengeksekusi database
                    CMD.ExecuteNonQuery()
                    MsgBox("Data berhasil di Hapus", MsgBoxStyle.Information, "Information")
                    Call TampilGrid()
                    Call TampilBulan()
            End Select
        ElseIf i = 13 Then 'ketika i=13 (kolom saldo)
            Call Koneksi()
            'string edit diisi dengan perintah update database
            Dim edit As String = "update TBL_ACC set Menang='0' where Username='" & ComboBox_Nama.Text & "'"
            'mengeksekusi string edit ke database
            CMD = New SqlCommand(edit, CONN)
            'mengeksekusi database
            CMD.ExecuteNonQuery()
            Call TampilGrid()
            Call TampilBulan()
        ElseIf TextBox_JumlahUang.Text = "" Or ComboBox_Bulan.Text = "" Or ComboBox_Nama.Text = "" Then
            MsgBox("Silahkan double klik data di bawah bosq!!")
            Exit Sub
        Else
            Call Koneksi()
            'string edit diisi dengan perintah update database
            Dim edit As String = "update TBL_ACC set " & ComboBox_Bulan.Text & "='0' where Username='" & ComboBox_Nama.Text & "'"
            'mengeksekusi string edit ke database
            CMD = New SqlCommand(edit, CONN)
            'mengeksekusi database
            CMD.ExecuteNonQuery()
            Call TampilGrid()
            Call TampilBulan()
        End If
    End Sub

    'button input di menu tagihan click
    Private Sub Button_InputTagihan_Click(sender As Object, e As EventArgs) Handles Button_InputTagihan.Click
        Dim jumlah As Integer
        If TextBox_JumlahUangNew.Text = "" Or TextBox_BulanNew.Text = "" Or TextBox_NamaNew.Text = "" Then
            MsgBox("Data belum lengkap, mohon diisi semua ya")
        Else
            Call Koneksi()
            CMD = New SqlCommand("select * from TBL_ACC where username='" & TextBox_NamaNew.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            jumlah = RD.Item(TextBox_BulanNew.Text) + TextBox_JumlahUangNew.Text
            RD.Close()
            'string edit diisi dengan kata2 untuk menupdate database
            Dim edit As String = "update TBL_ACC set " & TextBox_BulanNew.Text & "='" & jumlah & "' where Username='" & TextBox_NamaNew.Text & "'"
            'memasukkan string edit tadi untuk diproses oleh database
            CMD = New SqlCommand(edit, CONN)
            'mengexekusi database
            CMD.ExecuteNonQuery()
            Call TampilGrid()
            Call TampilBulan()
            Call TampilList()
            Call TampilBulanNew()
        End If
    End Sub
    'button delete di menu tagihan click
    Private Sub Button_DeleteTagihan_Click(sender As Object, e As EventArgs) Handles Button_DeleteTagihan.Click
        If TextBox_JumlahUangNew.Text = "" Or TextBox_BulanNew.Text = "" Or TextBox_NamaNew.Text = "" Then
            MsgBox("Silahkan klik data di samping")
            Exit Sub
        Else
            Call Koneksi()
            'string edit diisi dengan perintah update database
            Dim delete As String = "update TBL_ACC set " & TextBox_BulanNew.Text & "='0' where Username='" & TextBox_NamaNew.Text & "'"
            'mengeksekusi string edit ke database
            CMD = New SqlCommand(delete, CONN)
            'mengeksekusi database
            CMD.ExecuteNonQuery()
            Call TampilGrid()
            Call TampilBulan()
            If Label_Nama.Visible = True Then
                Call TampilList()
            Else
                Call TampilBulanNew()
            End If
        End If
    End Sub

    'ketika button undi diklik
    Private Sub Button_Undi_Click(sender As Object, e As EventArgs) Handles Button_Undi.Click
        PictureBox_Pemenang.Visible = True
        Panel1.Visible = True
        Panel2.Visible = True
        Label_EmailPemenang.Visible = True
        Label_UserPemenang.Visible = True
        Label_Selamat.Visible = True
        Label_Menang.Visible = False
        Dim max As Integer
        Dim i As Integer
        ProgressBar1.Value = 0
        i = 0
menang:
        i += 1
        Dim jumlah As Integer
        jumlah = DS.Tables("TBL_ACC").Rows.Count
        max = jumlah * 10
        Dim kode As Integer
        Dim hadiah As Integer
        Dim pemenang As String
        kode = random.Next(0, jumlah)
        pemenang = DataGridView_Admin.Item(0, kode).Value
        hadiah = 1000000 * jumlah
        Call Koneksi()
        CMD = New SqlCommand("select * from TBL_ACC where Username='" & pemenang & "'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.Item("Menang") = 0 Then
            RD.Close()
            Dim edit As String = "update TBL_ACC set Menang ='" & hadiah & "' where Username='" & pemenang & "'"
            CMD = New SqlCommand(edit, CONN)
            CMD.ExecuteNonQuery()
        Else
            If i = jumlah Then
                Select Case MsgBox("Semua sudah menang pada arisan ini, apakah anda ingin memulai dari awal?", MsgBoxStyle.YesNo)
                'ketika pilihannya yes
                    Case MsgBoxResult.Yes
                        Dim edit As String = "update TBL_ACC set Menang='0'"
                        'mengeksekusi string edit ke database
                        CMD = New SqlCommand(edit, CONN)
                        'mengeksekusi database
                        CMD.ExecuteNonQuery()
                        Call TampilGrid()
                        Call TampilBulan()
                        Panel_Undi.Visible = True
                End Select
            Else
                RD.Close()
                GoTo menang
            End If
        End If
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = max
        ProgressBar1.Step = 1
        For j = 1 To 10
            For i = 0 To jumlah - 1
                If ProgressBar1.Value = max Then
                    Exit For
                End If
                ProgressBar1.Value += 1
                Call Koneksi()
                CMD = New SqlCommand("select * from TBL_ACC where Username='" & DataGridView_Admin.Item(0, i).Value & "'", CONN)
                RD = CMD.ExecuteReader
                RD.Read()
                If RD.Item("Path") = "" Then
                    PictureBox_Pemenang.Image = Aplikasi_Keuangan.My.Resources.Resources.blank_user
                Else
                    PictureBox_Pemenang.ImageLocation = RD.Item("Path")
                End If
                Label_EmailPemenang.Text = RD.Item("Email")
                Label_UserPemenang.Text = RD.Item("Username")
                RD.Close()
                System.Threading.Thread.Sleep(100)
                Application.DoEvents()
            Next
        Next
        If ProgressBar1.Value = max Then
            Call Koneksi()
            CMD = New SqlCommand("select * from TBL_ACC where Username='" & pemenang & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.Item("Path") = "" Then
                PictureBox_Pemenang.Image = Aplikasi_Keuangan.My.Resources.Resources.blank_user
            Else
                PictureBox_Pemenang.ImageLocation = RD.Item("Path")
            End If
            Label_EmailPemenang.Text = RD.Item("Email")
            Label_UserPemenang.Text = RD.Item("Username")
            RD.Close()
            Label_Menang.Visible = True
        End If
        Call TampilGrid()
        Call TampilBulan()
    End Sub
#End Region
#Region "Data Click"
    'data gridview click
    Private Sub DataGridView_Admin_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView_Admin.CellClick
        Call Koneksi()
        Dim i As Integer
        'mendeklarasikan i sebagai indeks baris yang diklik
        i = DataGridView_Admin.CurrentRow.Index
        Dim j As Integer
        'mendeklarasikan j sebagai indeks kolom yang diklik
        j = DataGridView_Admin.CurrentCell.ColumnIndex
        If j = 0 Then
            TextBox_JumlahUang.Text = "0"
        Else
            TextBox_JumlahUang.Text = DataGridView_Admin.Item(j, i).Value
        End If
        ComboBox_Nama.SelectedItem = DataGridView_Admin.Item(0, i).Value
        'case untuk setiap j, untuk memasukkan nama bulan ke combobox1
        Select Case j
            Case 1
                ComboBox_Bulan.SelectedItem = "Januari"
            Case 2
                ComboBox_Bulan.SelectedItem = "Februari"
            Case 3
                ComboBox_Bulan.SelectedItem = "Maret"
            Case 4
                ComboBox_Bulan.SelectedItem = "April"
            Case 5
                ComboBox_Bulan.SelectedItem = "Mei"
            Case 6
                ComboBox_Bulan.SelectedItem = "Juni"
            Case 7
                ComboBox_Bulan.SelectedItem = "Juli"
            Case 8
                ComboBox_Bulan.SelectedItem = "Agustus"
            Case 9
                ComboBox_Bulan.SelectedItem = "September"
            Case 10
                ComboBox_Bulan.SelectedItem = "Oktober"
            Case 11
                ComboBox_Bulan.SelectedItem = "November"
            Case 12
                ComboBox_Bulan.SelectedItem = "Desember"
            Case Else
                ComboBox_Bulan.SelectedItem = ""
        End Select
        TextBox_JumlahUang.Focus()
    End Sub
    'listbox tagihan click
    Private Sub ListBox_Tagihan_Click(sender As Object, e As EventArgs) Handles ListBox_Tagihan.Click
        If Label_Nama.Visible = True Then
            TextBox_NamaNew.Text = ComboBox_NamaTagihan.Text
            ComboBox_Bulan.SelectedIndex = ListBox_Tagihan.SelectedIndex
            TextBox_BulanNew.Text = ComboBox_Bulan.Text
        ElseIf Label_Bulan.Visible = True Then
            TextBox_BulanNew.Text = ComboBox_BulanTagihan.Text
            ComboBox_Nama.SelectedIndex = ListBox_Tagihan.SelectedIndex
            TextBox_NamaNew.Text = ComboBox_Nama.Text
        End If
        If TextBox_NamaNew.Text = "" Or TextBox_BulanNew.Text = "" Then
            MsgBox("Silahkan Pilih Data")
        Else
            Call Koneksi()
            CMD = New SqlCommand("select * from TBL_ACC where username='" & TextBox_NamaNew.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            Dim jumlah As Integer = 1000000 - RD.Item(TextBox_BulanNew.Text)
            TextBox_JumlahUangNew.Text = jumlah
            RD.Close()
        End If
    End Sub
#End Region
#Region "Menu"
    'menu tabel arisan click
    Private Sub Button_TabelArisan_Click(sender As Object, e As EventArgs) Handles Button_TabelArisan.Click
        Panel_TabelArisan.Visible = True
        Panel_TagihanArisan.Visible = False
        Panel_Undi.Visible = False
    End Sub
    'menu tagihan arisan click
    Private Sub Button_TagihanArisan_Click(sender As Object, e As EventArgs) Handles Button_TagihanArisan.Click
        Panel_TabelArisan.Visible = False
        Panel_TagihanArisan.Visible = True
        Panel_Undi.Visible = False
    End Sub
    'menu perorangan click
    Private Sub Button_Perorangan_Click(sender As Object, e As EventArgs) Handles Button_Perorangan.Click
        Panel_Perorangan.Visible = True
        Label_Bulan.Visible = False
        Label_Nama.Visible = True
        tampilnamatagihan()
        ComboBox_BulanTagihan.Visible = False
        ComboBox_NamaTagihan.Visible = True
        ListBox_Tagihan.Items.Clear()
    End Sub
    'menu perbulan click
    Private Sub Button_Perbulan_Click(sender As Object, e As EventArgs) Handles Button_Perbulan.Click
        Panel_Perorangan.Visible = True
        Label_Bulan.Visible = True
        Label_Nama.Visible = False
        ComboBox_BulanTagihan.Visible = True
        ComboBox_NamaTagihan.Visible = False
        ListBox_Tagihan.Items.Clear()
        tampilbulantagihan()
        ComboBox_NamaTagihan.SelectedItem = ""
        ComboBox_BulanTagihan.SelectedItem = ""
    End Sub
    'menu undi click
    Private Sub Button_MenuUndi_Click(sender As Object, e As EventArgs) Handles Button_MenuUndi.Click
        Panel_Undi.Visible = True
        Panel_TagihanArisan.Visible = False
        Panel_TabelArisan.Visible = False
        PictureBox_Pemenang.Visible = False
        Panel1.Visible = False
        Panel2.Visible = False
        Label_EmailPemenang.Visible = False
        Label_UserPemenang.Visible = False
        Label_Selamat.Visible = False
    End Sub
#End Region
#Region "Others"
    Private Sub FormAdmin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox_JumlahUang.Focus()
        Call TampilGrid()
        Call TampilBulan()
        TextBox_JumlahUang.Text = ""
        Label_Username.Text = user
        Label_Email.Text = mail
        PictureBox_Admin.ImageLocation = path
        If path = "" Then
            PictureBox_Admin.Image = Aplikasi_Keuangan.My.Resources.Resources.blank_user
        End If
        Panel_TabelArisan.Visible = False
        Panel_TagihanArisan.Visible = False
        Panel_Undi.Visible = False
    End Sub

    'logout label click
    Private Sub Label_Logout_Click(sender As Object, e As EventArgs) Handles Label_Logout.Click
        Me.Close()
    End Sub
    'picturebox click
    Private Sub PictureBox_Admin_Click(sender As Object, e As EventArgs) Handles PictureBox_Admin.Click
        ShowPict.id = id
        ShowPict.path = path
        ShowPict.ShowDialog()
    End Sub

    Private Sub TextBox_JumlahUang_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox_JumlahUang.KeyPress
        If e.KeyChar = Chr(13) Then Button_Input.PerformClick()
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 42 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub
    Private Sub TextBox_JumlahUangNew_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox_JumlahUangNew.KeyPress
        If e.KeyChar = Chr(13) Then Button_InputTagihan.PerformClick()
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 42 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub
#End Region
#Region "Combobox"
    'pilihan pada combobox berubah
    Private Sub ComboBox_BulanTagihan_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_BulanTagihan.SelectedIndexChanged
        Call TampilBulanNew()
    End Sub
    Private Sub ComboBox_NamaTagihan_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_NamaTagihan.SelectedIndexChanged
        Call TampilList()
    End Sub
#End Region
#Region "Panel Atas"
    'panel kanan atas
    Private Sub Button___Click(sender As Object, e As EventArgs) Handles Button__.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub
    Private Sub Button_Max_Click(sender As Object, e As EventArgs) Handles Button_Max.Click
        If Me.WindowState = FormWindowState.Maximized Then
            Me.WindowState = FormWindowState.Normal
        Else
            Me.WindowState = FormWindowState.Maximized
        End If
    End Sub
    Private Sub Button_X_Click(sender As Object, e As EventArgs) Handles Button_X.Click
        Me.Close()
    End Sub
#End Region
#Region "Resize Form"
    'agar form bisa di resize
    Public Enum ResizeDirection
        None = 0
        Left = 1
        TopLeft = 2
        Top = 3
        TopRight = 4
        Right = 5
        BottomRight = 6
        Bottom = 7
        BottomLeft = 8
    End Enum
    Public Property resizeDir() As ResizeDirection
        Get
            Return _resizeDir
        End Get
        Set(ByVal value As ResizeDirection)
            _resizeDir = value
            'Change cursor
            Select Case value
                Case ResizeDirection.Left
                    Me.Cursor = Cursors.SizeWE
                Case ResizeDirection.Right
                    Me.Cursor = Cursors.SizeWE
                Case ResizeDirection.Top
                    Me.Cursor = Cursors.SizeNS
                Case ResizeDirection.Bottom
                    Me.Cursor = Cursors.SizeNS
                Case ResizeDirection.BottomLeft
                    Me.Cursor = Cursors.SizeNESW
                Case ResizeDirection.TopRight
                    Me.Cursor = Cursors.SizeNESW
                Case ResizeDirection.BottomRight
                    Me.Cursor = Cursors.SizeNWSE
                Case ResizeDirection.TopLeft
                    Me.Cursor = Cursors.SizeNWSE
                Case Else
                    Me.Cursor = Cursors.Default
            End Select
        End Set
    End Property
#Region " Functions and Constants "

    <DllImport("user32.dll")>
    Public Shared Function ReleaseCapture() As Boolean
    End Function

    <DllImport("user32.dll")>
    Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    End Function

    Private Const WM_NCLBUTTONDOWN As Integer = &HA1
    Private Const HTBORDER As Integer = 18
    Private Const HTBOTTOM As Integer = 15
    Private Const HTBOTTOMLEFT As Integer = 16
    Private Const HTBOTTOMRIGHT As Integer = 17
    Private Const HTCAPTION As Integer = 2
    Private Const HTLEFT As Integer = 10
    Private Const HTRIGHT As Integer = 11
    Private Const HTTOP As Integer = 12
    Private Const HTTOPLEFT As Integer = 13
    Private Const HTTOPRIGHT As Integer = 14

#End Region
#Region " Moving & Resizing methods "

    Private Sub MoveForm()
        ReleaseCapture()
        SendMessage(Me.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End Sub

    Private Sub ResizeForm(ByVal direction As ResizeDirection)
        Dim dir As Integer = -1
        Select Case direction
            Case ResizeDirection.Left
                dir = HTLEFT
            Case ResizeDirection.TopLeft
                dir = HTTOPLEFT
            Case ResizeDirection.Top
                dir = HTTOP
            Case ResizeDirection.TopRight
                dir = HTTOPRIGHT
            Case ResizeDirection.Right
                dir = HTRIGHT
            Case ResizeDirection.BottomRight
                dir = HTBOTTOMRIGHT
            Case ResizeDirection.Bottom
                dir = HTBOTTOM
            Case ResizeDirection.BottomLeft
                dir = HTBOTTOMLEFT
        End Select

        If dir <> -1 Then
            ReleaseCapture()
            SendMessage(Me.Handle, WM_NCLBUTTONDOWN, dir, 0)
        End If
    End Sub
#End Region
    Private Sub Form1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Left And Me.WindowState <> FormWindowState.Maximized Then
            ResizeForm(resizeDir)
        End If
    End Sub
    Private Sub Form1_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
        'Calculate which direction to resize based on mouse position

        If e.Location.X < BorderWidth And e.Location.Y < BorderWidth Then
            resizeDir = ResizeDirection.TopLeft

        ElseIf e.Location.X < BorderWidth And e.Location.Y > Me.Height - BorderWidth Then
            resizeDir = ResizeDirection.BottomLeft

        ElseIf e.Location.X > Me.Width - BorderWidth And e.Location.Y > Me.Height - BorderWidth Then
            resizeDir = ResizeDirection.BottomRight

        ElseIf e.Location.X > Me.Width - BorderWidth And e.Location.Y < BorderWidth Then
            resizeDir = ResizeDirection.TopRight

        ElseIf e.Location.X < BorderWidth Then
            resizeDir = ResizeDirection.Left

        ElseIf e.Location.X > Me.Width - BorderWidth Then
            resizeDir = ResizeDirection.Right

        ElseIf e.Location.Y < BorderWidth Then
            resizeDir = ResizeDirection.Top

        ElseIf e.Location.Y > Me.Height - BorderWidth Then
            resizeDir = ResizeDirection.Bottom

        Else
            resizeDir = ResizeDirection.None
        End If
    End Sub
#End Region
#Region "Drag Form"
    'agar form bisa di drag
    Private Sub Panel_Drag_MouseDown(sender As Object, e As MouseEventArgs) Handles Panel_Drag.MouseDown
        If e.Button = MouseButtons.Left Then
            drag = True
            xM = e.X
            yM = e.Y
        End If
    End Sub
    Private Sub Panel_Drag_MouseUp(sender As Object, e As MouseEventArgs) Handles Panel_Drag.MouseUp
        If e.Button = MouseButtons.Left Then
            drag = False
        End If
    End Sub
    Private Sub Panel_Drag_MouseMove(sender As Object, e As MouseEventArgs) Handles Panel_Drag.MouseMove
        If drag = True Then
            Dim tmp As Point = New Point()
            tmp.X = Me.Location.X + (e.X - xM)
            tmp.Y = Me.Location.Y + (e.Y - yM)
            Me.Location = tmp
            tmp = Nothing
        End If
    End Sub
    Private Sub PanelDrag_DoubleClick(sender As Object, e As EventArgs) Handles Panel_Drag.DoubleClick
        If Me.WindowState = FormWindowState.Maximized Then
            Me.WindowState = FormWindowState.Normal
        Else
            Me.WindowState = FormWindowState.Maximized
        End If
    End Sub
    Private Sub Label_ArisanDesktop_MouseDown(sender As Object, e As MouseEventArgs) Handles Label_ArisanDesktop.MouseDown
        If e.Button = MouseButtons.Left Then
            drag = True
            xM = e.X
            yM = e.Y
        End If
    End Sub
    Private Sub Label_ArisanDesktop_MouseUp(sender As Object, e As MouseEventArgs) Handles Label_ArisanDesktop.MouseUp
        If e.Button = MouseButtons.Left Then
            drag = False
        End If
    End Sub
    Private Sub Label_ArisanDesktop_MouseMove(sender As Object, e As MouseEventArgs) Handles Label_ArisanDesktop.MouseMove
        If drag = True Then
            Dim tmp As Point = New Point()
            tmp.X = Me.Location.X + (e.X - xM)
            tmp.Y = Me.Location.Y + (e.Y - yM)
            Me.Location = tmp
            tmp = Nothing
        End If
    End Sub
    Private Sub Label_ArisanDesktop_DoubleClick(sender As Object, e As EventArgs) Handles Label_ArisanDesktop.DoubleClick
        If Me.WindowState = FormWindowState.Maximized Then
            Me.WindowState = FormWindowState.Normal
        Else
            Me.WindowState = FormWindowState.Maximized
        End If
    End Sub
#End Region
End Class