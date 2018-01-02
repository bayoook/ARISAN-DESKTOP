Imports System.Data.SqlClient
Imports System.Runtime.InteropServices
Public Class FormUser
#Region "Variabel"
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
    Sub TampilGrid()
        Call Koneksi()
        CMD = New SqlCommand("select * from TBL_ACC where Id='" & id & "'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        DataGridView_User.Rows.Clear()
        DataGridView_User.ColumnCount = 2
        DataGridView_User.Columns(0).Name = "Bulan"
        DataGridView_User.Columns(1).Name = "Pembayaran"
        For i = 0 To 1
            'Mengecilkan jarak antar kolom
            DataGridView_User.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        Next
        Dim row As String() = New String() {"Januari", RD.Item("Januari")}
        DataGridView_User.Rows.Add(row)
        row = New String() {"Februari", RD.Item("Februari")}
        DataGridView_User.Rows.Add(row)
        row = New String() {"Maret", RD.Item("Maret")}
        DataGridView_User.Rows.Add(row)
        row = New String() {"April", RD.Item("April")}
        DataGridView_User.Rows.Add(row)
        row = New String() {"Mei", RD.Item("Mei")}
        DataGridView_User.Rows.Add(row)
        row = New String() {"Juni", RD.Item("Juni")}
        DataGridView_User.Rows.Add(row)
        row = New String() {"Juli", RD.Item("Juli")}
        DataGridView_User.Rows.Add(row)
        row = New String() {"Agustus", RD.Item("Agustus")}
        DataGridView_User.Rows.Add(row)
        row = New String() {"September", RD.Item("September")}
        DataGridView_User.Rows.Add(row)
        row = New String() {"Oktober", RD.Item("Oktober")}
        DataGridView_User.Rows.Add(row)
        row = New String() {"November", RD.Item("November")}
        DataGridView_User.Rows.Add(row)
        row = New String() {"Desember", RD.Item("Desember")}
        DataGridView_User.Rows.Add(row)
        RD.Close()
    End Sub
    Sub TampilList()
        Call Koneksi()
        CMD = New SqlCommand("select * from TBL_ACC where username='" & user & "'", CONN)
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
    End Sub
    Sub TampilBulan()
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
#End Region
#Region "Button"
    'ketika button input di klik
    Private Sub Button_Input_Click(sender As Object, e As EventArgs) Handles Button_Input.Click
        Dim jumlah As Integer
        If TextBox_JumlahUang.Text = "" Or ComboBox_Bulan.Text = "" Then
            MsgBox("Data belum lengkap, mohon diisi semua ya")
            Exit Sub
        Else
            Call Koneksi()
            CMD = New SqlCommand("select * from TBL_ACC where username='" & user & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            jumlah = RD.Item(ComboBox_Bulan.Text) + TextBox_JumlahUang.Text
            RD.Close()
            'string edit diisi dengan kata2 untuk menupdate database
            Dim edit As String = "update TBL_ACC set " & ComboBox_Bulan.Text & "='" & jumlah & "' where Username='" & user & "'"
            'memasukkan string edit tadi untuk diproses oleh database
            CMD = New SqlCommand(edit, CONN)
            'mengexekusi database
            CMD.ExecuteNonQuery()
            Call TampilGrid()
            Call TampilBulan()
            Call TampilList()
        End If
    End Sub
    'ketika button delete di click
    Private Sub Button_Delete_Click(sender As Object, e As EventArgs) Handles Button_Delete.Click
        If TextBox_JumlahUang.Text = "" Or ComboBox_Bulan.Text = "" Then
            MsgBox("Silahkan klik data di bawah")
            Exit Sub
        Else
            Call Koneksi()
            'string edit diisi dengan perintah update database
            Dim delete As String = "update TBL_ACC set " & ComboBox_Bulan.Text & "='0' where Username='" & user & "'"
            'mengeksekusi string edit ke database
            CMD = New SqlCommand(delete, CONN)
            'mengeksekusi database
            CMD.ExecuteNonQuery()
            Call TampilGrid()
            Call TampilBulan()
            Call TampilList()
        End If
    End Sub
#End Region
#Region "Data Click"
    'ketika data grid view di klik
    Private Sub DataGridView_User_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView_User.CellClick
        Call Koneksi()
        Dim i As Integer
        i = DataGridView_User.CurrentRow.Index
        ComboBox_Bulan.SelectedItem = DataGridView_User.Item(0, i).Value
        TextBox_JumlahUang.Text = DataGridView_User.Item(1, i).Value
        TextBox_JumlahUang.Focus()
    End Sub
    'listbox tagihan click
    Private Sub ListBox_Tagihan_Click(sender As Object, e As EventArgs) Handles ListBox_Tagihan.Click
        ComboBox_Bulan.SelectedIndex = ListBox_Tagihan.SelectedIndex
        Call Koneksi()
        CMD = New SqlCommand("select * from TBL_ACC where username='" & user & "'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        Dim jumlah As Integer = 1000000 - RD.Item(ComboBox_Bulan.SelectedItem)
        TextBox_JumlahUang.Text = jumlah
    End Sub
#End Region
#Region "Menu"
    Private Sub Button_TagihanArisan_Click(sender As Object, e As EventArgs) Handles Button_TagihanArisan.Click
        DataGridView_User.Visible = False
        ListBox_Tagihan.Visible = True
    End Sub
    Private Sub Button_TabelArisan_Click(sender As Object, e As EventArgs) Handles Button_TabelArisan.Click
        DataGridView_User.Visible = True
        ListBox_Tagihan.Visible = False
    End Sub
#End Region
#Region "Others"
    Private Sub Picture_User_Click(sender As Object, e As EventArgs) Handles Picture_User.Click
        ShowPict.id = id
        ShowPict.path = path
        ShowPict.ShowDialog()
    End Sub
    Private Sub Label_Logout_Click_1(sender As Object, e As EventArgs) Handles Label_Logout.Click
        Me.Close()
    End Sub
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        TampilBulan()
    End Sub
    Private Sub TextBox_JumlahUang_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox_JumlahUang.KeyPress
        If e.KeyChar = Chr(13) Then Button_Input.PerformClick()
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub
    'Ketika form diaktifkan
    Private Sub FormUser_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TampilBulan()
        TampilGrid()
        TampilList()
        TextBox_JumlahUang.Focus()
        TextBox_JumlahUang.Text = ""
        Label_Username.Text = user
        Label_Email.Text = mail
        Picture_User.ImageLocation = path
        If path = "" Then
            Picture_User.Image = Aplikasi_Keuangan.My.Resources.Resources.blank_user
        End If
    End Sub
    Private Sub FormUser_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        TampilBulan()
        TampilGrid()
        TampilList()
        TextBox_JumlahUang.Focus()
        Label_Username.Text = user
        Label_Email.Text = mail
        Picture_User.ImageLocation = path
        If path = "" Then
            Picture_User.Image = Aplikasi_Keuangan.My.Resources.Resources.blank_user
        End If
    End Sub
#End Region
#Region "Panel Atas"
    'button _ click
    Private Sub Button___Click(sender As Object, e As EventArgs) Handles Button__.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub\
    'button maximize click
    Private Sub Button_Max_Click(sender As Object, e As EventArgs) Handles Button_Max.Click
        If Me.WindowState = FormWindowState.Maximized Then
            Me.WindowState = FormWindowState.Normal
        Else
            Me.WindowState = FormWindowState.Maximized
        End If
    End Sub
    'button x click
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
    Private Sub Panel_Drag_DoubleClick(sender As Object, e As EventArgs) Handles Panel_Drag.DoubleClick
        If Me.WindowState = FormWindowState.Maximized Then
            Me.WindowState = FormWindowState.Normal
        Else
            Me.WindowState = FormWindowState.Maximized
        End If
    End Sub
#End Region
End Class