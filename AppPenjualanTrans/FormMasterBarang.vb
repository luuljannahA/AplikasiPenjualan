Imports System.Data.OleDb

Public Class FormMasterBarang
    Public tbl_barang As String
    Sub SiapIsi()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        ComboBox1.Enabled = True
        Call MunculSatuan()


    End Sub

    Sub KondisiAwal()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        ComboBox1.Text = ""
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        ComboBox1.Enabled = False

        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button1.Text = "Input"
        Button2.Text = "Edit"
        Button3.Text = "Hapus"
        Button4.Text = "Tutup"
        Call MunculGrid()
    End Sub

    Sub Isi()
        TextBox2.Clear()
        TextBox2.Focus()
    End Sub

    Sub MunculGrid()
        Call koneksi()
        Da = New OleDbDataAdapter("Select * From tbl_barang", Conn)
        Ds = New DataSet
        Ds.Clear()
        Da.Fill(Ds, "tbl_barang")
        DataGridView1.DataSource = (Ds.Tables("tbl_barang"))
    End Sub

    Sub MunculSatuan()
        Call koneksi()
        Cmd = New OleDbCommand("Select distinct satuan_barang from tbl_barang", Conn)
        Rd = Cmd.ExecuteReader
        ComboBox1.Items.Clear()
        Do While Rd.Read
            ComboBox1.Items.Add(Rd.Item("satuan_barang"))
        Loop

    End Sub

    Private Sub FormMasterBarang_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call KondisiAwal()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call koneksi()
        If Button1.Text = "Input" Then
            Button1.Text = "Simpan"
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Text = "Batal"
            Call SiapIsi()
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
                MsgBox("Data Ada Yang Kosong!", MsgBoxStyle.Information, "INFORMASI")
            Else
                Call koneksi()
                Dim SimpanData As String = "Insert into tbl_barang values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & ComboBox1.Text & "')"
                Cmd = New OleDbCommand(SimpanData, Conn)
                Cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil di Tambahkan!", MsgBoxStyle.Information, "INFORMASI")
                Call KondisiAwal()
                Call SiapIsi()
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call koneksi()
        If Button2.Text = "Edit" Then
            Button2.Text = "Simpan"
            Button1.Enabled = False
            Button3.Enabled = False
            Button4.Text = "Batal"
            TextBox1.Focus()
            Call SiapIsi()
        Else
            Call koneksi()
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
                MsgBox("Pastikan Data diisi Lengkap!", MsgBoxStyle.Information, "INFORMASI")
            Else
                Call koneksi()
                Dim EditData As String
                EditData = "UPDATE tbl_barang SET nama_barang='" & TextBox2.Text & "',harga_barang='" & TextBox3.Text & "',satuan_barang='" & ComboBox1.Text & "' where kode_barang='" & TextBox1.Text & "'"
                Cmd = New OleDbCommand(EditData, Conn)
                Cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil di Update!", MsgBoxStyle.Information, "INFORMASI")
                Call KondisiAwal()
                Call SiapIsi()
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Button3.Text = "Hapus" Then
            Button3.Text = "Simpan"
            Button1.Enabled = False
            Button2.Enabled = False
            Button4.Text = "Batal"
            TextBox1.Focus()
            Call SiapIsi()
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
                MsgBox("Data Ada Yang Kosong!", MsgBoxStyle.Information, "INFORMASI")
            Else
                Call koneksi()
                Dim HapusData As String = "Delete from tbl_barang where kode_barang= '" & TextBox1.Text & "'"
                Cmd = New OleDbCommand(HapusData, Conn)
                Cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil di Hapus!", MsgBoxStyle.Information, "INFORMASI")
                Call KondisiAwal()
                Call SiapIsi()
            End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If Button4.Text = "Tutup" Then
            Me.Close()
        Else
            Call KondisiAwal()
        End If
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Cmd = New OleDbCommand("select * from tbl_barang where kode_barang='" & TextBox1.Text & "'", Conn)
            Rd = Cmd.ExecuteReader
            Rd.Read()
            If Rd.HasRows Then
                MsgBox("Kode barang tidak ada!", MsgBoxStyle.Information, "INFORMASI")
                TextBox1.Focus()
            Else
                TextBox1.Text = Rd.Item("kode_barang")
                TextBox2.Text = Rd.Item("nama_barang")
                TextBox3.Text = Rd.Item("harga_barang")
                ComboBox1.Text = Rd.Item("satuan_barang")
            End If
                Call Isi()
            End If
    End Sub


    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim i As Integer
        i = Me.DataGridView1.CurrentRow.Index
        With DataGridView1.Rows.Item(i)
            Me.TextBox1.Text = .Cells(0).Value
            Me.TextBox2.Text = .Cells(1).Value
            Me.TextBox3.Text = .Cells(2).Value
            Me.ComboBox1.Text = .Cells(3).Value
        End With
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = Chr(13) Then
            Cmd = New OleDbCommand("select * from tbl_barang where kode_barang='" & TextBox1.Text & "'", Conn)
            Rd = Cmd.ExecuteReader
            Rd.Read()
            If Rd.HasRows Then
                MsgBox("Kode barang tidak ada!")
                TextBox2.Text = Rd.Item(1)
                TextBox2.Focus()
            Else
                TextBox1.Text = Rd.Item("kode_barang")
                TextBox2.Text = Rd.Item("nama_barang")
                TextBox3.Text = Rd.Item("harga_barang")
                ComboBox1.Text = Rd.Item("satuan_barang")
            End If
        End If
    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = vbBack) Then e.Handled = True
    End Sub

    Private Sub TextBox4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Not ((e.KeyChar >= "0" And e.KeyChar <= "9") Or e.KeyChar = vbBack) Then e.Handled = True
    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Call koneksi()
            Da = New OleDbDataAdapter("select * from tbl_barang where kode_barang like '" & TextBox4.Text & "%'", Conn)
        Ds = New DataSet
        Dim data As String
        data = Da.Fill(Ds)

        If data > 0 Then
            DataGridView1.DataSource = Ds.Tables(0)
            DataGridView1.ReadOnly = True
            Button1.Enabled = True
            Button2.Enabled = True
            Button3.Enabled = True
            Button4.Text = "BATAL"
        Else
            MsgBox("Data Tidak Ada!", MsgBoxStyle.Information, "INFORMASI")
            DataGridView1.DataSource = Nothing
        End If
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged

    End Sub

End Class