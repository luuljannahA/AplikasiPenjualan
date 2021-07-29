Imports System.Data.OleDb

Public Class FormMasterPelanggan

    Sub SiapIsi()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True

    End Sub
    Sub KondisiAwal()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False

        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button1.Text = "Input"
        Button2.Text = "Edit"
        Button3.Text = "Hapus"
        Button4.Text = "Tutup"
        Call MunculGrid()

    End Sub
    Sub MunculGrid()
        Call koneksi()
        Da = New OleDbDataAdapter("Select * From tbl_pelanggan", Conn)
        Ds = New DataSet
        Ds.Clear()
        Da.Fill(Ds, "tbl_pelanggan")
        DataGridView1.DataSource = (Ds.Tables("tbl_pelanggan"))

    End Sub

    Private Sub FormMasterPelanggan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call KondisiAwal()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call koneksi()
        If Button1.Text = "Input" Then
            Button1.Text = "Simpan"
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Text = "Batal"
            TextBox1.Focus()
            Call SiapIsi()
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MsgBox("Pastikan Data diisi Lengkap!", MsgBoxStyle.Information, "INFORMASI")
            Else
                Call koneksi()
                Dim SimpanData As String = "Insert into tbl_pelanggan values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "')"
                Cmd = New OleDbCommand(SimpanData, Conn)
                Cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil di Tambahkan", MsgBoxStyle.Information, "INFORMASI")
                Call KondisiAwal()
                Call SiapIsi()
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call koneksi()
        If Button3.Text = "Hapus" Then
            Button3.Text = "Hapus Data"
            Button1.Enabled = False
            Button2.Enabled = False
            Button4.Text = "Batal"
            TextBox1.Focus()
            Call SiapIsi()
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MsgBox("Pastikan Data diisi Lengkap!", MsgBoxStyle.Information, "INFORMASI")
            Else
                Call koneksi()
                Dim HapusData As String = "Delete from tbl_pelanggan where kode_pelanggan= '" & TextBox1.Text & "'"
                Cmd = New OleDbCommand(HapusData, Conn)
                Cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil di Hapus", MsgBoxStyle.Information, "INFORMASI")
                Call KondisiAwal()
                Call SiapIsi()
            End If
        End If

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If Button4.Text = "Batal" Then
            Call KondisiAwal()
        Else
            Me.Close()
        End If
    End Sub


    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = Chr(13) Then
            Cmd = New OleDbCommand("select * from tbl_pelanggan where kode_pelanggan='" & TextBox1.Text & "'", Conn)
            Rd = Cmd.ExecuteReader
            Rd.Read()
            If Rd.HasRows Then
                TextBox2.Text = Rd.Item(1)
                TextBox2.Focus()
            Else
                TextBox1.Text = Rd.Item("KodePelanggan")
                TextBox2.Text = Rd.Item("NamaPelanggan")
                TextBox3.Text = Rd.Item("AlamatPelanggan")
                TextBox4.Text = Rd.Item("TeleponPelanggan")
            End If
        End If
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim i As Integer
        i = Me.DataGridView1.CurrentRow.Index
        With DataGridView1.Rows.Item(i)
            Me.TextBox1.Text = .Cells(0).Value
            Me.TextBox2.Text = .Cells(1).Value
            Me.TextBox3.Text = .Cells(2).Value
            Me.TextBox4.Text = .Cells(3).Value
        End With
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call koneksi()
            Cmd = New OleDbCommand("select * from tbl_pelanggan where kode_pelanggan='" & TextBox1.Text & "'", Conn)
            Rd = Cmd.ExecuteReader
            Rd.Read()
            If Rd.HasRows Then
                TextBox2.Text = Rd.Item("nama_pelanggan")
                TextBox3.Text = Rd.Item("alamat_pelanggan")
                TextBox4.Text = Rd.Item("telepon_pelanggan")
            Else
                MsgBox("Data tidak ada!", MsgBoxStyle.Information, "INFORMASI")
            End If
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

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
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MsgBox("Pastikan Data diisi Lengkap!", MsgBoxStyle.Information, "INFORMASI")
            Else
                Call koneksi()
                Dim EditData As String = "Update tbl_pelanggan set nama_pelanggan='" & TextBox2.Text & "',alamat_pelanggan= '" & TextBox3.Text & "',telepon_pelanggan='" & TextBox4.Text & "' where kode_pelanggan= '" & TextBox1.Text & "'"
                Cmd = New OleDbCommand(EditData, Conn)
                Cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil di Update", MsgBoxStyle.Information, "INFORMASI")
                Call KondisiAwal()
                Call SiapIsi()
            End If
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Call koneksi()
        Da = New OleDbDataAdapter("select * from tbl_pelanggan where kode_pelanggan like '" & TextBox5.Text & "%'", Conn)
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
End Class