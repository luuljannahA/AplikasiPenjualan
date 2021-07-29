Imports System.Data
Imports System.Data.OleDb

Public Class FormTransJual

    Sub KondisiAwal()
        LBLNamaPlg.Text = ""
        LBLAlamat.Text = ""
        LBLTelepon.Text = ""
        LBLTanggal.Text = Today
        LBLAdmin.Text = FormMenuUtama.STLabel4.Text
        LBLKembali.Text = ""
        ComboBox2.Text = ""
        ComboBox1.Text = ""
        LBLNamaBarang.Text = ""
        LBLHargaBarang.Text = ""
        TextBox1.Text = ""
        TextBox4.Text = ""
        LBLItem.Text = ""
        Call tampilKodeBarangComboBox()
        Call tambahNoJual()
        Call BuatKolom()
        Label9.Text = "0"
        TextBox1.Text = ""


    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        LBLJam.Text = TimeOfDay
    End Sub
    Sub BuatKolom()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("kode", "kode")
        DataGridView1.Columns.Add("nama", "nama_barang")
        DataGridView1.Columns.Add("harga", "harga")
        DataGridView1.Columns.Add("jumlah", "jumlah")
        DataGridView1.Columns.Add("subtotal", "subtotal")
    End Sub

    Sub MunculGrid()
        Call koneksi()
        Da = New OleDbDataAdapter("Select * From tbl_jual", Conn)
        Ds = New DataSet
        Ds.Clear()
        Da.Fill(Ds, "tbl_jual")
        DataGridView1.DataSource = (Ds.Tables("tbl_jual"))

    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
       
      
    End Sub

    Sub RumusSubtotal()
        Dim hitung As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            hitung = hitung + DataGridView1.Rows(i).Cells(4).Value
            Label9.Text = hitung
        Next
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            If Val(TextBox1.Text) < Val(Label9.Text) Then
                MsgBox("Pembayaran kurang!", MsgBoxStyle.Information, "INFORMASI")
            ElseIf Val(TextBox1.Text) = Val(Label9.Text) Then
                LBLKembali.Text = 0
            ElseIf Val(TextBox1.Text) > Val(Label9.Text) Then
                LBLKembali.Text = Val(TextBox1.Text) - Val(Label9.Text)
                Button1.Focus()
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Button2.Text = "Tutup" Then
            Me.Close()
        Else
            Call KondisiAwal()
        End If

    End Sub

    Private Sub FormTransJual_Load_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call koneksi()
        Call KondisiAwal()
        Call BuatKolom()
        Label9.Text = "0"
        Call tampilKodePelangganComboBox()
        Call tampilKodeBarangComboBox()
        Call tambahNoJual()
    End Sub

    Sub tampilKodePelangganComboBox()
        Try
            Call koneksi()
            Cmd = New OleDbCommand("Select * from tbl_pelanggan", Conn)
            Rd = Cmd.ExecuteReader
            If Rd.HasRows Then
                Do While Rd.Read
                    ComboBox1.Items.Add(Rd("kode_pelanggan"))
                Loop
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Sub tampilKodeBarangComboBox()
        Try
            Call koneksi()
            ComboBox2.Items.Clear()
            Cmd = New OleDbCommand("Select * from tbl_barang", Conn)
            Rd = Cmd.ExecuteReader
            If Rd.HasRows Then
                Do While Rd.Read
                    ComboBox2.Items.Add(Rd("kode_barang"))
                Loop
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Sub tampilTextBox()
        Try

            Call koneksi()
            Dim str As String
            str = "Select * from tbl_pelanggan where kode_pelanggan = '" & ComboBox1.Text & "'"
            Cmd = New OleDbCommand(str, Conn)
            Rd = Cmd.ExecuteReader
            Rd.Read()
            If Rd.HasRows Then
                LBLNamaPlg.Text = Rd.Item("nama_pelanggan")
                LBLAlamat.Text = Rd.Item("alamat_pelanggan")
                LBLTelepon.Text = Rd.Item("telepon_pelanggan")
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Sub tampilTextBoxBarang()
        Try

            Call koneksi()
            Dim str As String
            str = "Select * from tbl_barang where kode_barang = '" & ComboBox2.Text & "'"
            Cmd = New OleDbCommand(str, Conn)
            Rd = Cmd.ExecuteReader
            Rd.Read()
            If Rd.HasRows Then
                LBLNamaBarang.Text = Rd.Item("nama_barang")
                LBLHargaBarang.Text = Rd.Item("harga_barang")
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub


    Sub tambahNoJual()
        Try
            Call koneksi()
            Cmd = New OleDbCommand("Select * from tbl_jual where no_jual in (select max(no_jual) from tbl_jual)", Conn)
            Dim urutan As String
            Dim hitung As Long
            Rd = Cmd.ExecuteReader
            Rd.Read()
            If Not Rd.HasRows Then
                urutan = Format(Now, "ddMMyy") + "001"
            Else
                hitung = Microsoft.VisualBasic.Right(Rd.GetString(0), 3) + 1
                urutan = Format(Now, "ddMMyy") + Microsoft.VisualBasic.Right("000" & hitung, 3)
            End If
            LBLNoJual.Text = urutan
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call tampilTextBox()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If LBLNamaBarang.Text = "" Or TextBox4.Text = "" Then
            MsgBox("silahkan masukkan kode barang dan tekan enter", MsgBoxStyle.Information, "INFORMASI")
        Else
            DataGridView1.Rows.Add(New String() {ComboBox2.Text, LBLNamaBarang.Text, LBLHargaBarang.Text, TextBox4.Text, Val(LBLHargaBarang.Text) * Val(TextBox4.Text)})
            Call RumusSubtotal()
            ComboBox2.Text = ""
            LBLNamaBarang.Text = ""
            LBLHargaBarang.Text = ""
            Call RumusCariItem()
        End If
    End Sub


    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Call tampilTextBoxBarang()
    End Sub
    Sub RumusCariItem()
        Dim HitungItem As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            HitungItem = HitungItem + DataGridView1.Rows(i).Cells(3).Value
            LBLItem.Text = HitungItem
        Next
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If LBLKembali.Text = "" Or LBLNamaPlg.Text = "" Or Label9.Text = "" Then
            MsgBox("Transaksi tidak ada, silahkan lakukan transaksi terlebih dahulu!", MsgBoxStyle.Information, "INFORMASI")
        Else
            Call koneksi()
            Dim SimpanJual As String = "insert into tbl_jual values ('" & LBLNoJual.Text & "','" & LBLTanggal.Text & "','" & LBLJam.Text & "','" & LBLItem.Text & "','" & Label9.Text & "','" & TextBox1.Text & "','" & LBLKembali.Text & "','" & ComboBox1.Text & "','" & FormMenuUtama.STLabel2.Text & "') "
            Cmd = New OleDbCommand(SimpanJual, Conn)
            Cmd.ExecuteNonQuery()
            Call KondisiAwal()
            Dim konfirmasi As String
            konfirmasi = MsgBox("Apakah Anda ingin Mencetak Nota?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "KONFIRMASI")
            If konfirmasi = vbYes Then
                FNotaJual.NotaPenjualan1.SetParameterValue("NomorJual", LBLNoJual.Text)
                FNotaJual.Show()
            End If
            Me.Close()
        End If
    End Sub
End Class