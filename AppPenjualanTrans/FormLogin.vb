Imports System.Data.OleDb

Public Class FormLogin

    Private Sub FormLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call KondisiAwal()
    End Sub
    Sub KondisiAwal()
        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub
    Sub Terbuka()
        FormMenuUtama.LoginToolStripMenuItem.Enabled = False
        FormMenuUtama.LogoutToolStripMenuItem.Enabled = True
        FormMenuUtama.MasterToolStripMenuItem.Enabled = True
        FormMenuUtama.TransaksiToolStripMenuItem.Enabled = True
        FormMenuUtama.ToolStripMenuItem1.Enabled = True
        FormMenuUtama.LaporanToolStripMenuItem.Enabled = True
        FormMenuUtama.BarangToolStripMenuItem.Enabled = True
        FormMenuUtama.PenjualanToolStripMenuItem.Enabled = True
        FormMenuUtama.PelangganToolStripMenuItem.Enabled = True
        FormMenuUtama.BackupDataToolStripMenuItem.Enabled = True



    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Data belum lengkap, silahkan isi data!", MsgBoxStyle.Information, "INFORMASI")
        Else
            Call koneksi()
            Cmd = New OleDbCommand("select * from tbl_admin where kode_admin ='" & TextBox1.Text & "' and password_admin ='" & TextBox2.Text & "'", Conn)
            Rd = Cmd.ExecuteReader
            Rd.Read()
            If Rd.HasRows Then
                Me.Close()
                Call Terbuka()
                FormMenuUtama.STLabel2.Text = Rd!kode_admin
                FormMenuUtama.STLabel4.Text = Rd!nama_admin
                FormMenuUtama.STLabel6.Text = Rd!level_admin
                MsgBox("Selamat anda telah login!", MsgBoxStyle.Information, "INFORMASI")
            Else
                MsgBox("Kode admin atau password salah!", MsgBoxStyle.Information, "INFORMASI")
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        FormMenuUtama.Show()
        Me.Hide()
    End Sub

End Class
