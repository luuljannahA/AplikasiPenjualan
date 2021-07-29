Public Class FormMenuUtama
    Sub Terkunci()
        LoginToolStripMenuItem.Enabled = True
        LogoutToolStripMenuItem.Enabled = False
        MasterToolStripMenuItem.Enabled = False
        BarangToolStripMenuItem.Enabled = False
        TransaksiToolStripMenuItem.Enabled = False
        ToolStripMenuItem1.Enabled = False
        LaporanToolStripMenuItem.Enabled = False
        PelangganToolStripMenuItem.Enabled = False
        STLabel2.Text = ""
        STLabel4.Text = ""
        STLabel6.Text = ""

    End Sub

    Private Sub LoginToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoginToolStripMenuItem.Click
        FormLogin.Show()
    End Sub

    Private Sub KeluarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KeluarToolStripMenuItem.Click
        Dim konfirmasi As String
        konfirmasi = MsgBox("Apakah Anda yakin akan KELUAR?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "KONFIRMASI")
        If konfirmasi = vbYes Then
            Me.Close()
            End
        End If
    End Sub

    Private Sub LogoutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogoutToolStripMenuItem.Click
        Dim konfirmasi As String
        konfirmasi = MsgBox("Apakah Anda yakin akan LOGOUT?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "KONFIRMASI")
        If konfirmasi = vbYes Then
            Call Terkunci()
            MsgBox("Anda berhasil LOGOUT!", MsgBoxStyle.Information, "INFORMASI")
        End If
    End Sub

    Private Sub MasterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MasterToolStripMenuItem.Click
        FormMasterAdmin.Show()
    End Sub


    Private Sub PelangganToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub BarangToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarangToolStripMenuItem.Click
        FormMasterBarang.Show()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        STLabel8.Text = TimeOfDay
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click

    End Sub

    Private Sub PenjualanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PenjualanToolStripMenuItem.Click
        FormTransJual.Show()
    End Sub

    Private Sub TransaksiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TransaksiToolStripMenuItem.Click

    End Sub
    Private Sub FormMenuUtama_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Terkunci()
        STLabel10.Text = Today
    End Sub

    Private Sub PelangganToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PelangganToolStripMenuItem.Click
        FormMasterPelanggan.Show()
    End Sub

 
    Private Sub LaporanDataMasterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaporanDataMasterToolStripMenuItem.Click
        FormLapDataMaster.Show()
    End Sub

    Private Sub LaporanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaporanToolStripMenuItem.Click

    End Sub

    Private Sub BackupDataToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        FormBackup.Show()
    End Sub

    Private Sub BackupDataToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackupDataToolStripMenuItem.Click
        FormBackup.Show()
    End Sub
End Class