Public Class FormLapDataMaster

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        FDaftarAdmin.CrystalReportViewer1.RefreshReport()
        FDaftarAdmin.Show()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        FDaftarPelanggan.CrystalReportViewer1.RefreshReport()
        FDaftarPelanggan.Show()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        FDaftarBarang.CrystalReportViewer1.RefreshReport()
        FDaftarBarang.Show()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        FDataJual.CrystalReportViewer1.RefreshReport()
        FDataJual.Show()
    End Sub

    Private Sub FormLapDataMaster_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class