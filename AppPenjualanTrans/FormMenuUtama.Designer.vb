<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormMenuUtama
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LoginToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LogoutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.KeluarToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.MasterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.BarangToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PelangganToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TransaksiToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PenjualanToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LaporanToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LaporanDataMasterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.STLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.STLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.STLabel3 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.STLabel4 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.STLabel5 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.STLabel6 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.STLabel7 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.STLabel8 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.STLabel9 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.STLabel10 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.BackupDataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.ToolStripMenuItem1, Me.TransaksiToolStripMenuItem, Me.LaporanToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1034, 33)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.LoginToolStripMenuItem, Me.LogoutToolStripMenuItem, Me.KeluarToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(50, 29)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'LoginToolStripMenuItem
        '
        Me.LoginToolStripMenuItem.Name = "LoginToolStripMenuItem"
        Me.LoginToolStripMenuItem.Size = New System.Drawing.Size(141, 30)
        Me.LoginToolStripMenuItem.Text = "Login"
        '
        'LogoutToolStripMenuItem
        '
        Me.LogoutToolStripMenuItem.Name = "LogoutToolStripMenuItem"
        Me.LogoutToolStripMenuItem.Size = New System.Drawing.Size(141, 30)
        Me.LogoutToolStripMenuItem.Text = "Logout"
        '
        'KeluarToolStripMenuItem
        '
        Me.KeluarToolStripMenuItem.Name = "KeluarToolStripMenuItem"
        Me.KeluarToolStripMenuItem.Size = New System.Drawing.Size(141, 30)
        Me.KeluarToolStripMenuItem.Text = "Keluar"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MasterToolStripMenuItem, Me.BarangToolStripMenuItem, Me.PelangganToolStripMenuItem})
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(78, 29)
        Me.ToolStripMenuItem1.Text = "Master"
        '
        'MasterToolStripMenuItem
        '
        Me.MasterToolStripMenuItem.Name = "MasterToolStripMenuItem"
        Me.MasterToolStripMenuItem.Size = New System.Drawing.Size(166, 30)
        Me.MasterToolStripMenuItem.Text = "Admin"
        '
        'BarangToolStripMenuItem
        '
        Me.BarangToolStripMenuItem.Name = "BarangToolStripMenuItem"
        Me.BarangToolStripMenuItem.Size = New System.Drawing.Size(166, 30)
        Me.BarangToolStripMenuItem.Text = "Barang"
        '
        'PelangganToolStripMenuItem
        '
        Me.PelangganToolStripMenuItem.Name = "PelangganToolStripMenuItem"
        Me.PelangganToolStripMenuItem.Size = New System.Drawing.Size(166, 30)
        Me.PelangganToolStripMenuItem.Text = "Pelanggan"
        '
        'TransaksiToolStripMenuItem
        '
        Me.TransaksiToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.PenjualanToolStripMenuItem})
        Me.TransaksiToolStripMenuItem.Name = "TransaksiToolStripMenuItem"
        Me.TransaksiToolStripMenuItem.Size = New System.Drawing.Size(94, 29)
        Me.TransaksiToolStripMenuItem.Text = "Transaksi"
        '
        'PenjualanToolStripMenuItem
        '
        Me.PenjualanToolStripMenuItem.Name = "PenjualanToolStripMenuItem"
        Me.PenjualanToolStripMenuItem.Size = New System.Drawing.Size(158, 30)
        Me.PenjualanToolStripMenuItem.Text = "Penjualan"
        '
        'LaporanToolStripMenuItem
        '
        Me.LaporanToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.LaporanDataMasterToolStripMenuItem, Me.BackupDataToolStripMenuItem})
        Me.LaporanToolStripMenuItem.Name = "LaporanToolStripMenuItem"
        Me.LaporanToolStripMenuItem.Size = New System.Drawing.Size(93, 29)
        Me.LaporanToolStripMenuItem.Text = "Laporan "
        '
        'LaporanDataMasterToolStripMenuItem
        '
        Me.LaporanDataMasterToolStripMenuItem.Name = "LaporanDataMasterToolStripMenuItem"
        Me.LaporanDataMasterToolStripMenuItem.Size = New System.Drawing.Size(195, 30)
        Me.LaporanDataMasterToolStripMenuItem.Text = "Laporan Data "
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.STLabel1, Me.STLabel2, Me.STLabel3, Me.STLabel4, Me.STLabel5, Me.STLabel6, Me.STLabel7, Me.STLabel8, Me.STLabel9, Me.STLabel10})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 548)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(2, 0, 14, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(1034, 30)
        Me.StatusStrip1.TabIndex = 2
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'STLabel1
        '
        Me.STLabel1.Name = "STLabel1"
        Me.STLabel1.Size = New System.Drawing.Size(68, 25)
        Me.STLabel1.Text = "Nama :"
        '
        'STLabel2
        '
        Me.STLabel2.Name = "STLabel2"
        Me.STLabel2.Size = New System.Drawing.Size(0, 25)
        '
        'STLabel3
        '
        Me.STLabel3.Name = "STLabel3"
        Me.STLabel3.Size = New System.Drawing.Size(68, 25)
        Me.STLabel3.Text = "Nama :"
        '
        'STLabel4
        '
        Me.STLabel4.Name = "STLabel4"
        Me.STLabel4.Size = New System.Drawing.Size(0, 25)
        '
        'STLabel5
        '
        Me.STLabel5.Name = "STLabel5"
        Me.STLabel5.Size = New System.Drawing.Size(60, 25)
        Me.STLabel5.Text = "Level :"
        '
        'STLabel6
        '
        Me.STLabel6.Name = "STLabel6"
        Me.STLabel6.Size = New System.Drawing.Size(0, 25)
        '
        'STLabel7
        '
        Me.STLabel7.Name = "STLabel7"
        Me.STLabel7.Size = New System.Drawing.Size(52, 25)
        Me.STLabel7.Text = "Jam :"
        '
        'STLabel8
        '
        Me.STLabel8.Name = "STLabel8"
        Me.STLabel8.Size = New System.Drawing.Size(0, 25)
        '
        'STLabel9
        '
        Me.STLabel9.Name = "STLabel9"
        Me.STLabel9.Size = New System.Drawing.Size(82, 25)
        Me.STLabel9.Text = "Tanggal :"
        '
        'STLabel10
        '
        Me.STLabel10.Name = "STLabel10"
        Me.STLabel10.Size = New System.Drawing.Size(0, 25)
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label1.Location = New System.Drawing.Point(47, 350)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(939, 37)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "APLIKASI PENGELOLAAN DATA TRANSAKSI PENJUALAN"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Kristen ITC", 36.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Red
        Me.Label2.Location = New System.Drawing.Point(204, 127)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(639, 98)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = " Emping Top Sari"
        '
        'BackupDataToolStripMenuItem
        '
        Me.BackupDataToolStripMenuItem.Name = "BackupDataToolStripMenuItem"
        Me.BackupDataToolStripMenuItem.Size = New System.Drawing.Size(195, 30)
        Me.BackupDataToolStripMenuItem.Text = "Backup Data"
        '
        'FormMenuUtama
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SkyBlue
        Me.ClientSize = New System.Drawing.Size(1034, 578)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Name = "FormMenuUtama"
        Me.Text = "FormMenuUtama"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LoginToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LogoutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents KeluarToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MasterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BarangToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TransaksiToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PenjualanToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LaporanToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents STLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents STLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents STLabel3 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents STLabel4 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents STLabel5 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents STLabel6 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents STLabel7 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents STLabel8 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents STLabel9 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents STLabel10 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents PelangganToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LaporanDataMasterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents BackupDataToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
