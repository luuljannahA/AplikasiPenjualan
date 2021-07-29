Imports System.IO

Public Class FormBackup
    Sub clear()
        TxtDatabase.Clear()
        TxtLocation.Clear()
        TxtBackup.Clear()
        TxtLokasiBackUp.Clear()
    End Sub
    Private Sub Backup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        clear()
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            If TxtBackup.Text = "" Then
                MsgBox("Tentukan database yang akan dibackup", MsgBoxStyle.Critical)
                Exit Sub
            End If
            If TxtLokasiBackup.Text = "" Then
                MsgBox("Tentukan lokasi penyimpanan database yang akan dibackup", MsgBoxStyle.Critical)
            End If
            IO.File.Copy(TxtLocation.Text, TxtLokasiBackup.Text + "\" + TxtBackup.Text, True)
            MsgBox("Database berhasil dibackup", MsgBoxStyle.Information)
            Process.Start(TxtLokasiBackup.Text)
            clear()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try

            With OpenFileDialog1
                .FileName = ""
                .Filter = "Access 2000-2003 (*.mdb)|*.mdb|Access 2007 (*.accdb)|*accdb"
                .Title = "Select Database !"
                .InitialDirectory = Application.StartupPath
            End With

            Dim x, y, name As String
            If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                y = OpenFileDialog1.FileName
                x = Split(y, "\").Length - 1
                name = Split(y, "\")(x)
                TxtDatabase.Text = name
                TxtLocation.Text = y
                TxtBackup.Text = "backup_" + TxtDatabase.Text
            End If
        Catch ex As Exception
            : Exit Sub
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            TxtLokasiBackup.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub
End Class