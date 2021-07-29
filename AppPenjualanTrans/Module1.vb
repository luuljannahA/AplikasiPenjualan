Imports System.Data.OleDb

Module Module1
    Public Conn As OleDbConnection
    Public Cmd As OleDbCommand
    Public Da As OleDbDataAdapter
    Public Ds As DataSet
    Public Dt As DataTable
    Public Rd As OleDbDataReader
    Public LokasiData As String
    Public Sub koneksi()
        LokasiData = "provider=microsoft.jet.oledb.4.0;data source=dbPengelolaan.mdb"
        Conn = New OleDbConnection(LokasiData)
        If Conn.State = ConnectionState.Closed Then Conn.Open()
    End Sub
End Module
