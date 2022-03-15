Imports System.Data.OleDb
Module Koneksi
    Public PerintahDatabase As OleDbCommand
    Public DataAdapter As OleDbDataAdapter
    Public DataSet As DataSet
    Public DataReader As OleDbDataReader
    Public SqlQuery As String

    Public NamaUser As String = "Admin"
    Public Conn As New OleDbConnection("provider=microsoft.ace.oledb.12.0;data source=wp.mdb")
    Public cmd As OleDbCommand
    Public da As OleDbDataAdapter
    Public ds As DataSet
    Public rd As OleDbDataReader
    Public query As String
    Public QueryEbbim As String

    Public Cmdlogin As OleDbCommand
    Public Dalogin As OleDbDataAdapter
    Public Dslogin As DataSet
    Public Drlogin As OleDbDataReader

    Public Cmddaftar As OleDbCommand
    Public Dadaftar As OleDbDataAdapter
    Public Dsdaftar As DataSet
    Public Drdaftar As OleDbDataReader

    Sub konek()
        If Conn.State = ConnectionState.Closed Then
            Conn.Open()
        End If
    End Sub
End Module
