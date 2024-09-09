Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb

Imports System.Configuration
Imports System.Drawing.Imaging
Imports System.IO
Module ModuleConexion
    Public cnn As New SqlConnection()
    Public cmd As New SqlCommand()
    Public Conn As SqlClient.SqlConnection
    Public Reader As OleDb.OleDbDataReader
    Public xservidor As String
    Public xusuario As String
    Public xclave As String
    Public base As String
    Public Dsn As String
    Public Inipath As String

    Public Function Conexion()
        'xservidor = "USUARIO1" 'HostServer
        'xusuario = "sa"
        'xclave = "sa"
        'base = "TOPAZ_DATOS"
        'Dsn = "DsnPruebas"

        Conn = New SqlClient.SqlConnection()
        Conn.ConnectionString = "data source=" & xservidor & ";database=" & base & ";User ID  =" & xusuario & ";Password =" & xclave & ";Trusted_Connection=yes;"
        'Conn.ConnectionString = "DSN=" & Dsn & ";database=" & base & ";User ID  =" & xusuario & ";Password =" & xclave & ""
        Conn.Open()
    End Function
    Public Function LogintextoPlano() As Boolean
        Try
            Dim reader As IO.StreamReader = New StreamReader("C:\Inicio.ini")
            Dim contenido As String
            Dim sSplit() As String
            contenido = reader.ReadToEnd()
            reader.Close()

            sSplit = Split(contenido, ";")

            xservidor = sSplit(0)
            xusuario = sSplit(1)
            xclave = sSplit(2)
            base = sSplit(3)
            Dsn = sSplit(4)

            MsgBox("  Dsn :   " & Dsn & "  SERVIDOR    :" & xservidor & " USUARIO : " & xusuario & "  CLAVE : " & xclave & "  BASE :" & base)
            Conn = New SqlClient.SqlConnection()
            Conn.ConnectionString = "data source=" & xservidor & ";database=" & base & ";User ID  =" & xusuario & ";Password =" & xclave & ";Trusted_Connection=yes;"
            'Conn.ConnectionString = "DSN=" & Dsn & ";database=" & base & ";User ID  =" & xusuario & ";Password =" & xclave & ""
            Conn.Open()
            LogintextoPlano = True
        Catch var As SqlException
            MessageBox.Show(var.ToString, "Conexion Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            LogintextoPlano = False
        Catch var As Exception
            MsgBox(var.Message)
            LogintextoPlano = False
        End Try
    End Function
End Module
