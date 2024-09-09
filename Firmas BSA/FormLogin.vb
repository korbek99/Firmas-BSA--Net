
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Drawing.Imaging
Imports System.IO
Public Class FormLogin

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If LogintextoPlano() = True Then
            My.Forms.Formmenu.Show()
            'My.Forms.FormLogin.Close()
        Else
            MsgBox("Problemas en carga de datos para conexion al servidor servidor", MsgBoxStyle.Critical)

        End If
    End Sub
 
End Class