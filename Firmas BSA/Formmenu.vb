Public Class Formmenu

    Private Sub AgregarFirmaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AgregarFirmaToolStripMenuItem.Click
        My.Forms.FormIngreso.Show()
    End Sub

    Private Sub Formmenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        My.Forms.FormLogin.Hide()
        'xservidor = sSplit(0)
        'xusuario = sSplit(1)
        'xclave = sSplit(2)
        'base = sSplit(3)
        'Dsn = sSplit(4)
        Me.ToolStripStatusLabel1.Text = "Servidor Sql Server       :" & xservidor
        Me.ToolStripStatusLabel1.Width = 2000
        Me.ToolStripStatusLabel2.Text = "Usuario        :" & xusuario
        Me.ToolStripStatusLabel2.Width = 2000
        Me.ToolStripStatusLabel3.Text = "Base Datos Sql Server      :" & base
        Me.ToolStripStatusLabel3.Width = 2000
        Me.ToolStripStatusLabel4.Text = "Fecha          :" & DateTime.Today
        Me.ToolStripStatusLabel4.Width = 2000
    End Sub

    Private Sub BuscarFirmasCuentasToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BuscarFirmasCuentasToolStripMenuItem1.Click
        My.Forms.FormFirmas.Show()
    End Sub

    Private Sub SalirToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalirToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub BuscarFirmasCuentasToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BuscarFirmasCuentasToolStripMenuItem.Click

    End Sub
End Class