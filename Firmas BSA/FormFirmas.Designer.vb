<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormFirmas
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Button2 = New System.Windows.Forms.Button
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.GrillasFirmas = New System.Windows.Forms.DataGridView
        Me.CheckBox = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Button4 = New System.Windows.Forms.Button
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.Textfinal = New System.Windows.Forms.TextBox
        Me.Textini = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.Texttotal = New System.Windows.Forms.TextBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Textmax = New System.Windows.Forms.TextBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Textmin = New System.Windows.Forms.TextBox
        Me.Button5 = New System.Windows.Forms.Button
        Me.PanProgress = New System.Windows.Forms.GroupBox
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        CType(Me.GrillasFirmas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.PanProgress.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(33, 538)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(202, 23)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Generar Imagenes en Directorio"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox1.Location = New System.Drawing.Point(33, 154)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(132, 17)
        Me.CheckBox1.TabIndex = 2
        Me.CheckBox1.Text = "Seleccionar Todas"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'GrillasFirmas
        '
        Me.GrillasFirmas.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GrillasFirmas.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CheckBox})
        Me.GrillasFirmas.Location = New System.Drawing.Point(33, 187)
        Me.GrillasFirmas.Name = "GrillasFirmas"
        Me.GrillasFirmas.Size = New System.Drawing.Size(723, 345)
        Me.GrillasFirmas.TabIndex = 3
        '
        'CheckBox
        '
        Me.CheckBox.HeaderText = "Check"
        Me.CheckBox.Name = "CheckBox"
        Me.CheckBox.TrueValue = ""
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.PictureBox1)
        Me.GroupBox1.Controls.Add(Me.Button4)
        Me.GroupBox1.Controls.Add(Me.GroupBox5)
        Me.GroupBox1.Controls.Add(Me.GroupBox4)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(33, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(723, 136)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Datos Firma"
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(636, 16)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(81, 21)
        Me.Button4.TabIndex = 15
        Me.Button4.Text = "Salir"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.Button3)
        Me.GroupBox5.Controls.Add(Me.Textfinal)
        Me.GroupBox5.Controls.Add(Me.Textini)
        Me.GroupBox5.Controls.Add(Me.Button1)
        Me.GroupBox5.Location = New System.Drawing.Point(6, 64)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(454, 66)
        Me.GroupBox5.TabIndex = 14
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Buscar Rangos"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(361, 29)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(87, 22)
        Me.Button3.TabIndex = 12
        Me.Button3.Text = "Nueva Busqueda"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Textfinal
        '
        Me.Textfinal.Location = New System.Drawing.Point(139, 31)
        Me.Textfinal.Name = "Textfinal"
        Me.Textfinal.Size = New System.Drawing.Size(100, 20)
        Me.Textfinal.TabIndex = 11
        '
        'Textini
        '
        Me.Textini.Location = New System.Drawing.Point(16, 30)
        Me.Textini.Name = "Textini"
        Me.Textini.Size = New System.Drawing.Size(100, 20)
        Me.Textini.TabIndex = 10
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(272, 29)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 9
        Me.Button1.Text = "Buscar"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Texttotal)
        Me.GroupBox4.Location = New System.Drawing.Point(304, 16)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(156, 42)
        Me.GroupBox4.TabIndex = 13
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Total Firmas"
        '
        'Texttotal
        '
        Me.Texttotal.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Texttotal.Enabled = False
        Me.Texttotal.Location = New System.Drawing.Point(6, 16)
        Me.Texttotal.Name = "Texttotal"
        Me.Texttotal.Size = New System.Drawing.Size(107, 20)
        Me.Texttotal.TabIndex = 0
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Textmax)
        Me.GroupBox3.Location = New System.Drawing.Point(155, 16)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(143, 42)
        Me.GroupBox3.TabIndex = 12
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Rango Maximo"
        '
        'Textmax
        '
        Me.Textmax.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Textmax.Enabled = False
        Me.Textmax.Location = New System.Drawing.Point(15, 16)
        Me.Textmax.Name = "Textmax"
        Me.Textmax.Size = New System.Drawing.Size(100, 20)
        Me.Textmax.TabIndex = 11
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Textmin)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 16)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(143, 42)
        Me.GroupBox2.TabIndex = 11
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Rango Minimo"
        '
        'Textmin
        '
        Me.Textmin.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Textmin.Enabled = False
        Me.Textmin.Location = New System.Drawing.Point(16, 16)
        Me.Textmin.Name = "Textmin"
        Me.Textmin.Size = New System.Drawing.Size(100, 20)
        Me.Textmin.TabIndex = 10
        '
        'Button5
        '
        Me.Button5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.Location = New System.Drawing.Point(245, 538)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(183, 22)
        Me.Button5.TabIndex = 7
        Me.Button5.Text = "Detalle Cuentas Firmas"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'PanProgress
        '
        Me.PanProgress.Controls.Add(Me.ProgressBar1)
        Me.PanProgress.Location = New System.Drawing.Point(214, 285)
        Me.PanProgress.Name = "PanProgress"
        Me.PanProgress.Size = New System.Drawing.Size(379, 44)
        Me.PanProgress.TabIndex = 8
        Me.PanProgress.TabStop = False
        Me.PanProgress.Text = "Procesando..."
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(15, 12)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(358, 24)
        Me.ProgressBar1.TabIndex = 0
        '
        'PictureBox1
        '
        Me.PictureBox1.Location = New System.Drawing.Point(466, 16)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(135, 115)
        Me.PictureBox1.TabIndex = 16
        Me.PictureBox1.TabStop = False
        '
        'FormFirmas
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(796, 573)
        Me.Controls.Add(Me.PanProgress)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GrillasFirmas)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Button2)
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Name = "FormFirmas"
        Me.Text = "Firmas 3.0 -Visual 2005"
        CType(Me.GrillasFirmas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.PanProgress.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents GrillasFirmas As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents Textfinal As System.Windows.Forms.TextBox
    Friend WithEvents Textini As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Texttotal As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Textmax As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Textmin As System.Windows.Forms.TextBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents PanProgress As System.Windows.Forms.GroupBox
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents CheckBox As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox

End Class
