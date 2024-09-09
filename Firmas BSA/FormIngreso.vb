Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Drawing.Imaging
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.IO.Compression
Public Class FormIngreso
    ' Array extensiones para los archivos gráficos
    Private sExtension() As String = {"*.jpg", "*.bmp", "*.png", _
                                      "*.ico", "*.gif", "*.wmf", _
                                      "*.jpeg", "*.tif", "*.psd"}

    ' Estrucutura para usar con la función SHGetFileInfo _
    ' y recuperar el ícono asociado al archivo
    Private Structure SHFILEINFO
        Public hIcon As IntPtr            ' : icon
        Public iIcon As Integer           ' : icondex
        Public dwAttributes As Integer    ' : SFGAO_ flags
         _
        Public szDisplayName As String
         _
        Public szTypeName As String
    End Structure

    Private Declare Auto Function SHGetFileInfo Lib "shell32.dll" _
                (ByVal pszPath As String, _
                 ByVal dwFileAttributes As Integer, _
                 ByRef psfi As SHFILEINFO, _
                 ByVal cbFileInfo As Integer, _
                 ByVal uFlags As Integer) As IntPtr

    ' Constantes para SHGetFileInfo
    Private Const SHGFI_ICON = &H100
    Private Const SHGFI_SMALLICON = &H1
    Private Const SHGFI_LARGEICON = &H0    ' Large icon
    Private Const MAX_PATH = 260

    ' Cargar los íconos en el imagelist para el Listview
    Private Sub cargar_iconos_de_Archivos( _
        ByVal sPath As String, _
        ByVal ImageList As ImageList)


        On Error Resume Next
        Dim shInfo As SHFILEINFO
        shInfo = New SHFILEINFO()

        ' buffers
        With shInfo
            .szDisplayName = New String(vbNullChar, MAX_PATH)
            .szTypeName = New String(vbNullChar, 80)
        End With
        Dim hIcon As IntPtr

        ' recuperar el handle de la imagen
        hIcon = SHGetFileInfo( _
                        sPath, _
                        0, _
                        shInfo, _
                        Marshal.SizeOf(shInfo), _
                        SHGFI_ICON Or SHGFI_SMALLICON)

        ' crear el ícono y agregarlo al ImageList
        Dim MyIcon As Drawing.Bitmap
        MyIcon = Drawing.Icon.FromHandle(shInfo.hIcon).ToBitmap

        With ImageList
            .Images.Add(sPath.ToString(), MyIcon)
        End With

        On Error GoTo 0

    End Sub

    Dim cmd As New SqlCommand
    Dim Reader As SqlDataReader
    Dim numeroEstado As Integer
    'Dim Conn As New SqlConnection(ConfigurationSettings.AppSettings("ConsultaCDE"))
    Dim Estado As Integer
    Public intRetorno As Integer
    Public myReader, myReader2 As SqlDataReader
    Public myCommand, myCommand2 As SqlCommand
    Public myParam, myParam2, myParamReturn, myParamReturn2 As SqlParameter
    Public dt As New System.Data.DataTable()
    Public dr As DataRow
    Public MyPhoto() As Byte
    Public RutaFoto As String
    Public WinRar As String, WinZip As String
    Public ArchivoTEMP As String
    Public TipoImagen As Integer
    Public arrImage() As Byte
    Public RutaArchivoIMagen As String
    Public TYpeRecText As Integer


    Public Function LlenaGrillaTodos()
        Dim DataSet As New DataSet
        Dim adaptador As New SqlDataAdapter
        Try

            If Conn.State <> ConnectionState.Closed Then Conn.Close()
            Dim myCommando As New SqlDataAdapter("PRC_FIRMAS_SEL_TRAETODOS", Conn)
            With myCommando
                .SelectCommand.CommandType = CommandType.StoredProcedure
                '.SelectCommand.Parameters.Add("@RangoCInicio", SqlDbType.Int).Value = rango1
                '.SelectCommand.Parameters.Add("@RangoCFinal", SqlDbType.Int).Value = rango2
            End With
            Conn.Open()
            myCommando.Fill(DataSet, "Materia")
            'Adapter.Fill(DataSet, "Materia")
            GrillasFirmas.DataSource = DataSet
            GrillasFirmas.DataMember = "Materia"
            GrillasFirmas.Refresh()

            Conn.Close()
            Return DataSet
        Catch var As SqlException

        Catch var As Exception
            'Return 0
        End Try
    End Function
    Private Sub FormInicio_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LlenaGrillaTodos()
        HabilitarNuevo()
        GroupTipo.Visible = False
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        If CheckBox1.Checked = True Then
            'BuscarPorCuentayCorrelativo(Val(Textcta.Text), Val(Textcorr.Text))
            BuscarImagenDeCuenta(Val(Textcta.Text), Val(Textcorr.Text))
            HabilitarBuscar()
        Else
            BuscarCuenta(Val(Textcta.Text))
            HabilitarBuscar()
        End If
    End Sub
    Public Function BuscarCuenta(ByVal Cta As Integer)
        Dim DataSet As New DataSet
        Dim adaptador As New SqlDataAdapter
        Try

            If Conn.State <> ConnectionState.Closed Then Conn.Close()
            Dim myCommando As New SqlDataAdapter("PRC_FIRMAS_SEL_FirmasPorCuentas", Conn)
            With myCommando
                .SelectCommand.CommandType = CommandType.StoredProcedure
                .SelectCommand.Parameters.Add("@Cuenta", SqlDbType.Int).Value = Cta

            End With
            Conn.Open()
            myCommando.Fill(DataSet, "Materia")
            'Adapter.Fill(DataSet, "Materia")
            GrillasFirmas.DataSource = DataSet
            GrillasFirmas.DataMember = "Materia"
            GrillasFirmas.Refresh()

            Conn.Close()
            Return DataSet
        Catch var As SqlException

        Catch var As Exception
            'Return 0
        End Try

    End Function
    Public Function BuscarPorCuentayCorrelativo(ByVal Cta As Integer, ByVal Correlativo As Integer)
        Dim DataSet As New DataSet
        Dim adaptador As New SqlDataAdapter
        Try

            If Conn.State <> ConnectionState.Closed Then Conn.Close()
            Dim myCommando As New SqlDataAdapter("PRC_BUSCAR_FIRMAS_RANGOS", Conn)
            With myCommando
                .SelectCommand.CommandType = CommandType.StoredProcedure
                .SelectCommand.Parameters.Add("@Cuenta", SqlDbType.Int).Value = Cta
                .SelectCommand.Parameters.Add("@TipoTexto", SqlDbType.Int).Value = Correlativo
            End With
            Conn.Open()
            myCommando.Fill(DataSet, "Materia")
            'Adapter.Fill(DataSet, "Materia")
            GrillasFirmas.DataSource = DataSet
            GrillasFirmas.DataMember = "Materia"
            GrillasFirmas.Refresh()

            Conn.Close()
            Return DataSet
        Catch var As SqlException

        Catch var As Exception
            'Return 0
        End Try

    End Function
    Public Function BuscarImagenDeCuenta(ByVal Cta As Integer, ByVal Correlativo As Integer)
        Dim ms As New MemoryStream(ExtraerImagen(Val(Textcta.Text), Val(Textcorr.Text)))
        Me.PictureBox1.Image = Image.FromStream(ms)

        'Try
        '    If Conn.State = ConnectionState.Open Then Conn.Close()

        '    myCommand = New SqlCommand("PRC_FIRMAS_SEL_BUSCAR", Conn)
        '    myCommand.CommandType = CommandType.StoredProcedure
        '    Conn.Open()
        '    myReader = myCommand.ExecuteReader()
        '    If myReader.Read() Then

        '        Textcta.Text = myReader("Cuenta")
        '        Textcorr.Text = myReader("TipoTexto")
        '        'dr("Foto") = TablaNavegar.Image2Bytes(Me.PictureBox1.Image)
        '        'dr["Foto"] = TablaNavegar.Image2Bytes(this.PictureBox1.Image);


        '        'Textcorr.Text = myReader("Imagen") '// IMAGEN DEL SQL SERVER




        '        myReader.Close()
        '    End If
        '    Conn.Close()

        '    'Return dst
        'Catch var As SqlException

        'Catch var As Exception

        'End Try
    End Function
    Public Shared Function Image2Bytes(ByVal img As Image) As Byte()
        Dim sTemp As String = Path.GetTempFileName()
        Dim fs As New IO.FileStream(sTemp, FileMode.OpenOrCreate, FileAccess.ReadWrite)
        img.Save(fs, System.Drawing.Imaging.ImageFormat.Png)
        fs.Position = 0
        '
        Dim imgLength As Integer = CInt(fs.Length)
        Dim bytes(0 To imgLength - 1) As Byte
        fs.Read(bytes, 0, imgLength)
        fs.Close()
        Return bytes
    End Function
   
  
    Function ExtraerImagen(ByVal Cuenta As Integer, ByVal correlativo As Integer) As Byte()
        Dim SqlSelect As String = "SELECT Imagen FROM Textos WHERE Cuenta=" & Val(Cuenta) & "" _
            & " and  TipoTexto =" & Val(correlativo) & "" _
            & " and tz_lock=0"
        Dim Command As New SqlCommand(SqlSelect, Conn)
        Conn.Open()
        Dim MyPhoto() As Byte = CType(Command.ExecuteScalar(), Byte())
        Conn.Close()
        Return MyPhoto
    End Function
    

    Private Sub ButtonAdjuntar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonAdjuntar.Click
        GroupTipo.Visible = True


        Select Case TipoImagen


            Case 1
                With OpenFileDialog1
                    .InitialDirectory = ""
                    .Filter = "Todos los Archivos|*.*|PCX|*.pcx"
                    .FilterIndex = 2
                End With
                If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    With PictureBox1
                        '.Image = Image.FromFile(OpenFileDialog1.FileName)
                        '.SizeMode = PictureBoxSizeMode.CenterImage
                        '.BorderStyle = BorderStyle.Fixed3D

                        'Me.BtnInsertar.Enabled = True
                        RutaFoto = OpenFileDialog1.FileName
                    End With


                End If

            Case 2
                With OpenFileDialog1
                    .InitialDirectory = ""
                    .Filter = "Todos los Archivos|*.*|JPG|*.jpg|GIFs|*.gif|Bitmaps|*.bmp|*.bmp|"
                    .FilterIndex = 2
                End With
                If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    With PictureBox1
                        .Image = Image.FromFile(OpenFileDialog1.FileName)
                        .SizeMode = PictureBoxSizeMode.CenterImage
                        .BorderStyle = BorderStyle.Fixed3D

                        'Me.BtnInsertar.Enabled = True
                        RutaFoto = OpenFileDialog1.FileName
                    End With


                End If
            Case 3

                With OpenFileDialog1
                    .InitialDirectory = ""
                    .Filter = "Todos los Archivos|*.*|TIFF|*.Tiff|"
                    .FilterIndex = 2
                End With
                If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    With PictureBox1
                        .Image = Image.FromFile(OpenFileDialog1.FileName)
                        .SizeMode = PictureBoxSizeMode.CenterImage
                        .BorderStyle = BorderStyle.Fixed3D

                        'Me.BtnInsertar.Enabled = True
                        RutaFoto = OpenFileDialog1.FileName
                    End With


                End If
        End Select
    End Sub

    Private Sub GrillasFirmas_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GrillasFirmas.CellContentClick
        Try
            Textcta.Text = Val(GrillasFirmas.Rows(GrillasFirmas.SelectionMode).Cells(1).Value)
            'Cuenta = Val(GrillasFirmas.Rows(Fila).Cells(2).Value)
            Textcorr.Text = Val(GrillasFirmas.Rows(GrillasFirmas.SelectionMode).Cells(2).Value)
            ''Me.txtDescripcion.Text = CStr(Me.GrillasFirmas.SelectedCells(2).Value)
            ''El MemoryStream nos permite crear un almacen de memoria..
            'Dim ms As New MemoryStream(ExtraerImagen(Val(Textcta.Text), Val(Textcorr.Text)))
            'Me.PictureBox1.Image = Image.FromStream(ms)
        Catch ex As Exception
        End Try

    End Sub
    Sub HabilitarNuevo()
        Me.ButtonNuevo.Enabled = True
        Me.ButtonGrabar.Enabled = True
        Me.ButtonMod.Enabled = False
        Me.ButtonEliminar.Enabled = False
        Me.ButtonAdjuntar.Enabled = True
    End Sub
    Sub HabilitarBuscar()
        Me.ButtonNuevo.Enabled = True
        Me.ButtonGrabar.Enabled = False
        Me.ButtonMod.Enabled = True
        Me.ButtonEliminar.Enabled = True
        Me.ButtonAdjuntar.Enabled = True
    End Sub
    Private Sub ButtonNuevo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNuevo.Click
        Textcta.Text = ""
        Textcorr.Text = ""
        HabilitarNuevo()
    End Sub
  

    Private Sub ButtonGrabar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonGrabar.Click

       
        Dim bm As Bitmap = Nothing
        Try
            'TraeCorrelativoCuenta(Val(Textcta.Text))
            Conn.Open()
            If RutaArchivoIMagen = "" Then
                MessageBox.Show("Seleccione una Imagen antes de grabar")
            Else


                'RutaArchivoIMagen = OpenFileDialog1.FileName
                Dim fs As New FileStream(RutaArchivoIMagen, FileMode.Open, FileAccess.Read)
                ' Creamos un array de bytes, cuyo límite superior se corresponderá
                ' con la longitud en bytes de la secuencia.
                Dim data() As Byte = New Byte(Convert.ToInt32(fs.Length)) {}
                ' Al leer la secuencia, se rellenará la matriz.
                fs.Read(data, 0, Convert.ToInt32(fs.Length))

                ' Cerramos la secuencia.
                '
                fs.Close()
                ' Devolvemos el array de bytes

                'Dim arrFilename() As String = Split(Text, "\")
                'Array.Reverse(arrFilename)
                'Dim ms As New IO.MemoryStream

                'PictureBox1.Image.Save(ms, PictureBox1.Image.RawFormat)

                'Dim arrImage() As Byte = ms.GetBuffer




                If Conn.State = ConnectionState.Open Then Conn.Close()
                cmd = New SqlCommand("PRC_ING_FIRMAS", Conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("@tz_lock", SqlDbType.Int).Value = 0
                cmd.Parameters.Add("@Cuenta", SqlDbType.Int).Value = Textcta.Text
                cmd.Parameters.Add("@TipoTexto", SqlDbType.Int).Value = Textcorr.Text
                cmd.Parameters.Add("@Imagen", SqlDbType.VarBinary).Value = data 'arrImage
                cmd.Parameters.Add("@TypeRecTxt", SqlDbType.Char).Value = TYpeRecText
                cmd.Parameters.Add("@Aud_Usuario", SqlDbType.Char).Value = xusuario
                cmd.Parameters.Add("@Aud_WorkStation", SqlDbType.Char).Value = "WORK"
                cmd.Parameters.Add("@Aud_Aplicacion", SqlDbType.Char).Value = "aud "
                cmd.Parameters.Add("@FechaAlta", SqlDbType.DateTime).Value = DateTime.Today
                cmd.Parameters.Add("@Ingresador", SqlDbType.Char).Value = xusuario
                cmd.Parameters.Add("@Categoria", SqlDbType.Char).Value = "A"


                Conn.Open()
                cmd.ExecuteNonQuery()
                'Reader = cmd.ExecuteReader()
                Conn.Close()
            End If
        Catch var As SqlException
            'Return 0
        Catch var As Exception
            'Return 0

        End Try
        LlenaGrillaTodos()
    End Sub

    Private Sub ButtonMod_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonMod.Click


        Dim arrFilename() As String = Split(Text, "\")
        Array.Reverse(arrFilename)
        Dim ms As New MemoryStream
        PictureBox1.Image.Save(ms, PictureBox1.Image.RawFormat)
        Dim arrImage() As Byte = ms.GetBuffer


        Try




            If Conn.State = ConnectionState.Open Then Conn.Close()
            cmd = New SqlCommand("PRC_UPD_FIRMAS", Conn)
            cmd.Parameters.Add("@tz_lock", SqlDbType.Int).Value = 0
            cmd.Parameters.Add("@Cuenta", SqlDbType.Int).Value = Textcta.Text
            cmd.Parameters.Add("@TipoTexto", SqlDbType.Int).Value = Textcorr.Text
            cmd.Parameters.Add("@Imagen", SqlDbType.VarBinary).Value = arrImage
            cmd.Parameters.Add("@TypeRecTxt", SqlDbType.Char).Value = TYpeRecText
            cmd.Parameters.Add("@Aud_Usuario", SqlDbType.Char).Value = xusuario
            cmd.Parameters.Add("@Aud_WorkStation", SqlDbType.Char).Value = "WORK"
            cmd.Parameters.Add("@Aud_Aplicacion", SqlDbType.Char).Value = "aud "
            cmd.Parameters.Add("@FechaAlta", SqlDbType.DateTime).Value = DateTime.Today
            cmd.Parameters.Add("@Ingresador", SqlDbType.Char).Value = xusuario
            cmd.Parameters.Add("@Categoria", SqlDbType.Char).Value = "A"

            Conn.Open()
            cmd.ExecuteNonQuery()
            'Reader = cmd.ExecuteReader()
            Conn.Close()

        Catch var As SqlException
            'Return 0
        Catch var As Exception
            'Return 0
        End Try
        LlenaGrillaTodos()
    End Sub
    Public Function TraeCorrelativoCuenta(ByVal Cuenta As Integer)
        Try
            If Conn.State = ConnectionState.Open Then Conn.Close()

            myCommand = New SqlCommand("PRC_FIRMAS_SEL_FirmasPorCuentas", Conn)
            myCommand.CommandType = CommandType.StoredProcedure
            myCommand.Parameters.Add("@Cuenta", SqlDbType.Char).Value = Cuenta
            Conn.Open()
            myReader = myCommand.ExecuteReader()
            If myReader.Read() Then
                Textcorr.Text = myReader("TipoTexto")
                myReader.Close()
            Else
                Textcorr.Text = 1

            End If
            Conn.Close()

        Catch var As SqlException

        Catch var As Exception

        End Try
    End Function
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    
    Private Sub Textcta_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Textcta.LostFocus
        TraeCorrelativoCuenta(Val(Textcta.Text))
    End Sub

    Public Shared Function Bytes2Image(ByVal bytes() As Byte) As Image
        If bytes Is Nothing Then Return Nothing
        '
        Dim ms As New MemoryStream(bytes)
        Dim bm As Bitmap = Nothing
        Try
            bm = New Bitmap(ms)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try
        Return bm
    End Function

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged

        If RadioButton1.Checked = True Then
            TipoImagen = 1
            TYpeRecText = 1
            GroupTipo.Visible = False
            With OpenFileDialog1
                .InitialDirectory = ""
                .Filter = "Todos los Archivos|*.*|PCX|*.pcx"
                .FilterIndex = 2
            End With
            If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                With PictureBox1
                    RutaArchivoIMagen = OpenFileDialog1.FileName
                    Dim fs As New FileStream(RutaArchivoIMagen, FileMode.Open, FileAccess.Read)
                    ' Creamos un array de bytes, cuyo límite superior se corresponderá
                    ' con la longitud en bytes de la secuencia.
                    Dim data() As Byte = New Byte(Convert.ToInt32(fs.Length)) {}
                    ' Al leer la secuencia, se rellenará la matriz.
                    fs.Read(data, 0, Convert.ToInt32(fs.Length))
                    fs.Close()


                    'Dim Imagen As System.Drawing.Image
                    '.Image = Imagen.FromFile(OpenFileDialog1.FileName)
                    '.SizeMode = PictureBoxSizeMode.CenterImage
                    '.BorderStyle = BorderStyle.Fixed3D

                    ''Me.BtnInsertar.Enabled = True
                    MsgBox("Memoria Insuficiente para Mostrar Imagen por ser Formato .PCX")
                    RutaFoto = OpenFileDialog1.FileName
                End With


            End If
        End If

    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        TipoImagen = 2
        TYpeRecText = 2
        GroupTipo.Visible = False
        With OpenFileDialog1
            .InitialDirectory = ""
            .Filter = "Todos los Archivos|*.*|JPG|*.jpg|GIFs|*.gif|Bitmaps|*.bmp|*.bmp"
            .FilterIndex = 2
        End With
        RutaArchivoIMagen = OpenFileDialog1.FileName
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            With PictureBox1
                .Image = Image.FromFile(RutaArchivoIMagen)
                .SizeMode = PictureBoxSizeMode.CenterImage
                .BorderStyle = BorderStyle.Fixed3D

                'Me.BtnInsertar.Enabled = True
                RutaFoto = OpenFileDialog1.FileName
            End With


        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        TipoImagen = 3
        TYpeRecText = 3
        GroupTipo.Visible = False
        With OpenFileDialog1
            .InitialDirectory = ""
            .Filter = "Todos los Archivos|*.*|TIFF|*.Tiff"
            .FilterIndex = 2
        End With
        RutaArchivoIMagen = OpenFileDialog1.FileName
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            With PictureBox1
                'Dim stream As IO.StreamReader = New IO.StreamReader(OpenFileDialog1.FileName)
                'Dim im As Image = Image.FromStream(stream.BaseStream)
                .ImageLocation = RutaArchivoIMagen 'im
                .SizeMode = PictureBoxSizeMode.CenterImage
                .BorderStyle = BorderStyle.Fixed3D

                'Me.BtnInsertar.Enabled = True
                RutaFoto = OpenFileDialog1.FileName
            End With


        End If
    End Sub
End Class


 
