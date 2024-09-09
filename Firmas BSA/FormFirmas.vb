Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Drawing.Imaging
Imports System.IO

Public Class FormFirmas
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
    Public sPathTif As String
    Public Carp As String
    Public Numero As Integer
    Public WinRar As String, WinZip As String
    Public TextRecImagen As Integer
    
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        LLenasGrilla(Textini.Text, Textfinal.Text)
    End Sub
    Public Function LlenaGrillaTodos()
        Dim DataSet As New DataSet
        Dim adaptador As New SqlDataAdapter
        Try

            If Conn.State <> ConnectionState.Closed Then Conn.Close()
            Dim myCommando As New SqlDataAdapter("PRC_FIRMAS_SEL_TRAETODOS", Conn)
            With myCommando
                .SelectCommand.CommandType = CommandType.StoredProcedure
               
            End With
            Conn.Open()
            myCommando.Fill(DataSet, "Materia")

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
    Public Function LLenasGrilla(ByVal rango1 As Integer, ByVal rango2 As Integer)
        Dim DataSet As New DataSet
        Dim adaptador As New SqlDataAdapter
        Try

            If Conn.State <> ConnectionState.Closed Then Conn.Close()
            Dim myCommando As New SqlDataAdapter("PRC_BUSCAR_FIRMAS_RANGOS", Conn)
            With myCommando
                .SelectCommand.CommandType = CommandType.StoredProcedure
                .SelectCommand.Parameters.Add("@RangoCInicio", SqlDbType.Int).Value = rango1
                .SelectCommand.Parameters.Add("@RangoCFinal", SqlDbType.Int).Value = rango2
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

    Public Function TraeRangosCuentas()
        Try
            If Conn.State = ConnectionState.Open Then Conn.Close()

            myCommand = New SqlCommand("PRC_FIRMAS_MAXIMO_MINIMO_RANGOS", Conn)
            myCommand.CommandType = CommandType.StoredProcedure
            Conn.Open()
            myReader = myCommand.ExecuteReader()
            If myReader.Read() Then
                Textmax.Text = myReader("Maximo")
                Textmin.Text = myReader("Minimo")
                myReader.Close()
            End If
            Conn.Close()

            'Return dst
        Catch var As SqlException

        Catch var As Exception

        End Try
    End Function
    Public Function TraeCantidadCuentas()
        Try
            If Conn.State = ConnectionState.Open Then Conn.Close()

            myCommand = New SqlCommand("PRC_FIRMAS_SEL_CANTIDAD_FIRMAS", Conn)
            myCommand.CommandType = CommandType.StoredProcedure
            Conn.Open()
            myReader = myCommand.ExecuteReader()
            If myReader.Read() Then
                Texttotal.Text = myReader("CANTIDADFIRMAS")




                myReader.Close()
            End If
            Conn.Close()

        Catch var As SqlException

        Catch var As Exception

        End Try
    End Function


    Private Sub FormFirmas_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Button5.Enabled = False
        PanProgress.Visible = False
        Conexion()
        LlenaGrillaTodos()
        TraeRangosCuentas()
        TraeCantidadCuentas()
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

    Public Function MostrarPictureBox(ByVal Cuenta As Integer, ByVal Correlativo As Integer)

        TraeTYPERECTEXTCuentas(Cuenta, Correlativo) '// funcion que trae los TextRectext o tipo de archivos jpg,bmp o pcx

        Select Case TextRecImagen
            Case Is = 1 '// Archivo PCX
                Dim NombreFoto2 As String
                Dim SqlSelect As String = "EXEC PRC_UNAIMAGEN_FIRMAS " & Val(Cuenta) & "," & Val(Correlativo) & ""

                NombreFoto2 = Cuenta & "_" & Correlativo
                Dim Command As New SqlCommand(SqlSelect, Conn)
                Conn.Open()
                Dim MyPhoto() As Byte = CType(Command.ExecuteScalar(), Byte())
                MsgBox("Dim MyPhoto() As Byte = CType(Command.ExecuteScalar(), Byte())")

                MsgBox(" MemoryStream(MyPhoto)")


                Dim ms As New IO.MemoryStream(MyPhoto)
                Dim bm As Bitmap = Nothing
                MsgBox("bm = New Bitmap(ms) Y Dim img As Image = bm")
                bm = New Bitmap(ms)
                Dim img As Image = bm
                MsgBox("voy al Try")
                Try
                    MsgBox("Pase!")
                    ' bm = New Bitmap(ms)
                Catch ex As Exception
                    MsgBox(" No Pase!")
                    System.Diagnostics.Debug.WriteLine(ex.Message)
                End Try

                If img IsNot Nothing Then
                    MsgBox("PictureBox1.Image = img")
                    PictureBox1.Image = img
                    PictureBox1.Image.Save(sPathTif & "\" & NombreFoto2 & ".bmp", Imaging.ImageFormat.Bmp)
                    MsgBox("PictureBox1.Image.Save(sPathTif & " \ " & NombreFoto2.jpg")
                End If
            Case Is = 2 '//// Archivo JPG
                Dim NombreFoto2 As String
                Dim SqlSelect As String = "EXEC PRC_UNAIMAGEN_FIRMAS " & Val(Cuenta) & "," & Val(Correlativo) & ""

                NombreFoto2 = Cuenta & "_" & Correlativo
                Dim Command As New SqlCommand(SqlSelect, Conn)
                Conn.Open()
                Dim MyPhoto() As Byte = CType(Command.ExecuteScalar(), Byte())
                MsgBox("Dim MyPhoto() As Byte = CType(Command.ExecuteScalar(), Byte())")

                MsgBox(" MemoryStream(MyPhoto)")
                'Dim img As Image = Bytes2Image(MyPhoto)

                Dim ms As New IO.MemoryStream(MyPhoto)
                Dim bm As Bitmap = Nothing
                MsgBox("bm = New Bitmap(ms) Y Dim img As Image = bm")
                bm = New Bitmap(ms)
                Dim img As Image = bm
                MsgBox("voy al Try")
                Try
                    MsgBox("Pase!")
                    ' bm = New Bitmap(ms)
                Catch ex As Exception
                    MsgBox(" No Pase!")
                    System.Diagnostics.Debug.WriteLine(ex.Message)
                End Try

                If img IsNot Nothing Then
                    MsgBox("PictureBox1.Image = img")
                    PictureBox1.Image = img
                    PictureBox1.Image.Save(sPathTif & "\" & NombreFoto2 & ".jpeg", Imaging.ImageFormat.Jpeg)
                    MsgBox("PictureBox1.Image.Save(sPathTif & " \ " & NombreFoto2.jpg")
                End If
        End Select
    End Function
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click



        CreaCarpetaTemporal()

        Dim Fila As Integer
        Dim Cuenta As Integer
        Dim Correlativo As Integer
   

        For Fila = 0 To GrillasFirmas.RowCount - 1
            'If CType(GrillasFirmas.Rows.Item(Fila).Cells(0).Value, CheckBox).Checked Then
            '    '//Acción seleccionadas
            'Else
            '    '//Acción no seleccionadas
            'End If
            '        'Selec = Ctype(GrillasFirmas.Rows.Item(fila).Cells(0).Controls(1).CheckBox).Checked=true
            '    If (GrillasFirmas.Rows(Fila).Cells(0).Value) = 1 Then
            '        'if  Ctype(me.GrillasFirmas.Rows.Item(fila).Cells(0).Value.GetType.
            '        'If CType(GrillasFirmas.Rows.Item(Fila).Cells(0).GrillasFirmas.Controls(1), CheckBox).Checked = True Then
            '        Exit For
            '    Else
            Cuenta = Val(GrillasFirmas.Rows(Fila).Cells(2).Value)
            Correlativo = Val(GrillasFirmas.Rows(Fila).Cells(3).Value)
            CreaCarpetaTemporal()
            'MostrarPictureBox(Cuenta, Correlativo)
            LLenaCarpetaTemporal(Cuenta, Correlativo)
            'End If


            'End If

        Next
        'End If

        PanProgress.Visible = False
        MsgBox("Se Guardo con exito las imagenes en: C:\CarpetaTemporal", MsgBoxStyle.Information)

    End Sub
    Sub CreaCarpetaTemporal()
        Dim Directorio As String
        '// CREA LA CARPETA TEMPORAL AUTOMATICAMENTE
        Carp = Dir("C:\CarpetaTemporal", vbDirectory)
        If Carp <> "" Then
            sPathTif = "C:\CarpetaTemporal" '// PATH DONDE SE GUARDA LA RUTA DONDE SE GUARDAR LOS ARCHIVOS TEMPORALES QUE SEREAN TRANSFORMADOS
        Else
            Directorio = "C:\CarpetaTemporal"
            sPathTif = "C:\CarpetaTemporal"
            MkDir(Directorio)
            MsgBox("Las Fotos procesadas estaran en la Siguente direccion en: C:\CarpetaTemporal")
        End If
    End Sub
    Public Function LLenaCarpetaTemporal(ByVal Cuenta As Integer, ByVal correlativo As Integer)
        Dim NombreFoto As String
        If Cuenta = 0 Then

        Else
            TraeTYPERECTEXTCuentas(Cuenta, correlativo) '// funcion que trae los TextRectext o tipo de archivos jpg,bmp o pcx

            Select Case TextRecImagen
                Case Is = 1 '// Archivo PCX

                    Dim SqlSelect As String = "EXEC PRC_UNAIMAGEN_FIRMAS " & Val(Cuenta) & "," & Val(correlativo) & ""

                    Dim Command As New SqlCommand(SqlSelect, Conn)
                    Conn.Open()
                    Dim MyPhoto() As Byte = CType(Command.ExecuteScalar(), Byte())
                    Dim ms As New MemoryStream(MyPhoto)
                    'Dim bm As Bitmap = Nothing
                    Dim SideSize As Integer
                    SideSize = 1000
                    Dim bm As New Bitmap(SideSize, SideSize)
                    'Dim g As Graphics = Graphics.FromImage(bm)
                    'g.FillEllipse(New SolidBrush(Color.Red), 0, 0, SideSize, SideSize)
                    'g.DrawLine(New Pen(Color.Black), 0, 0, SideSize, SideSize)
                    'g.DrawLine(New Pen(Color.Black), SideSize, 0, 0, SideSize)
                    'g.Dispose()

                    Try
                        bm = New Bitmap(ms)
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine(ex.Message)
                    End Try
                    


                    'Dim ArchivoTIF As String
                    'ArchivoTIF = sPathTif
                    NombreFoto = Cuenta & "_" & correlativo
                    bm.Save(sPathTif & "\" & NombreFoto & ".bmp") ',' Imaging.ImageFormat.jpg)
                    ProgressBar(Numero)
                    Conn.Close()
                    Return MyPhoto


                Case Is = 2 '// Archivo JPG
                    Dim SqlSelect As String = "EXEC PRC_UNAIMAGEN_FIRMAS " & Val(Cuenta) & "," & Val(correlativo) & ""

                    Dim Command As New SqlCommand(SqlSelect, Conn)
                    Conn.Open()
                    Dim MyPhoto() As Byte = CType(Command.ExecuteScalar(), Byte())
                    Dim ms As New MemoryStream(MyPhoto)
                    Dim bm As Bitmap = Nothing
                    Try
                        bm = New Bitmap(ms)
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine(ex.Message)
                    End Try
                    Dim ArchivoTIF As String
                    ArchivoTIF = sPathTif
                    NombreFoto = Cuenta & "_" & correlativo
                    bm.Save(sPathTif & "\" & NombreFoto & ".jpg", Imaging.ImageFormat.Jpeg)
                    ProgressBar(Numero)
                    Conn.Close()
                    Return MyPhoto


                Case Is = 3 '// Archivo BMP


                    Dim SqlSelect As String = "EXEC PRC_UNAIMAGEN_FIRMAS " & Val(Cuenta) & "," & Val(correlativo) & ""
                    'MsgBox(SqlSelect)
                    Dim Command As New SqlCommand(SqlSelect, Conn)
                    Conn.Open()
                    Dim MyPhoto() As Byte = CType(Command.ExecuteScalar(), Byte())
                    Dim ms As New MemoryStream(MyPhoto)
                    Dim bm As Bitmap = Nothing
                    Try
                        bm = New Bitmap(ms)
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine(ex.Message)
                    End Try
                    Dim ArchivoTIF As String
                    ArchivoTIF = sPathTif
                    NombreFoto = Cuenta & "_" & correlativo
                    bm.Save(sPathTif & "\" & NombreFoto & ".jpg")

                    ProgressBar(Numero)
                    Conn.Close()
                    Return MyPhoto
            End Select




        End If

    End Function
    Public Function ProgressBar(ByVal Num As Integer) As String
        PanProgress.Visible = True
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = 25
        For Numero = 0 To 25
            ProgressBar1.Value = (Numero)

        Next
    End Function

    Public Function ConvertTifTo(ByVal cuenta As Integer, ByVal correlativo As Integer) As String
       
        Dim dimension As Imaging.FrameDimension
        Dim Imagen As Image
        Dim Item As Integer
        Dim item2 As Integer
        Dim ArchivoTIF As String
        Dim Archivoimg As String
        Dim Tipo As Imaging.ImageFormat
        Dim tipos As String
        Dim c As String

        ArchivoTIF = sPathTif '& "\" & cuenta & "_" & correlativo & ".TIF"  '"c:\imagenTiF\20602029.tif"
        Archivoimg = "c:\imagenTiF\"
        tipos = "jpg"
        'c = ConvertTifTo(ArchivoTIF, Archivoimg, Imaging.ImageFormat.Jpeg)
        Try
            'Se carga el archivo TIF a un Image
            Imagen = System.Drawing.Image.FromFile(ArchivoTIF)
            dimension = New Imaging.FrameDimension(Imagen.FrameDimensionsList(0))
            'Se realiza un ciclo para ver todas las imagenes que contiene la dimensión
            For Item = 0 To Imagen.GetFrameCount(dimension) '– 1
                'Se activa la imagen del multitif en Image
                Imagen.SelectActiveFrame(dimension, Item)
                item2 = Item + 1
                'Se Graba la imagen con el mismo nombre del multitiff
                'más correlativo más la extensión del documento
                'Imagen.Save(ArchivoIMG & “_” & Item & “.” & Tipo.ToString, Tipo)
                Imagen.Save(Archivoimg & item2 & "." & Tipo.ToString, Tipo)
            Next
            'Se liberan los recursos
            Imagen.Dispose()
            Imagen = Nothing
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try

    End Function

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        LlenaGrillaTodos()
    End Sub
  

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        My.Forms.FormIngreso.Show()
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Dim Fila As Integer
        

        For Fila = 0 To GrillasFirmas.RowCount - 1
            
            If CheckBox1.Checked = False Then
                'If Val(GrillasFirmas.Rows(Fila).Cells(1).Value) = 1 Then
                '    Exit For
                'Else
                ' GrillasFirmas.Rows.Item(Fila).Cells(0).ch = True
                GrillasFirmas.Rows(Fila).Cells(0).Value = False
                GrillasFirmas.Rows(Fila).Cells(0).Value = 0
                'End If
            Else
                GrillasFirmas.Rows(Fila).Cells(0).Value = True
                GrillasFirmas.Rows(Fila).Cells(0).Value = 1
            End If
        Next
    

       
    End Sub
    Public Function TraeTYPERECTEXTCuentas(ByVal cuenta As Integer, ByVal correlativo As Integer) As String
        Try
            If Conn.State = ConnectionState.Open Then Conn.Close()

            myCommand = New SqlCommand("PRC_FIRMAS_SEL_BUSCAR", Conn)
            myCommand.CommandType = CommandType.StoredProcedure
            myCommand.Parameters.Add("@Cuenta", SqlDbType.Char).Value = cuenta
            myCommand.Parameters.Add("@TipoTexto", SqlDbType.Char).Value = correlativo
            Conn.Open()
            myReader = myCommand.ExecuteReader()
            If myReader.Read() Then
                TextRecImagen = myReader("TypeRecTxt")

                myReader.Close()
            End If
            Conn.Close()

            'Return dst
        Catch var As SqlException

        Catch var As Exception

        End Try
    End Function

End Class
