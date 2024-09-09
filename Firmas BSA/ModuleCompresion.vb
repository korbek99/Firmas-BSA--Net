Imports System.IO
Imports System.IO.Compression


Module ModuleCompresion

    Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
    Dim WinRar As String, WinZip As String
    Public sPathT2 As String
    Public ArchivoZIPBytes As Byte
    Public Function GetShortPath(ByVal strFileName As String) As String
        Dim Ret As Long, strPath As String, sPathT2 As String
        'strPath = String$(165, 0)
        Ret = GetShortPathName(strFileName, strPath, 164)
        GetShortPath = Left(strPath, Ret)
    End Function
    Public Function Comprimir(ByVal NombreArchivo As String, ByVal RutaArchivo As String)
        On Error Resume Next
        Dim MiObjeto As Object, Ret As String
        MiObjeto = CreateObject("Wscript.Shell")
        Ret = MiObjeto.RegRead("HKEY_CLASSES_ROOT\WinRAR\shell\open\command\")
        WinRar = Left(Ret, InStrRev(Ret, " "))

        Ret = MiObjeto.RegRead("HKEY_CLASSES_ROOT\WinZip\shell\open\command\")
        WinZip = Left(Ret, InStrRev(Ret, " "))
        MiObjeto = Nothing
        
        If WinRar & WinZip = "" Then
            MsgBox("No se encuentra instalado ni WinZip ni WinRar :(")
            End
        End If
        Dim FileZipName As String, PathToCompress As String, DestPath As String, FullPathZip As String
        'Ruta existente en la que se comprimira el archivo
        DestPath = RutaArchivo '"C:\"
        sPathT2 = RutaArchivo
        'nombre que le asignaremos al zip
        FileZipName = NombreArchivo & ".zip" '"ArchivoComprimido.zip"

        'es nesesario utilizar el nombre corto para que funcione
        FullPathZip = GetShortPath(DestPath) & "\" & FileZipName '"\" & NombreArchivo

        'Archivo o carpeta exitente que se quiere comprimir
        PathToCompress = RutaArchivo & "\" & NombreArchivo & ".zip" '"C:\Archivos de programa\Windows NT\Pinball\SOUND1.WAV"
        'nombre corto
        PathToCompress = GetShortPath(PathToCompress) '(PathToCompress)

        If WinRar <> "" Then
            Shell(WinRar & " a " & FullPathZip & " " & PathToCompress, vbHide) ', vbHide
        Else
            Shell(WinZip & " -a " & FullPathZip & " " & PathToCompress, vbHide) ', vbHide
        End If
    End Function

    Public Function Descomprimir()
        Dim FullPathZip As String, FolderToDescompress As String
        'Ruta existente del archivo comprimirdo
        FullPathZip = "C:\ArchivoComprimido.zip"
        'carpeta existente que se va a descomprimir el zip
        FolderToDescompress = "C:\"
        'Combercion a nombre corto
        FullPathZip = GetShortPath(FullPathZip)
        FolderToDescompress = GetShortPath(FolderToDescompress)
        If WinRar <> "" Then
            Shell(WinRar & " e " & FullPathZip & " " & FolderToDescompress) ', vbHide
        Else
            Shell(WinZip & " -e " & FullPathZip & " " & FolderToDescompress) ', vbHide
        End If
    End Function
    Public Function comprimirStandard(ByVal ArchivoOriginal As String, ByVal NomArchivoZIP As String)
        Try

            Dim archivo As FileStream = File.OpenRead(ArchivoOriginal) 'Abrimos y leemos el archivo que esta en el Proyecto

            'creamos los archivos destino para comprimir con GZipStream y DeflateStream respectivamente
            Dim narchivo As FileStream = File.Create(NomArchivoZIP)
            Dim n1archivo As FileStream = File.Create("Deflatedstrim.defleted")

            'Creamos los Archivos de compresion asignandole el archivo origen y el modo de compresion
            Dim archivo_zip As New GZipStream(narchivo, CompressionMode.Compress, True)
            Dim archivo_otro As New DeflateStream(n1archivo, CompressionMode.Compress, True)

            'creamos un bufer para leer fragmentos de bytes del archivo
            Dim tamaño_buffer As Integer = 16384
            Dim buffer(tamaño_buffer) As Byte
            Dim bytesleidos As Integer = 0

            'Escribimos en los archivos de compresion los fragmentos de bytes guardados en el buffer
            While bytesleidos = archivo.Read(buffer, 0, tamaño_buffer) <> -1

                archivo_zip.Write(buffer, 0, tamaño_buffer)
                archivo_otro.Write(buffer, 0, tamaño_buffer)

            End While

            'mostramos el tamoño de los Archivos
            Console.WriteLine("Archivo original = {0} Archivo .Zip = {1} Archivo .deflated = {2}", archivo.Length, narchivo.Length, n1archivo.Length)
            Console.ReadLine()

            'liberamos los recursos, cerrando todos los archivos utilizados
            archivo_zip.Close()
            archivo_otro.Close()

            archivo.Close()
            narchivo.Close()
            n1archivo.Close()

            'el bloque Catch se ejecuta si existe alguna exepcion en el bloque Try
        Catch ex As Exception
            Console.WriteLine(ex.Message) 'enviamos el mensage de la exepcion producida
            Console.ReadLine()
        End Try

        'Antes de Compilar seleccionamos las propiedades del Archivo.doc y le decimos en la propiedad Copy to output Directory= Copy if newer. Y Listo tenemos nuestro Archivo.doc comprimido en la carpeta bin/Debug del proyecto que crea visual studio normalmente en la ruta MisDocumentos/Visual studio 2005/[nombre del proyecto]/[nombre del proyecto]/bin/Debug alli deben encontrar un archivo .zip y otro .deflated los cuales se comprimieron con las clases GZipStream y DeflateStream respectivamente.
        'el Archivo .zip lo podemos descomprimir con Winzip o Winrar.
    End Function


End Module
