Imports WinSCP
Imports System.Threading
Imports Ionic.Zip
Imports System.Net.Mail
Imports System.Net
Imports System.IO
Imports System.Text

Module Module1
    ReadOnly fecha As Date = Date.Now
    ReadOnly d As Integer = Day(fecha)
    ReadOnly m As Integer = Month(fecha)
    ReadOnly a As Integer = Year(fecha)

    Dim NomArchResp As String

    Dim th As Thread
    Dim tim As TimeSpan
    Private ReadOnly sig As New TimeSpan(0, 0, 1)

    Dim continuaRel As Boolean = True

    Dim resultado As String = "El respaldo no se completó."

    Dim param As String = ""

    Dim Ruta As String

    Dim txtEscrito As Boolean

    ''' <summary>
    ''' Función que realiza la tarea del respaldo.
    ''' </summary>
    ''' <param name = "NomF"> Ruta del archivo a respaldar. </param>
    ''' <returns> Devuelve el resultado del respaldo. </returns>
    Function SubeRespaldo(ByVal NomF As String) As Boolean
        Dim sesionOpcionFTP As New SessionOptions
        th = New Thread(AddressOf Contador)
        SubeRespaldo = False

        With sesionOpcionFTP 'Las credenciales pueden variar. Esto resulta un ejemplo.
            .Protocol = Protocol.Ftp
            .HostName = ObtenIP()
            .UserName = "Usuario"
            .Password = "Contraseña"
            .PortNumber = 21
            .FtpSecure = FtpSecure.ExplicitSsl
            .SslHostCertificateFingerprint = "Credencial"
        End With

        Dim TransFerOP As New TransferOptions With {
            .TransferMode = TransferMode.Binary
        }

        Console.WriteLine("Preparando para subir: " & NomF)
        Console.WriteLine("Abriendo sesión FTP ...")

        Using sesFTP As New Session
            sesFTP.Open(sesionOpcionFTP)
            Dim TransResultado As TransferOperationResult

            Console.WriteLine("Transferencia en curso ...")
            th.Start() 'Iniciamos contador abriendo el hilo de ejecución.

            TransResultado = sesFTP.PutFiles(NomF, Destino(), False, TransFerOP)
            TransResultado.Check()

            continuaRel = False 'Terminamos contador.
            While th.IsAlive 'Nos aseguramos de haber cerrado el hilo de ejecución.
            End While

            Console.WriteLine("Transferencia terminada.")

            Dim Trans As TransferEventArgs
            For Each Trans In TransResultado.Transfers
                Console.WriteLine("Respaldo completado con éxito.")
                SubeRespaldo = True
            Next
        End Using
    End Function

    ''' <summary>
    ''' Método que procesa el parámetro pasado al iniciar esta aplicación para determinar el tipo de respaldo que se va a realizar.
    ''' </summary>
    Private Sub ProcesaParametro()
        If Environment.GetCommandLineArgs.Length > 1 Then
            param = Environment.GetCommandLineArgs(1)
            NomArchResp = ObtenNombAr()
            Ruta = Origen()
        Else
            NomArchResp = ObtenNombAr()
            Ruta = Origen()
        End If
    End Sub

    ''' <summary>
    ''' Función que obtiene el nombre del archivo a respaldar.
    ''' Puede cambiar con respecto a la fecha y hora del día que se realice el respaldo.
    ''' </summary>
    ''' <returns>
    ''' Nombre del archivo (comprimido) a respaldar.
    ''' </returns>
    Private Function ObtenNombAr() As String
        If param = "Parámetro1" Then
            ObtenNombAr = "RespaldoParámetro1" & a & "-" & ObtenerMes2(m) & ".bak"
        ElseIf param = "Parámetro2" Then
            ObtenNombAr = "RespaldoParámetro2" & d & "-" & m & "-" & a & ".zip"
        Else
            ObtenNombAr = "RespaldoParámetroX" & ObtenerMes(m) & "-" & a & "_Semana" & NumSemana(d) & ".txt"
        End If
    End Function

    ''' <summary>
    ''' Función que obtiene la dirección IP para realizar la transmisión de información mediante el FTP.
    ''' Dependiendo de la locación física del archivo (o carpeta comprimida) se puede cambiar la dirección IP.
    ''' </summary>
    ''' <returns>
    ''' Nombre del archivo (comprimido) a respaldar.
    ''' </returns>
    Private Function ObtenIP() As String
        If param = "Parámetro1" Then
            ObtenIP = "IPParametro1"
        ElseIf param = "Parámetro2" Then
            ObtenIP = "IPParametro2"
        Else
            ObtenIP = "IPParametroX"
        End If
    End Function

    ''' <summary>
    ''' Función que obtiene la ruta del archivo a respaldar.
    ''' </summary>
    ''' <returns> Ruta u origen del archivo a respaldar. </returns>
    Private Function Origen() As String
        If param = "Parámetro1" Then
            Origen = "RutaOrigenP1" & NomArchResp
        ElseIf param = "Parámetro2" Then
            Origen = "RutaOrigenP2"
        Else
            Origen = "RutaOrigenPX"
        End If
    End Function

    ''' <summary>
    ''' Función que obtiene la ruta donde se subirá el archivo.
    ''' </summary>
    ''' <returns> Ruta destino del archivo a respaldar. </returns>
    Private Function Destino() As String
        If param = "Parámetro1" Then
            Destino = "RutaDestinoP1"
        ElseIf param = "Parámetro2" Then
            Destino = "RutaDestinoP2"
        Else
            Destino = "RutaDestinoPX"
        End If

    End Function

    Sub Main()
        ProcesaParametro()
        Console.WriteLine("Inicio de respaldo.")
        Console.WriteLine("Fecha y hora de ejecución: " & fecha)

        Try
            If Directory.Exists(Ruta) Or File.Exists(Ruta) Then
                Console.WriteLine("Carpeta o archivo " & Origen() & " lista para respaldo.")

                'El archivo comprimido se guardará en el directorio donde se encuentra este programa.
                Comprimir(Ruta)

                Dim exito As Boolean = SubeRespaldo(RutaOrigenRespaldo())
                EnviarMail(exito)
                EscribirTxt()
            Else 'Caso en el que no existe la carpeta.
                resultado = "La carpeta " & Ruta & " no existe en pc."
                Console.WriteLine(resultado)
                EnviarMail(False) 'Envía Mail de que no encontró la carpeta a realizar respaldo.
                EscribirTxt()
            End If
        Catch ex As Exception
            resultado = ex.ToString()
            EnviarMail(False) 'Envía Mail de la excepción arrojada.
            EscribirTxt()
        End Try
    End Sub

    ''' <summary>
    ''' Función que determina la ruta del archivo a respaldar
    ''' </summary>
    ''' <returns> Ruta origen del archivo (comprimido) a respaldar. </returns>
    Private Function RutaOrigenRespaldo() As String
        If param = "Parámetro2" Then
            RutaOrigenRespaldo = Environment.CurrentDirectory & "\" & NomArchResp 'Obtiene el archivo comprimido de la carpeta.
        Else
            RutaOrigenRespaldo = Ruta
        End If
    End Function

    ''' <summary>
    ''' Función que realiza la compresión del archivo si este resulta una carpeta de archivos.
    ''' </summary>
    ''' <param name = "Ruta"> Ruta de la carpeta a comprimir. </param>
    ''' <returns> Resultado de la compresión. True en caso de éxito, o false en otro caso. </returns>
    Function Comprimir(ByVal Ruta As String) As Boolean
        If param = "Parámetro2" Then
            Console.WriteLine("Comprimiendo " & Ruta & " ...")

            th = New Thread(AddressOf Contador)
            th.Start() 'Iniciamos contador abriendo el hilo de ejecución.

            Using zip As New ZipFile()
                zip.AddDirectory(Ruta)
                zip.Save(NomArchResp)
                Console.WriteLine("Archivo " & NomArchResp & " creado.")
            End Using

            continuaRel = False 'Terminamos el contador
            While th.IsAlive 'Nos aseguramos de haber cerrado el hilo de ejecución.
            End While
        Else
        End If
        Comprimir = True
    End Function

    ''' <summary>
    ''' Función que obtiene el mes en formato de cadena.
    ''' </summary>
    ''' <param name = "numMes"> Número de mes (1 - 12) </param>
    ''' <returns> Mes en forma de cadena. </returns>
    Private Function ObtenerMes(ByVal numMes As Integer) As String
        If numMes = 1 Then
            Return "Enero"
        ElseIf numMes = 2 Then
            Return "Febrero"
        ElseIf numMes = 3 Then
            Return "Marzo"
        ElseIf numMes = 4 Then
            Return "Abril"
        ElseIf numMes = 5 Then
            Return "Mayo"
        ElseIf numMes = 6 Then
            Return "Junio"
        ElseIf numMes = 7 Then
            Return "Julio"
        ElseIf numMes = 8 Then
            Return "Agosto"
        ElseIf numMes = 9 Then
            Return "Septiembre"
        ElseIf numMes = 10 Then
            Return "Obtubre"
        ElseIf numMes = 11 Then
            Return "Noviembre"
        ElseIf numMes = 12 Then
            Return "Diciembre"
        Else
            Return ""
        End If
    End Function

    ''' <summary>
    ''' Función que obtiene el mes en formato de un número a dos dígitos.
    ''' </summary>
    ''' <param name = "numMes"> Número de mes (1 - 12). </param>
    ''' <returns> Número de mes a dos dígitos. </returns>
    Private Function ObtenerMes2(ByVal numMes As Integer) As String
        If numMes <= 9 Then
            Return "0" & numMes
        Else
            Return numMes
        End If
    End Function

    ''' <summary>
    ''' Función que obtiene el número de la semana del mes.
    ''' </summary>
    ''' <param name = "numDia"> Día del mes. </param>
    ''' <returns> Número de semana (1 - 4). </returns>
    Private Function NumSemana(ByVal numDia As Integer) As Integer
        NumSemana = 0
        If 1 <= numDia And numDia <= 7 Then
            Return 1
        ElseIf 8 <= numDia And numDia <= 14 Then
            Return 2
        ElseIf 15 <= numDia And numDia <= 21 Then
            Return 3
        ElseIf 22 <= numDia And numDia <= 31 Then
            Return 4
        End If

    End Function

    ''' <summary>
    ''' Método que simula un contador de tiempo de procesos que se refleja en terminal.
    ''' </summary>
    Private Sub Contador()
        Dim f1 As Date = Date.Now
        While continuaRel
            Dim f2 As Date = Date.Now
            If Not (Second(f1) = Second(f2)) Then
                'Console.Clear()
                Console.WriteLine(tim)
                f1 = f2
                tim = tim.Add(sig)
            End If
        End While

        'Reiniciamos valores para mantener consistencia al próximo inicio de otro contador.
        tim = New TimeSpan
        continuaRel = True
    End Sub

    ''' <summary>
    ''' Método que envía un correo del resultado de la operación. Devuelve el resultado del envío.
    ''' </summary>
    ''' <param name = "fueExito"> Bandera que indica si la operación se realizó exitosamente y elaborar el mensaje correspondiente. </param>
    Private Sub EnviarMail(ByVal fueExito As Boolean)
        If fueExito Then
            resultado = "El respaldo de " & NomArchResp & " fue llevado a cabo exitosamente."
        Else
            resultado = "El respaldo no se realizó correctamente debido al siguiente evento: " & resultado
        End If
        Try
            Dim bodyMail As String = resultado & "<br>Fecha y hora de la operación: " & fecha
            Dim smtp As New SmtpClient("ServidorSMTP", 0)
            Dim mail As New MailMessage()

            smtp.Credentials = New NetworkCredential("CorreoRemitente", "Contraseña") 'Uso de credenciales para enviar correo.
            smtp.EnableSsl = True

            mail.From = New MailAddress("CorreoRemitente", "NombreAMostrar")
            mail.To.Add("CorreoDestinatario")
            mail.CC.Add("CorreoDestinatario") 'Envío de copia de correo.
            mail.Subject = ObtenAsunto()

            mail.IsBodyHtml = True
            mail.Body = bodyMail

            smtp.Send(mail)
        Catch ex As Exception
            resultado = "Error de envío de correo: " & ex.Message & Chr(13) & Chr(10) & "Resultado del respaldo: " & resultado
            Console.WriteLine(resultado)
            EscribirTxt()
            txtEscrito = True
        End Try
    End Sub

    ''' <summary>
    ''' Función que obtiene el asunto del correo a enviar dependiendo del proceso del respaldo que se realizó.
    ''' </summary>
    ''' <returns> Asunto de correo. </returns>
    Private Function ObtenAsunto() As String
        If param = "Parámetro1" Then
            ObtenAsunto = "RespaldoParámetro1" & d & "-" & ObtenerMes(m) & "-" & a
        ElseIf param = "Parámetro2" Then
            ObtenAsunto = "RespaldoParámetro2 " & d & "-" & ObtenerMes2(m) & "-" & a
        Else
            ObtenAsunto = "RespaldoParámetroX " & NumSemana(d) & "-" & m & "-" & a
        End If
    End Function

    ''' <summary>
    ''' Método que escribe un archivo en formato txt para obtener el resultado del respaldo localmente.
    ''' </summary>
    Sub EscribirTxt()
        If Not txtEscrito Then
            Dim path As String = Environment.CurrentDirectory & "\Log.txt"

            'Crea o sobreescribe el archivo.
            Dim fs As FileStream = File.Create(path)

            'Agrega el texto al archivo.
            Dim info As Byte() = New UTF8Encoding(True).GetBytes("Respaldo " & NomArchResp & Chr(13) & Chr(10) & resultado & Chr(13) & Chr(10) & "Fecha y hora de operación: " & fecha)
            fs.Write(info, 0, info.Length)
            fs.Close()
        End If
    End Sub
End Module