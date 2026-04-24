Imports System.Net.Mail
Imports System.Net
Module ModuloEmail
    Public Sub EnviarReportePorEmail(ByVal destinatario As String, ByVal rutaAdjunto As String, ByVal asunto As String)
        ' --- CONFIGURACIÓN DEL SERVIDOR (Ejemplo con Gmail) ---
        Dim servidor As String = "smtp.gmail.com"
        Dim puerto As Integer = 587
        Dim usuarioEmisor As String = "CORREOELECTRONICO" ' Tu correo
        Dim passEmisor As String = "PASSWORDAPP" ' Tu contraseña de aplicación

        Try
            ' 1. Crear el mensaje
            Dim correo As New MailMessage()
            correo.From = New MailAddress(usuarioEmisor)
            correo.To.Add(destinatario)
            correo.Subject = asunto
            correo.Body = "Buen día, adjunto se envía el reporte generado desde el Sistema ACN LINUX."
            correo.IsBodyHtml = False

            ' 2. Adjuntar el PDF
            If IO.File.Exists(rutaAdjunto) Then
                Dim adjunto As New Attachment(rutaAdjunto)
                correo.Attachments.Add(adjunto)
            End If

            ' 3. Configurar el Cliente SMTP
            Dim cliente As New SmtpClient(servidor)
            cliente.Port = puerto
            cliente.EnableSsl = True
            cliente.Credentials = New NetworkCredential(usuarioEmisor, passEmisor)

            ' 4. Enviar
            cliente.Send(correo)
            MsgBox("El reporte ha sido enviado con éxito a: " & destinatario, MsgBoxStyle.Information, "JosuOs - Email")

        Catch ex As Exception
            MsgBox("Error al enviar el correo: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub
End Module
