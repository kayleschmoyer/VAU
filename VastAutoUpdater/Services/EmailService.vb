Imports System.Net.Mail

''' <summary>
''' Service for sending SMTP notification emails.
''' Credentials and host information are read from <see cref="ConfigManager"/>.
''' </summary>
Public Class EmailService

    ''' <summary>
    ''' Send an email summarizing the update process result.
    ''' </summary>
    ''' <param name="success">Whether the update succeeded.</param>
    ''' <param name="details">Extra details about the update.</param>
    ''' <param name="exception">Optional exception if the update failed.</param>
    Public Sub SendSummary(success As Boolean, details As String, Optional exception As Exception = Nothing)
        Try
            Dim subject As String = If(success, "Update Success", "Update Failure")
            Dim body As String = $"Time: {DateTime.Now}\nResult: {subject}\nDetails: {details}"
            If Not success AndAlso exception IsNot Nothing Then
                body &= $"\nException: {exception.Message}"
            End If

            Using msg As New MailMessage()
                msg.From = New MailAddress(ConfigManager.EmailFrom)
                For Each addr In ConfigManager.EmailTo.Split(";"c)
                    If Not String.IsNullOrWhiteSpace(addr) Then msg.To.Add(addr)
                Next
                msg.Subject = subject
                msg.Body = body

                Using client As New SmtpClient(ConfigManager.SmtpHost, ConfigManager.SmtpPort)
                    client.Credentials = New Net.NetworkCredential(ConfigManager.SmtpUsername, ConfigManager.SmtpPassword)
                    client.EnableSsl = True
                    client.Send(msg)
                End Using
            End Using
            Logger.Log($"Notification email sent: {subject}", Logger.LogLevel.Info)
        Catch ex As Exception
            Logger.Log($"Failed to send notification email: {ex.Message}", Logger.LogLevel.Error)
        End Try
    End Sub
End Class
