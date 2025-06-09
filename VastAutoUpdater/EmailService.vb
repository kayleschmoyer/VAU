Imports System.Net.Mail

Public Class EmailService
    Public Sub SendNotification(success As Boolean, details As String)
        Try
            Dim subject As String = If(success, "Update Success", "Update Failure")
            Using msg As New MailMessage()
                msg.From = New MailAddress(ConfigManager.EmailFrom)
                For Each addr In ConfigManager.EmailTo.Split(";"c)
                    If Not String.IsNullOrWhiteSpace(addr) Then msg.To.Add(addr)
                Next
                msg.Subject = subject
                msg.Body = details
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
