Imports System.Net.Mail

''' <summary>
''' Service for sending SMTP notification emails.
''' Credentials and host information are read from <see cref="ConfigManager"/>.
''' All exceptions are caught and logged — email failures never crash the app.
''' </summary>
Public Class EmailService
    Implements IEmailService

    ''' <summary>
    ''' Send an email summarizing the update process result.
    ''' </summary>
    Public Sub SendSummary(success As Boolean, details As String, Optional exception As Exception = Nothing) Implements IEmailService.SendSummary
        Try
            Dim emailTo As String = ConfigManager.EmailTo
            Dim emailFrom As String = ConfigManager.EmailFrom
            Dim smtpHost As String = ConfigManager.SmtpHost

            Dim smtpUser As String = ConfigManager.SmtpUsername
            Dim smtpPass As String = ConfigManager.SmtpPassword

            ' Validate required email config before attempting send
            If String.IsNullOrWhiteSpace(emailTo) OrElse
               String.IsNullOrWhiteSpace(emailFrom) OrElse
               String.IsNullOrWhiteSpace(smtpHost) OrElse
               String.IsNullOrWhiteSpace(smtpUser) OrElse
               String.IsNullOrWhiteSpace(smtpPass) Then
                Logger.Log("Email configuration incomplete — skipping notification", Logger.LogLevel.Warning)
                Return
            End If

            Dim subject As String = If(success, "VAST Update Success", "VAST Update Failure")
            Dim body As String = $"Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}" & vbCrLf &
                                 $"Machine: {Environment.MachineName}" & vbCrLf &
                                 $"Result: {subject}" & vbCrLf &
                                 $"Details: {details}"

            If Not success AndAlso exception IsNot Nothing Then
                body &= vbCrLf & $"Exception: {exception.Message}"
                If exception.StackTrace IsNot Nothing Then
                    body &= vbCrLf & $"Stack Trace: {exception.StackTrace}"
                End If
            End If

            Using msg As New MailMessage()
                msg.From = New MailAddress(emailFrom)
                For Each addr As String In emailTo.Split(";"c)
                    Dim trimmed As String = addr.Trim()
                    If Not String.IsNullOrWhiteSpace(trimmed) Then
                        msg.To.Add(trimmed)
                    End If
                Next
                If msg.To.Count = 0 Then
                    Logger.Log("No valid email recipients — skipping notification", Logger.LogLevel.Warning)
                    Return
                End If

                msg.Subject = subject
                msg.Body = body

                Using client As New SmtpClient(smtpHost, ConfigManager.SmtpPort)
                    client.Credentials = New Net.NetworkCredential(smtpUser, smtpPass)
                    client.EnableSsl = True
                    client.Timeout = 30000 ' 30 second timeout
                    client.Send(msg)
                End Using
       