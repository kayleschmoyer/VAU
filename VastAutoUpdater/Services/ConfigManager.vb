Imports System.Configuration

''' <summary>
''' Centralized access to application configuration values from App.config.
''' </summary>
Public Module ConfigManager
    Public ReadOnly Property SftpHost As String
        Get
            Return ConfigurationManager.AppSettings("SftpHost")
        End Get
    End Property

    Public ReadOnly Property SftpUsername As String
        Get
            Return ConfigurationManager.AppSettings("SftpUsername")
        End Get
    End Property

    Public ReadOnly Property SftpPassword As String
        Get
            Return ConfigurationManager.AppSettings("SftpPassword")
        End Get
    End Property

    Public ReadOnly Property SmtpHost As String
        Get
            Return ConfigurationManager.AppSettings("SmtpHost")
        End Get
    End Property

    Public ReadOnly Property SmtpPort As Integer
        Get
            Return Integer.Parse(ConfigurationManager.AppSettings("SmtpPort"))
        End Get
    End Property

    Public ReadOnly Property SmtpUsername As String
        Get
            Return ConfigurationManager.AppSettings("SmtpUsername")
        End Get
    End Property

    Public ReadOnly Property SmtpPassword As String
        Get
            Return ConfigurationManager.AppSettings("SmtpPassword")
        End Get
    End Property

    Public ReadOnly Property EmailFrom As String
        Get
            Return ConfigurationManager.AppSettings("EmailFrom")
        End Get
    End Property

    Public ReadOnly Property EmailTo As String
        Get
            Return ConfigurationManager.AppSettings("EmailTo")
        End Get
    End Property

End Module
