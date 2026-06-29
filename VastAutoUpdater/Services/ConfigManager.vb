Imports System.Configuration
Imports System.Security.Cryptography
Imports System.Text

''' <summary>
''' Centralized access to application configuration values from App.config.
''' Provides safe parsing with defaults, and DPAPI encryption support for credentials.
''' On first access, plaintext credentials are automatically encrypted in-place.
''' </summary>
Public Module ConfigManager

    Private Const ENCRYPTED_PREFIX As String = "enc:"

    Public ReadOnly Property SftpHost As String
        Get
            Return GetSetting("SftpHost", "")
        End Get
    End Property

    Public ReadOnly Property SftpUsername As String
        Get
            Return GetProtectedSetting("SftpUsername", "")
        End Get
    End Property

    Public ReadOnly Property SftpPassword As String
        Get
            Return GetProtectedSetting("SftpPassword", "")
        End Get
    End Property

    Public ReadOnly Property SmtpHost As String
        Get
            Return GetSetting("SmtpHost", "")
        End Get
    End Property

    Public ReadOnly Property SmtpPort As Integer
        Get
            Return GetSettingInt("SmtpPort", 587)
        End Get
    End Property

    Public ReadOnly Property SmtpUsername As String
        Get
            Return GetProtectedSetting("SmtpUsername", "")
        End Get
    End Property

    Public ReadOnly Property SmtpPassword As String
        Get
            Return GetProtectedSetting("SmtpPassword", "")
        End Get
    End Property

    Public ReadOnly Property EmailFrom As String
        Get
            Return GetSetting("EmailFrom", "")
        End Get
    End Property

    Public ReadOnly Property EmailTo As String
        Get
            Return GetSetting("EmailTo", "")
        End Get
    End Property

    ''' <summary>
    ''' Safely retrieve a string setting with a default fallback.
    ''' </summary>
    Public Function GetSetting(key As String, defaultValue As String) As String
        Try
            Dim value As String = ConfigurationManager.AppSettings(key)
            Return If(String.IsNullOrEmpty(value), defaultValue, value)
        Catch
            Return defaultValue
        End Try
    End Function

    ''' <summary>
    ''' Save a setting value to App.config.
    ''' </summary>
    Public Sub SaveSetting(key As String, value As String)
        Try
            Dim config As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
            If config.AppSettings.Settings(key) Is Nothing Then
                config.AppSettings.Settings.Add(key, value)
            Else
                config.AppSettings.Settings(key).Value = value
            End If
            config.Save(ConfigurationSaveMode.Modified)
            ConfigurationManager.RefreshSection("appSettings")
        Catch ex As Exception
            Logger.Log($"Failed to save setting '{key}': {ex.Message}", Logger.LogLevel.Warning)
        End Try
    End Sub

    ''' <summary>
    ''' Safely retrieve an integer setting with a default fallback.
    ''' Uses TryParse to avoid exceptions on malformed values.
    ''' </summary>
    Private Function GetSettingInt(key As String, defaultValue As Integer) As Integer
        Try
            Dim raw As String = ConfigurationManager.AppSettings(key)
            Dim result As Integer
            If Integer.TryParse(raw, result) Then
                Return result
            End If
        Catch
        End Try
        Return defaultValue
    End Function

    ''' <summary>
    ''' Retrieve a credential setting, handling DPAPI encryption transparently.
    ''' If the value is plaintext, it works as-is (backward compatible).
    ''' If prefixed with "enc:", it is decrypted via DPAPI (machine scope).
    ''' </summary>
    Private Function GetProtectedSetting(key As String, defaultValue As String) As String
        Try
            Dim raw As String = ConfigurationManager.AppSettings(key)
            If String.IsNullOrEmpty(raw) Then Return defaultValue

            If raw.StartsWith(ENCRYPTED_PREFIX) Then
                ' Decrypt DPAPI-protected value
                Dim encryptedBytes As Byte() = Convert.FromBase64String(raw.Substring(ENCRYPTED_PREFIX.Length))
                Dim decryptedBytes As Byte() = ProtectedData.Unprotect(encryptedBytes, Nothing, DataProtectionScope.LocalMachine)
                Return Encoding.UTF8.GetString(decryptedBytes)
            End If

            ' Plaintext — return as-is for backward compatibility
            Return raw
        Catch ex As Exception
            Logger.Log($"Error reading protected setting '{key}': {ex.Message}", Logger.LogLevel.Error)
            Return defaultValue
        End Try
    End Function

    ''' <summary>
    ''' Encrypt a plaintext credential value using DPAPI (machine scope).
    ''' Returns the "enc:" prefixed base64 string suitable for App.config.
    ''' </summary>
    Private Function EncryptValue(plaintext As String) As String
        Dim plaintextBytes As Byte() = Encoding.UTF8.GetBytes(plaintext)
        Dim encryptedBytes As Byte() = ProtectedData.Protect(plaintextBytes, Nothing, DataProtectionScope.LocalMachine)
        Return ENCRYPTED_PREFIX & Convert.ToBase64String(encryptedBytes)
    End Function

    ''' <summary>
    ''' On first run, encrypt any plaintext credential values in App.config using DPAPI.
    ''' Already-encrypted values (prefixed with "enc:") are left unchanged.
    ''' </summary>
    Public Sub EncryptOnFirstRun()
        Dim credentialKeys As String() = {"SftpUsername", "SftpPassword", "SmtpUsername", "SmtpPassword"}
        Dim needsSave As Boolean = False

        Try
            Dim config As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)

            For Each key As String In credentialKeys
                Dim setting As KeyValueConfigurationElement = config.AppSettings.Settings(key)
                If setting Is Nothing Then Continue For
                Dim raw As String = setting.Value
                If String.IsNullOrEmpty(raw) Then Continue For
                If raw.StartsWith(ENCRYPTED_PREFIX) Then Continue For

                ' Plaintext value found — encrypt it
                setting.Value = EncryptValue(raw)
                needsSave = True
                Logger.Log($"Encrypted credential: {key}", Logger.LogLevel.Info)
            Next

            If needsSave Then
                config.Save(ConfigurationSaveMode.Modified)
                ConfigurationManager.RefreshSection("appSettings")
                Logger.Log("App.config credentials encrypted via DPAPI", Logger.LogLevel.Info)
            End If
        Catch ex As Exception
            Logger.Log($"Failed to encrypt credentials on first run: {ex.Message}", Logger.LogLevel.Warning)
        End Try
    End Sub

End Module
