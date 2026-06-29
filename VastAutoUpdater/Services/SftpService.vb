Imports Renci.SshNet
Imports Renci.SshNet.Common
Imports System.IO
Imports System.Linq

''' <summary>
''' Handles communication with the remote SFTP server to check and download updates.
''' Uses a single connection per update cycle, verifies host keys, and disconnects properly.
''' </summary>
Public Class SftpService
    Implements IDisposable

    Private ReadOnly host As String = ConfigManager.SftpHost
    Private ReadOnly remoteDir As String = "/VASTAutoInstall/"
    Private Const CONNECTION_TIMEOUT_SECONDS As Integer = 30
    Private Const OPERATION_TIMEOUT_SECONDS As Integer = 300

    Private client As SftpClient = Nothing
    Private disposed As Boolean = False

    ''' <summary>
    ''' Connect to the SFTP server with host key verification.
    ''' Reuses the existing connection if already connected.
    ''' </summary>
    Public Sub Connect(username As String, password As String)
        If client IsNot Nothing AndAlso client.IsConnected Then Return

        ' Dispose stale client before creating a new one
        If client IsNot Nothing Then
            Try
                client.Dispose()
            Catch
            End Try
            client = Nothing
        End If

        client = New SftpClient(host, username, password)
        client.ConnectionInfo.Timeout = TimeSpan.FromSeconds(CONNECTION_TIMEOUT_SECONDS)
        client.OperationTimeout = TimeSpan.FromSeconds(OPERATION_TIMEOUT_SECONDS)

        ' Verify host key to prevent MITM attacks
        AddHandler client.HostKeyReceived, AddressOf ValidateHostKey

        client.Connect()
        Logger.Log($"Connected to SFTP host: {host}", Logger.LogLevel.Info)
    End Sub

    ''' <summary>
    ''' Validate the SSH host key fingerprint.
    ''' On first connection, the fingerprint is stored in App.config.
    ''' On subsequent connections, it must match the stored value.
    ''' </summary>
    Private Sub ValidateHostKey(sender As Object, e As HostKeyEventArgs)
        Dim fingerprint As String = BitConverter.ToString(e.FingerPrint).Replace("-", ":")
        Dim storedFingerprint As String = ConfigManager.GetSetting("SftpHostFingerprint", "")

        If String.IsNullOrEmpty(storedFingerprint) Then
            ' First connection — trust on first use, store the fingerprint
            ConfigManager.SaveSetting("SftpHostFingerprint", fingerprint)
            Logger.Log($"SFTP host key fingerprint stored: {fingerprint}", Logger.LogLevel.Info)
            e.CanTrust = True
        ElseIf storedFingerprint.Equals(fingerprint, StringComparison.OrdinalIgnoreCase) Then
            e.CanTrust = True
        Else
            ' Fingerprint mismatch — potential MITM
            Logger.Log($"SFTP host key mismatch! Expected: {storedFingerprint}, Got: {fingerprint}", Logger.LogLevel.Error)
            e.CanTrust = False
        End If
    End Sub

    ''' <summary>
    ''' Disconnect from the server cleanly.
    ''' </summary>
    Public Sub Disconnect()
        If client IsNot Nothing AndAlso client.IsConnected Then
            Try
                client.Disconnect()
                Logger.Log("SFTP disconnected", Logger.LogLevel.Info)
            Catch ex As Exception
                Logger.Log($"SFTP disconnect error: {ex.Message}", Logger.LogLevel.Warning)
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Query the SFTP server for the latest version matching the given major.minor prefix.
    ''' Returns "0.0.0" if no matching version is found.
    ''' Must call Connect() first.
    ''' </summary>
    Public Function GetLatestVersion(prefix As String) As String
        Dim files = client.ListDirectory(remoteDir).
            Where(Function(f) f.IsRegularFile AndAlso f.Name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))

        Dim versions As New List(Of Version)()
        For Each f In files
            Dim name As String = Path.GetFileNameWithoutExtension(f.Name)
            Dim v As Version = Nothing
            If Version.TryParse(name, v) AndAlso $"{v.Major}.{v.Minor}" = prefix Then
                versions.Add(v)
            End If
        Next

        If versions.Count = 0 Then
            Logger.Log($"No matching version found on SFTP for prefix: {prefix}", Logger.LogLevel.Warning)
            Return "0.0.0"
        End If

        versions.Sort()
        Dim latest As String = versions.Last().ToString()
        Logger.Log($"Latest version available for {prefix}: {latest}", Logger.LogLevel.Info)
        Return latest
    End Function

    ''' <summary>
    ''' Get the size of a remote file in bytes.
    ''' Returns 0 if the file doesn't exist.
    ''' </summary>
    Public Function GetRemoteFileSize(version As String) As Long
        Try
            Dim remotePath As String = $"{remoteDir}{version}.exe"
            Dim attrs = client.GetAttributes(remotePath)
            Return attrs.Size
        Catch
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' Download the installer for the specified version to localPath.
    ''' Returns True on success, False on failure.
    ''' Must call Connect() first.
    ''' </summary>
    Public Function DownloadFile(version As String, localPath As String, progress As Action(Of ULong)) As Boolean
        Dim remotePath As String = $"{remoteDir}{version}.exe"
        Dim tempPath As String = localPath & ".tmp"

        Logger.Log($"Downloading: {remotePath} -> {tempPath}", Logger.LogLevel.Info)

        Try
            Using fs As New FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None)
                client.DownloadFile(remotePath, fs, Sub(d) progress(d))
            End Using

            ' Verify temp file was actually written
            Dim fi As New FileInfo(tempPath)
            If Not fi.Exists OrElse fi.Length = 0 Then
                Logger.Log($"Download produced empty file: {tempPath}", Logger.LogLevel.Error)
                Return False
            End If

            ' Atomic rename: delete target if it exists, then move temp into place
            If File.Exists(localPath) Then File.Delete(localPath)
            File.Move(tempPath, localPath)

            Logger.Log($"Download completed: {localPath} ({fi.Length} bytes)", Logger.LogLevel.Info)
            Return True
        Catch ex As Exception
            ' Clean up partial download on failure
            Try
                If File.Exists(tempPath) Then File.Delete(tempPath)
            Catch
            End Try
            Throw
        End Try
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose
        If Not disposed Then
            Disconnect()
            If client IsNot Nothing Then
                client.Dispose()
                client = Nothing
            End If
            disposed = True
        End If
    End Sub
End Class
