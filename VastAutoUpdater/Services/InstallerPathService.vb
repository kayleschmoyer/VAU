Imports System.IO
Imports System.Security.AccessControl
Imports System.Security.Principal
Imports System.Text.RegularExpressions

''' <summary>
''' Helper methods related to installer storage paths.
''' Manages the %ProgramData%\VASTUpdater\NewPatchInstall directory.
''' </summary>
Public Module InstallerPathService
    Private ReadOnly BasePath As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
        "VASTUpdater", "NewPatchInstall")

    ''' <summary>
    ''' Ensure the update folder exists in ProgramData, writable by all users.
    ''' Whichever account creates a ProgramData folder first owns it; without
    ''' an explicit grant, other accounts (SYSTEM task vs. interactive user)
    ''' get "Access denied" replacing each other's downloaded installers.
    ''' </summary>
    Public Sub EnsureUpdateFolderExists()
        If Not Directory.Exists(BasePath) Then
            Directory.CreateDirectory(BasePath)
            GrantUsersModify(BasePath)
            Logger.Log($"Created update folder: {BasePath}", Logger.LogLevel.Info)
        End If
    End Sub

    ''' <summary>
    ''' Best-effort grant of Modify to BUILTIN\Users on a folder (inherited by
    ''' its files). Failures are logged, never fatal — the installer also sets
    ''' these ACLs at install time.
    ''' </summary>
    Private Sub GrantUsersModify(folderPath As String)
        Try
            Dim dirInfo As New DirectoryInfo(folderPath)
            Dim security As DirectorySecurity = dirInfo.GetAccessControl()
            Dim usersSid As New SecurityIdentifier(WellKnownSidType.BuiltinUsersSid, Nothing)
            security.AddAccessRule(New FileSystemAccessRule(
                usersSid,
                FileSystemRights.Modify Or FileSystemRights.Synchronize,
                InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit,
                PropagationFlags.None,
                AccessControlType.Allow))
            dirInfo.SetAccessControl(security)
        Catch ex As Exception
            Logger.Log($"Could not grant Users write access on '{folderPath}': {ex.Message}", Logger.LogLevel.Warning)
        End Try
    End Sub

    ''' <summary>
    ''' Build the full path for a downloaded installer version.
    ''' </summary>
    Public Function GetInstallPath(version As String) As String
        ' Validate version contains only digits and dots to prevent path traversal
        If String.IsNullOrEmpty(version) OrElse Not Regex.IsMatch(version, "^\d+(\.\d+){1,3}$") Then
            Throw New ArgumentException($"Invalid version format: {version}")
        End If
        EnsureUpdateFolderExists()
        Dim result As String = Path.Combine(BasePath, $"{version}.exe")
        ' Belt-and-suspenders: verify the resolved path is inside BasePath
        If Not Path.GetFullPath(result).StartsWith(Path.GetFullPath(BasePath), StringComparison.OrdinalIgnoreCase) Then
            Throw New InvalidOperationException($"Path traversal detected: {result}")
        End If
        Return result
    End Function

    ''' <summary>
    ''' Remove old installer files, keeping only the current version.
    ''' Prevents ProgramData from accumulating stale installers.
    ''' </summary>
    Public Sub CleanupOldInstallers(currentVersion As String)
        Try
            If Not Directory.Exists(BasePath) Then Return

            Dim currentFile As String = $"{currentVersion}.exe"

            ' Clean up old .exe installers
            For Each filePath In Directory.GetFiles(BasePath, "*.exe")
                Dim fileName As String = Path.GetFileName(filePath)
                If Not fileName.Equals(currentFile, StringComparison.OrdinalIgnoreCase) Then
                    Try
                        File.Delete(filePath)
                        Logger.Log($"Cleaned up old installer: {fileName}", Logger.LogLevel.Info)
                    Catch ex As Exception
                        Logger.Log($"Could not delete old installer '{fileName}': {ex.Message}", Logger.LogLevel.Warning)
                    End Try
                End If
            Next

            ' Clean up stale .tmp partial downloads
            For Each filePath In Directory.GetFiles(BasePath, "*.tmp")
                Try
                    File.Delete(filePath)
                    Logger.Log($"Cleaned up partial download: {Path.GetFileName(filePath)}", Logger.LogLevel.Info)
                Catch ex As Exception
                    Logger.Log($"Could not delete temp file '{Path.GetFileName(filePath)}': {ex.Message}", Logger.LogLevel.Warning)
                End Try
            Next
        Catch ex As Exception
            Logger.Log($"Error during installer cleanup: {ex.Message}", Logger.LogLevel.Warning)
        End Try
    End Sub
End Module
