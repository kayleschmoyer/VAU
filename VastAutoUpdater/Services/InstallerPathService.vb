Imports System.IO
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
    ''' Ensure the update folder exists in ProgramData.
    ''' </summary>
    Public Sub EnsureUpdateFolderExists()
        If Not Directory.Exists(BasePath) Then
            Directory.CreateDirectory(BasePath)
            Logger.Log($"Created update folder: {BasePath}", Logger.LogLevel.Info)
        End If
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
        Catch ex As Exception
            Logger.Log($"Error during installer cleanup: {ex.Message}", Logger.LogLevel.Warning)
        End Try
    End Sub
End Module
