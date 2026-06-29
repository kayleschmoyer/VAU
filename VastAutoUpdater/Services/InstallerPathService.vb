Imports System.IO

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
        EnsureUpdateFolderExists()
        Return Path.Combine(BasePath, $"{version}.exe")
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
