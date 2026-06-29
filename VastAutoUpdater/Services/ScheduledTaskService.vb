Imports System.Diagnostics
Imports System.IO

''' <summary>
''' Encapsulates creation of Windows scheduled tasks for running the updater.
''' Uses schtasks.exe directly (not via cmd.exe) to avoid command injection.
''' </summary>
Public Module ScheduledTaskService
    Private Const TASK_NAME As String = "VASTAutoUpdate"

    ''' <summary>
    ''' Create a weekly scheduled task that runs the updater silently.
    ''' Calls schtasks.exe directly with proper argument quoting.
    ''' </summary>
    Public Sub CreateTask()
        Try
            Dim exePath As String = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                "VASTUpdater", "VastAutoUpdater.exe")

            If Not File.Exists(exePath) Then
                Logger.Log($"VastAutoUpdater.exe not found at: {exePath}", Logger.LogLevel.Error)
                Return
            End If

            ' Build the /tr value with proper quoting for paths with spaces
            Dim taskRun As String = $"""{exePath}"" silent"

            ' Call schtasks.exe directly — no cmd.exe wrapper
            Dim processInfo As New ProcessStartInfo With {
                .FileName = "schtasks.exe",
                .Arguments = $"/create /tn ""{TASK_NAME}"" /tr ""{taskRun}"" /sc weekly /d SUN /st 02:00 /ru SYSTEM /rl HIGHEST /f",
                .UseShellExecute = False,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True,
                .CreateNoWindow = True
            }

            Logger.Log($"Creating scheduled task: {TASK_NAME}", Logger.LogLevel.Info)

            Using proc As Process = Process.Start(processInfo)
                Dim output As String = proc.StandardOutput.ReadToEnd()
                Dim errorOutput As String = proc.StandardError.ReadToEnd()
                proc.WaitForExit()

                If proc.ExitCode = 0 Then
                    Logger.Log($"Scheduled task created successfully. Output: {output.Trim()}", Logger.LogLevel.Info)
                Else
                    Logger.Log($"Failed to create scheduled task. Exit code: {proc.ExitCode}, Error: {errorOutput.Trim()}", Logger.LogLevel.Warning)
                End If
            End Using
        Catch ex As Exception
            Logger.Log($"Error creating scheduled task: {ex.Message}", Logger.LogLevel.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Check if the scheduled task already exists.
    ''' </summary>
    Public Function TaskExists() As Boolean
        Try
            Dim processInfo As New ProcessStartInfo With {
                .FileName = "schtasks.exe",
                .Arguments = $"/query /tn ""{TASK_NAME}""",
                .UseShellExecute = False,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True,
                .CreateNoWindow = True
            }
            Using proc As Process = Process.Start(processInfo)
                proc.StandardOutput.ReadToEnd()
                proc.WaitForExit()
                Return proc.ExitCode = 0
            End Using
        Catch
            Return False
        End Try
    End Function
End Module
