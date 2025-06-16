Imports System.Diagnostics
Imports System.IO

''' <summary>
''' Encapsulates creation of Windows scheduled tasks for running the updater.
''' </summary>
Public Module ScheduledTaskService
    ''' <summary>
    ''' Attempt to create a weekly scheduled task that runs the updater silently.
    ''' </summary>
    Public Sub CreateTask()
        Try
            Dim exePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "VASTUpdater\VastAutoUpdater.exe")
            If Not File.Exists(exePath) Then
                Logger.Log($"VastAutoUpdater.exe not found at: {exePath}", Logger.LogLevel.Error)
                Return
            End If

            Dim taskCommand As String = $"schtasks /create /tn \"VASTAutoUpdate\" /tr \"\"\"{exePath}\" silent\" /sc weekly /d SUN /st 02:00 /ru SYSTEM /rl HIGHEST /f"
            Logger.Log($"Creating scheduled task with command: {taskCommand}", Logger.LogLevel.Info)

            Dim processInfo As New ProcessStartInfo With {
                .FileName = "cmd.exe",
                .Arguments = "/C " & taskCommand,
                .UseShellExecute = False,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True,
                .CreateNoWindow = True
            }

            Using proc As Process = Process.Start(processInfo)
                Dim output As String = proc.StandardOutput.ReadToEnd()
                Dim errorOutput As String = proc.StandardError.ReadToEnd()
                proc.WaitForExit()

                If proc.ExitCode = 0 Then
                    Logger.Log("Scheduled Task created successfully.", Logger.LogLevel.Info)
                    Logger.Log($"Task creation output: {output}", Logger.LogLevel.Info)
                Else
                    Logger.Log($"Failed to create scheduled task. Exit Code: {proc.ExitCode} - Error: {errorOutput}", Logger.LogLevel.Warning)
                End If
            End Using
        Catch ex As Exception
            Logger.Log($"Error creating scheduled task: {ex.Message}", Logger.LogLevel.Warning)
        End Try
    End Sub
End Module
