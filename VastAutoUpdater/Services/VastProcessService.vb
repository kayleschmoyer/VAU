Option Strict On

Imports System.Diagnostics

''' <summary>
''' Finds and closes running VAST programs before a patch is applied.
''' Used by the update engine in both interactive and silent modes.
''' Tries a graceful window close first, then kills any process that
''' does not exit in time.
''' </summary>
Public Module VastProcessService

    ''' <summary>
    ''' Process names (without .exe) that must not be running while a
    ''' VAST patch installs.
    ''' </summary>
    Public ReadOnly VastProcessNames As String() = {
        "Vast",
        "VastMaint",
        "VASTRealtimeUploader",
        "VASTReportingCR",
        "CommShop",
        "InventoryUpload"
    }

    Private Const GRACEFUL_CLOSE_WAIT_MS As Integer = 10000
    Private Const KILL_WAIT_MS As Integer = 5000

    ''' <summary>
    ''' Names of VAST programs currently running (with .exe suffix), for
    ''' logging and status display. Empty when none are running.
    ''' </summary>
    Public Function GetRunningVastPrograms() As List(Of String)
        Dim running As New List(Of String)()
        For Each name As String In VastProcessNames
            Dim procs As Process() = Process.GetProcessesByName(name)
            Try
                If procs.Length > 0 Then running.Add(name & ".exe")
            Finally
                For Each p As Process In procs
                    p.Dispose()
                Next
            End Try
        Next
        Return running
    End Function

    ''' <summary>
    ''' Close every running VAST program: ask each main window to close,
    ''' then kill anything still alive after the grace period. Windowless
    ''' processes (services, uploaders) are killed directly.
    ''' Returns True when nothing in the list is left running.
    ''' </summary>
    Public Function CloseVastPrograms() As Boolean
        Dim allClosed As Boolean = True

        For Each name As String In VastProcessNames
            For Each p As Process In Process.GetProcessesByName(name)
                Try
                    Logger.Log($"Closing {name}.exe (PID {p.Id}) before update", Logger.LogLevel.Info)

                    Dim exited As Boolean = False
                    ' Graceful close only works for processes with a window
                    ' we can reach; otherwise go straight to Kill
                    If p.CloseMainWindow() Then
                        exited = p.WaitForExit(GRACEFUL_CLOSE_WAIT_MS)
                    End If

                    If Not exited AndAlso Not p.HasExited Then
                        p.Kill()
                        exited = p.WaitForExit(KILL_WAIT_MS)
                    End If

                    If p.HasExited Then
                        Logger.Log($"{name}.exe closed", Logger.LogLevel.Info)
                    Else
                        Logger.Log($"{name}.exe (PID {p.Id}) is still running after close/kill", Logger.LogLevel.Error)
                        allClosed = False
                    End If
                Catch ex As Exception
                    ' The process may have exited between enumeration and here;
                    ' anything else means we could not stop it
                    Try
                        If Not p.HasExited Then
                            Logger.Log($"Failed to close {name}.exe: {ex.Message}", Logger.LogLevel.Error)
                            allClosed = False
                        End If
                    Catch
                        Logger.Log($"Failed to close {name}.exe: {ex.Message}", Logger.LogLevel.Warning)
                    End Try
                Finally
                    p.Dispose()
                End Try
            Next
        Next

        Return allClosed
    End Function

End Module
