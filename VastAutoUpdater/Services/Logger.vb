Imports System.Diagnostics
Imports System.IO

''' <summary>
''' Logging wrapper that writes to both Windows Event Log and a fallback file log.
''' If Event Log is unavailable (source not registered, permissions), the file log
''' ensures diagnostics are never lost.
''' </summary>
Public Module Logger
    Private ReadOnly EVENT_SOURCE As String = "VASTUpdater"
    Private ReadOnly LOG_NAME As String = "Application"
    Private ReadOnly LogFilePath As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
        "VASTUpdater", "Logs", "updater.log")
    Private ReadOnly LogLock As New Object()
    Private _eventLogAvailable As Boolean? = Nothing

    Private Sub EnsureLogDirectory()
        Try
            Dim dir = Path.GetDirectoryName(LogFilePath)
            If Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If
        Catch
            ' If we can't create the log directory, we'll fall through to Trace
        End Try
    End Sub

    Private Function IsEventLogAvailable() As Boolean
        If _eventLogAvailable.HasValue Then Return _eventLogAvailable.Value
        Try
            If Not EventLog.SourceExists(EVENT_SOURCE) Then
                EventLog.CreateEventSource(EVENT_SOURCE, LOG_NAME)
            End If
            _eventLogAvailable = True
        Catch ex As Exception
            _eventLogAvailable = False
            WriteToFile($"Event Log source '{EVENT_SOURCE}' unavailable: {ex.Message}", "Warning")
        End Try
        Return _eventLogAvailable.Value
    End Function

    Public Sub Log(message As String, level As LogLevel)
        Dim levelStr As String = level.ToString()

        ' Always write to file log as primary diagnostic record
        WriteToFile(message, levelStr)

        ' Attempt Event Log as secondary
        If IsEventLogAvailable() Then
            Try
                Dim entryType As EventLogEntryType = EventLogEntryType.Information
                Select Case level
                    Case LogLevel.Warning
                        entryType = EventLogEntryType.Warning
                    Case LogLevel.Error
                        entryType = EventLogEntryType.Error
                End Select
                EventLog.WriteEntry(EVENT_SOURCE, message, entryType)
            Catch ex As Exception
                WriteToFile($"Event Log write failed: {ex.Message}", "Warning")
            End Try
        End If
    End Sub

    Private Sub WriteToFile(message As String, level As String)
        Try
            EnsureLogDirectory()
            SyncLock LogLock
                Using writer As New StreamWriter(LogFilePath, True)
                    writer.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}")
                End Using
            End SyncLock
        Catch ex As Exception
            ' Last resort — Trace output
            Trace.WriteLine($"[{level}] {message} (file log failed: {ex.Message})")
        End Try
    End Sub

    ''' <summary>
    ''' Trim log file if it exceeds maxSizeBytes. Keeps the most recent entries.
    ''' Called at startup to prevent unbounded growth.
    ''' </summary>
    Public Sub TrimLogFile(Optional maxSizeBytes As Long = 5242880) ' 5 MB default
        Try
            If Not File.Exists(LogFilePath) Then Return
            Dim fi As New FileInfo(LogFilePath)
            If fi.Length <= maxSizeBytes Then Return

            SyncLock LogLock
                ' Read and write inside the same lock to prevent lost writes
                Dim lines = File.ReadAllLines(LogFilePath)
                Dim keepFrom As Integer = lines.Length \ 2
                File.WriteAllLines(LogFilePath, lines.Skip(keepFrom).ToArray())
            End SyncLock
            Log("Log file trimmed to stay under size limit", LogLevel.Info)
        Catch
            ' Non-critical — don't let log maintenance crash the app
        End Try
    End Sub

    Public Enum LogLevel
        Info
        Warning
        [Error]
    End Enum
End Module
