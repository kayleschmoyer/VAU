Imports System.Diagnostics

''' <summary>
''' Simple wrapper around <see cref="EventLog"/> used throughout the application.
''' </summary>
Public Module Logger
    Private ReadOnly EVENT_SOURCE As String = "VASTUpdater"
    Private ReadOnly LOG_NAME As String = "Application"

    Private Sub EnsureSource()
        Try
            If Not EventLog.SourceExists(EVENT_SOURCE) Then
                EventLog.CreateEventSource(EVENT_SOURCE, LOG_NAME)
            End If
        Catch ex As Exception
            Trace.WriteLine($"Failed to create event source: {ex.Message}")
        End Try
    End Sub

    Public Sub Log(message As String, level As LogLevel)
        EnsureSource()
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
            Trace.WriteLine($"Failed to write to event log: {ex.Message} - {message}")
        End Try
    End Sub

    Public Enum LogLevel
        Info
        Warning
        [Error]
    End Enum
End Module
