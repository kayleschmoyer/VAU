Option Strict On

Imports System.Net.Http
Imports System.Text
Imports Microsoft.Win32

''' <summary>
''' Service for reporting status events to the VAU-Dashboard API.
''' Sends HTTP POST requests to the configured endpoint with an API key header.
''' All sends are fire-and-forget: failures are logged and never block or
''' fail the update workflow. If the dashboard is not configured, events
''' are silently skipped.
''' </summary>
Public Class DashboardService
    Implements IDashboardService

    ''' <summary>
    ''' Shared HttpClient — reused across requests to avoid socket exhaustion.
    ''' </summary>
    Private Shared ReadOnly Client As New HttpClient() With {
        .Timeout = TimeSpan.FromSeconds(15)
    }

    ''' <summary>
    ''' Report a status event to the dashboard without blocking the caller.
    ''' </summary>
    Public Sub ReportEvent(eventType As DashboardEventType,
                           Optional version As String = "",
                           Optional targetVersion As String = "",
                           Optional message As String = "") Implements IDashboardService.ReportEvent
        Try
            Dim apiUrl As String = ConfigManager.DashboardApiUrl
            Dim apiKey As String = ConfigManager.DashboardApiKey

            If String.IsNullOrWhiteSpace(apiUrl) OrElse String.IsNullOrWhiteSpace(apiKey) Then
                Logger.Log("Dashboard not configured — skipping status report", Logger.LogLevel.Info)
                Return
            End If

            Dim payload As String = BuildPayload(
                ConfigManager.CustomerName,
                ConfigManager.SiteName,
                Environment.MachineName,
                GetMachineKey(),
                EventTypeName(eventType),
                version,
                targetVersion,
                ResultName(eventType),
                message,
                Environment.OSVersion.VersionString)

            ' Fire-and-forget: do not await — dashboard reporting must never
            ' block or fail the update workflow.
            Dim ignored As Task = Task.Run(Function() SendAsync(apiUrl, apiKey, payload, EventTypeName(eventType)))
        Catch ex As Exception
            Logger.Log($"Failed to report dashboard event: {ex.Message}", Logger.LogLevel.Warning)
        End Try
    End Sub

    ''' <summary>
    ''' POST the JSON payload to the dashboard API. Runs on a background task;
    ''' all exceptions are caught and logged.
    ''' </summary>
    Private Shared Async Function SendAsync(apiUrl As String, apiKey As String, payload As String, eventName As String) As Task
        Try
            Using request As New HttpRequestMessage(HttpMethod.Post, apiUrl)
                request.Headers.Add("x-api-key", apiKey)
                request.Content = New StringContent(payload, Encoding.UTF8, "application/json")

                Using response As HttpResponseMessage = Await Client.SendAsync(request)
                    If response.IsSuccessStatusCode Then
                        Logger.Log($"Dashboard event sent: {eventName}", Logger.LogLevel.Info)
                    Else
                        Logger.Log($"Dashboard API returned {CInt(response.StatusCode)} for event {eventName}", Logger.LogLevel.Warning)
                    End If
                End Using
            End Using
        Catch ex As Exception
            Logger.Log($"Failed to send dashboard event {eventName}: {ex.Message}", Logger.LogLevel.Warning)
        End Try
    End Function

    ''' <summary>
    ''' Build the JSON payload for a status event.
    ''' Public and Shared so the payload format can be unit tested.
    ''' </summary>
    Public Shared Function BuildPayload(customer As String, site As String, hostname As String,
                                        machineKey As String, eventType As String, version As String,
                                        targetVersion As String, result As String, message As String,
                                        osVersion As String) As String
        Dim sb As New StringBuilder()
        sb.Append("{")
        sb.Append($"""customer"":""{EscapeJson(customer)}"",")
        sb.Append($"""site"":""{EscapeJson(site)}"",")
        sb.Append($"""hostname"":""{EscapeJson(hostname)}"",")
        sb.Append($"""machineKey"":""{EscapeJson(machineKey)}"",")
        sb.Append($"""eventType"":""{EscapeJson(eventType)}"",")
        sb.Append($"""version"":""{EscapeJson(version)}"",")
        sb.Append($"""targetVersion"":""{EscapeJson(targetVersion)}"",")
        sb.Append($"""result"":""{EscapeJson(result)}"",")
        sb.Append($"""message"":""{EscapeJson(message)}"",")
        sb.Append($"""osVersion"":""{EscapeJson(osVersion)}""")
        sb.Append("}")
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Escape a string value for safe inclusion in a JSON document.
    ''' </summary>
    Public Shared Function EscapeJson(value As String) As String
        If String.IsNullOrEmpty(value) Then Return String.Empty

        Dim sb As New StringBuilder(value.Length)
        For Each c As Char In value
            Select Case c
                Case """"c
                    sb.Append("\""")
                Case "\"c
                    sb.Append("\\")
                Case ControlChars.Cr
                    sb.Append("\r")
                Case ControlChars.Lf
                    sb.Append("\n")
                Case ControlChars.Tab
                    sb.Append("\t")
                Case Else
                    If AscW(c) < 32 Then
                        sb.Append("\u").Append(AscW(c).ToString("x4"))
                    Else
                        sb.Append(c)
                    End If
            End Select
        Next
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Map a <see cref="DashboardEventType"/> to its API wire name.
    ''' </summary>
    Public Shared Function EventTypeName(eventType As DashboardEventType) As String
        Select Case eventType
            Case DashboardEventType.Heartbeat
                Return "heartbeat"
            Case DashboardEventType.UpdateStart
                Return "update_start"
            Case DashboardEventType.UpdateSuccess
                Return "update_success"
            Case DashboardEventType.UpdateFailure
                Return "update_failure"
            Case Else
                Return "heartbeat"
        End Select
    End Function

    ''' <summary>
    ''' Derive the result field from the event type.
    ''' Only success/failure events carry a result value.
    ''' </summary>
    Public Shared Function ResultName(eventType As DashboardEventType) As String
        Select Case eventType
            Case DashboardEventType.UpdateSuccess
                Return "success"
            Case DashboardEventType.UpdateFailure
                Return "failure"
            Case Else
                Return String.Empty
        End Select
    End Function

    ''' <summary>
    ''' Read the Windows machine GUID from the registry to uniquely identify
    ''' this machine. Returns an empty string if unavailable.
    ''' </summary>
    Public Shared Function GetMachineKey() As String
        Try
            Using baseKey As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64)
                Using cryptoKey As RegistryKey = baseKey.OpenSubKey("SOFTWARE\Microsoft\Cryptography")
                    If cryptoKey IsNot Nothing Then
                        Return CStr(cryptoKey.GetValue("MachineGuid", String.Empty))
                    End If
                End Using
            End Using
        Catch ex As Exception
            Logger.Log($"Could not read machine GUID: {ex.Message}", Logger.LogLevel.Warning)
        End Try
        Return String.Empty
    End Function

End Class
