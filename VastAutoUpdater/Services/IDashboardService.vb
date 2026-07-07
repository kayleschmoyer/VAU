Option Strict On

''' <summary>
''' Status event types reported to the VAU-Dashboard API.
''' </summary>
Public Enum DashboardEventType
    Heartbeat
    UpdateStart
    UpdateSuccess
    UpdateFailure
End Enum

''' <summary>
''' Interface for reporting status events to the VAU-Dashboard API.
''' Enables dependency injection and unit testing with mocks.
''' </summary>
Public Interface IDashboardService
    ''' <summary>
    ''' Report a status event to the dashboard. Implementations must be
    ''' fire-and-forget: never block the caller and never throw.
    ''' </summary>
    ''' <param name="eventType">The kind of event being reported.</param>
    ''' <param name="version">Currently installed VAST version, if known.</param>
    ''' <param name="targetVersion">Version being installed, if an update is in progress.</param>
    ''' <param name="message">Failure reason or additional detail, if applicable.</param>
    Sub ReportEvent(eventType As DashboardEventType,
                    Optional version As String = "",
                    Optional targetVersion As String = "",
                    Optional message As String = "")
End Interface
