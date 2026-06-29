''' <summary>
''' Interface for email notification used by the update engine.
''' Enables dependency injection and unit testing with mocks.
''' </summary>
Public Interface IEmailService
    ''' <summary>
    ''' Send an email summarizing the update process result.
    ''' </summary>
    Sub SendSummary(success As Boolean, details As String, Optional exception As Exception = Nothing)
End Interface
