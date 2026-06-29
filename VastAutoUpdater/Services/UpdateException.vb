Option Strict On

''' <summary>
''' Categorized error codes for the update workflow.
''' Enables programmatic handling of specific failure modes.
''' </summary>
Public Enum UpdateErrorCode
    ''' <summary>No specific error code assigned.</summary>
    Unknown = 0
    ''' <summary>VAST.exe could not be found on any fixed drive.</summary>
    VastNotFound = 1
    ''' <summary>The current version string could not be parsed.</summary>
    VersionParseError = 2
    ''' <summary>SFTP connection failed after all retry attempts.</summary>
    ConnectionFailed = 3
    ''' <summary>No matching version was found on the server.</summary>
    NoUpdateAvailable = 4
    ''' <summary>The installer file download failed or produced an empty file.</summary>
    DownloadFailed = 5
    ''' <summary>SHA-256 hash of the downloaded installer does not match the sidecar.</summary>
    HashMismatch = 6
    ''' <summary>The installer process failed to start or exited with a non-zero code.</summary>
    InstallerFailed = 7
    ''' <summary>The update was cancelled by the user.</summary>
    Cancelled = 8
    ''' <summary>SFTP credentials are missing or empty.</summary>
    CredentialsMissing = 9
    ''' <summary>SFTP host configuration is missing.</summary>
    ConfigurationError = 10
End Enum

''' <summary>
''' Custom exception for update workflow failures.
''' Carries a structured <see cref="UpdateErrorCode"/> for programmatic handling
''' in addition to the human-readable message.
''' </summary>
Public Class UpdateException
    Inherits Exception

    ''' <summary>
    ''' Gets the structured error code identifying the failure category.
    ''' </summary>
    Public ReadOnly Property ErrorCode As UpdateErrorCode

    ''' <summary>
    ''' Initializes a new <see cref="UpdateException"/> with the specified error code and message.
    ''' </summary>
    ''' <param name="errorCode">The categorized error code.</param>
    ''' <param name="message">A human-readable description of the error.</param>
    Public Sub New(errorCode As UpdateErrorCode, message As String)
        MyBase.New(message)
        Me.ErrorCode = errorCode
    End Sub

    ''' <summary>
    ''' Initializes a new <see cref="UpdateException"/> with the specified error code, message, and inner exception.
    ''' </summary>
    ''' <param name="errorCode">The categorized error code.</param>
    ''' <param name="message">A human-readable description of the error.</param>
    ''' <param name="innerException">The exception that caused this failure.</param>
    Public Sub New(errorCode As UpdateErrorCode, message As String, innerException As Exception)
        MyBase.New(message, innerException)
        Me.ErrorCode = errorCode
    End Sub
End Class
