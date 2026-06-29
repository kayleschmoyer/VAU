''' <summary>
''' Interface for SFTP operations used by the update engine.
''' Enables dependency injection and unit testing with mocks.
''' </summary>
Public Interface ISftpService
    Inherits IDisposable

    ''' <summary>
    ''' Connect to the SFTP server.
    ''' </summary>
    Sub Connect(username As String, password As String)

    ''' <summary>
    ''' Query the server for the latest version matching the given major.minor prefix.
    ''' Returns "0.0.0" if no matching version is found.
    ''' </summary>
    Function GetLatestVersion(prefix As String) As String

    ''' <summary>
    ''' Get the size of a remote file in bytes. Returns 0 if the file doesn't exist.
    ''' </summary>
    Function GetRemoteFileSize(remoteFileName As String) As Long

    ''' <summary>
    ''' Download a remote file to localPath. Returns True on success.
    ''' </summary>
    Function DownloadFile(remoteFileName As String, localPath As String, progress As Action(Of ULong)) As Boolean

    ''' <summary>
    ''' Disconnect from the server.
    ''' </summary>
    Sub Disconnect()
End Interface
