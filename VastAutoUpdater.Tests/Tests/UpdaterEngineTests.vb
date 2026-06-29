Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports System.Threading

Option Strict On

''' <summary>
''' Tests for UpdaterEngine using mock SFTP and email services.
''' </summary>
<TestClass>
Public Class UpdaterEngineTests

    ''' <summary>
    ''' Mock SFTP service for testing the engine without a real server.
    ''' </summary>
    Private Class MockSftpService
        Implements ISftpService

        Public Property ConnectCalled As Boolean = False
        Public Property DisconnectCalled As Boolean = False
        Public Property LatestVersion As String = "0.0.0"
        Public Property RemoteFileSize As Long = 0
        Public Property DownloadResult As Boolean = True
        Public Property ShouldThrowOnConnect As Boolean = False

        Public Sub Connect(username As String, password As String) Implements ISftpService.Connect
            If ShouldThrowOnConnect Then
                Throw New InvalidOperationException("Connection failed")
            End If
            ConnectCalled = True
        End Sub

        Public Function GetLatestVersion(prefix As String) As String Implements ISftpService.GetLatestVersion
            Return LatestVersion
        End Function

        Public Function GetRemoteFileSize(remoteFileName As String) As Long Implements ISftpService.GetRemoteFileSize
            Return RemoteFileSize
        End Function

        Public Function DownloadFile(remoteFileName As String, localPath As String, progress As Action(Of ULong)) As Boolean Implements ISftpService.DownloadFile
            ' Simulate download progress
            progress(CULng(1024))
            Return DownloadResult
        End Function

        Public Sub Disconnect() Implements ISftpService.Disconnect
            DisconnectCalled = True
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Disconnect()
        End Sub
    End Class

    ''' <summary>
    ''' Mock email service for testing without SMTP.
    ''' </summary>
    Private Class MockEmailService
        Implements IEmailService

        Public Property SendCalled As Boolean = False
        Public Property LastSuccess As Boolean = False
        Public Property LastDetails As String = String.Empty

        Public Sub SendSummary(success As Boolean, details As String, Optional exception As Exception = Nothing) Implements IEmailService.SendSummary
            SendCalled = True
            LastSuccess = success
            LastDetails = details
        End Sub
    End Class

    Private _mockSftp As MockSftpService
    Private _mockEmail As MockEmailService
    Private _engine As UpdaterEngine
    Private _lastProgress As Integer
    Private _lastStatus As String

    <TestInitialize>
    Public Sub Setup()
        _mockSftp = New MockSftpService()
        _mockEmail = New MockEmailService()
        _engine = New UpdaterEngine(_mockEmail, Function() _mockSftp)
        _lastProgress = 0
        _lastStatus = String.Empty
    End Sub

    Private Sub TrackProgress(p As Integer, status As String)
        _lastProgress = p
        _lastStatus = status
    End Sub

    <TestMethod>
    Public Async Function NoUpdateAvailable_ReportsSuccess() As Task
        _mockSftp.LatestVersion = "0.0.0"

        Await _engine.PerformUpdateCheck("user", "pass", AddressOf TrackProgress)

        Assert.IsTrue(_mockSftp.ConnectCalled, "Should have connected to SFTP")
        Assert.IsTrue(_mockEmail.SendCalled, "Should have sent summary email")
        Assert.IsTrue(_mockEmail.LastSuccess, "Should report success when no update available")
        Assert.AreEqual(100, _lastProgress, "Should reach 100% progress")
    End Function

    <TestMethod>
    <ExpectedException(GetType(UpdateException))>
    Public Async Function ConnectionFailure_ThrowsUpdateException() As Task
        _mockSftp.ShouldThrowOnConnect = True

        Await _engine.PerformUpdateCheck("user", "pass", AddressOf TrackProgress)
    End Function

    <TestMethod>
    Public Async Function AlreadyUpToDate_ReportsSuccess() As Task
        ' GetLatestVersion returns "0.0.0" by default (no update)
        _mockSftp.LatestVersion = "0.0.0"

        Await _engine.PerformUpdateCheck("user", "pass", AddressOf TrackProgress)

        Assert.IsTrue(_mockEmail.LastSuccess)
        Assert.IsTrue(_lastStatus.Contains("No update available") OrElse _lastStatus.Contains("up-to-date"))
    End Function

    <TestMethod>
    Public Sub CancellationToken_PropagatesCancellation()
        Dim cts As New CancellationTokenSource()
        cts.Cancel()

        Dim task = _engine.PerformUpdateCheck("user", "pass", AddressOf TrackProgress, cts.Token)

        Assert.IsTrue(task.IsCanceled OrElse task.IsFaulted, "Task should be cancelled or faulted")
        If task.IsFaulted Then
            Assert.IsInstanceOfType(task.Exception.InnerException, GetType(OperationCanceledException))
        End If
        cts.Dispose()
    End Sub

    <TestMethod>
    Public Sub EmailService_CalledEvenOnFailure()
        _mockSftp.ShouldThrowOnConnect = True

        Try
            _engine.PerformUpdateCheck("user", "pass", AddressOf TrackProgress).Wait()
        Catch ex As AggregateException
            ' Expected
        End Try

        Assert.IsTrue(_mockEmail.SendCalled, "Email should be sent even when update fails")
        Assert.IsFalse(_mockEmail.LastSuccess, "Email should report failure")
    End Sub

    <TestMethod>
    Public Sub Constructor_DefaultServices_DoesNotThrow()
        ' Verify the parameterless constructor works
        Dim engine As New UpdaterEngine()
        Assert.IsNotNull(engine)
    End Sub

End Class
