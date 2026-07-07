Option Strict On

Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports System.Threading
Imports System.Threading.Tasks

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

    ''' <summary>
    ''' Mock dashboard service for testing without HTTP.
    ''' </summary>
    Private Class MockDashboardService
        Implements IDashboardService

        Public Property Events As New List(Of DashboardEventType)()
        Public Property LastVersion As String = String.Empty
        Public Property LastTargetVersion As String = String.Empty
        Public Property LastMessage As String = String.Empty

        Public Sub ReportEvent(eventType As DashboardEventType,
                               Optional version As String = "",
                               Optional targetVersion As String = "",
                               Optional message As String = "") Implements IDashboardService.ReportEvent
            Events.Add(eventType)
            LastVersion = version
            LastTargetVersion = targetVersion
            LastMessage = message
        End Sub
    End Class

    Private _mockSftp As MockSftpService
    Private _mockEmail As MockEmailService
    Private _mockDashboard As MockDashboardService
    Private _engine As UpdaterEngine
    Private _lastProgress As Integer
    Private _lastStatus As String

    <TestInitialize>
    Public Sub Setup()
        _mockSftp = New MockSftpService()
        _mockEmail = New MockEmailService()
        _mockDashboard = New MockDashboardService()
        _engine = New UpdaterEngine(_mockEmail, Function() _mockSftp, _mockDashboard)
        _lastProgress = 0
        _lastStatus = String.Empty
    End Sub

    Private Sub TrackProgress(p As Integer, status As String)
        _lastProgress = p
        _lastStatus = status
    End Sub

    ''' <summary>
    ''' Skip tests that need a real VAST installation when running on a
    ''' machine (e.g. a hosted CI runner) where VAST.exe is not installed.
    ''' </summary>
    Private Shared Sub AssumeVastInstalled()
        If String.IsNullOrEmpty(VersionService.FindVastExecutable()) Then
            Assert.Inconclusive("VAST.exe is not installed on this machine — skipping environment-dependent test")
        End If
    End Sub

    <TestMethod>
    Public Async Function NoUpdateAvailable_ReportsSuccess() As Task
        AssumeVastInstalled()
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
        AssumeVastInstalled()
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

        ' Wait for the task to finish so the assertions don't race its completion
        Try
            task.Wait(TimeSpan.FromSeconds(30))
        Catch ex As AggregateException
            ' Expected — cancellation surfaces here when waiting
        End Try

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
    Public Sub DashboardService_ReportsFailureOnError()
        _mockSftp.ShouldThrowOnConnect = True

        Try
            _engine.PerformUpdateCheck("user", "pass", AddressOf TrackProgress).Wait()
        Catch ex As AggregateException
            ' Expected
        End Try

        Assert.IsTrue(_mockDashboard.Events.Contains(DashboardEventType.UpdateFailure),
                      "Dashboard should receive update_failure when the update fails")
        Assert.IsFalse(String.IsNullOrEmpty(_mockDashboard.LastMessage),
                       "Failure report should include a message")
    End Sub

    <TestMethod>
    Public Sub DashboardService_NoUpdateEventsWhenNoUpdateAvailable()
        _mockSftp.LatestVersion = "0.0.0"

        Try
            _engine.PerformUpdateCheck("user", "pass", AddressOf TrackProgress).Wait()
        Catch ex As AggregateException
            ' VAST.exe may not exist in the test environment
        End Try

        Assert.IsFalse(_mockDashboard.Events.Contains(DashboardEventType.UpdateStart),
                       "update_start should not be reported when no update is needed")
        Assert.IsFalse(_mockDashboard.Events.Contains(DashboardEventType.UpdateSuccess),
                       "update_success should not be reported when no update is needed")
    End Sub

    <TestMethod>
    Public Sub Constructor_DefaultServices_DoesNotThrow()
        ' Verify the parameterless constructor works
        Dim engine As New UpdaterEngine()
        Assert.IsNotNull(engine)
    End Sub

    <TestMethod>
    Public Sub VastNotFound_ThrowsUpdateExceptionWithCorrectCode()
        ' VAST.exe won't exist in the test environment, so the engine
        ' should throw UpdateException with VastNotFound error code
        Dim thrown As UpdateException = Nothing
        Try
            _engine.PerformUpdateCheck("user", "pass", AddressOf TrackProgress).Wait()
        Catch ex As AggregateException
            thrown = TryCast(ex.InnerException, UpdateException)
        End Try

        Assert.IsNotNull(thrown, "Should throw UpdateException when VAST.exe not found")
        Assert.AreEqual(UpdateErrorCode.VastNotFound, thrown.ErrorCode)
    End Sub

    <TestMethod>
    Public Sub VastNotFound_EmailStillSent()
        ' Even when VAST.exe is not found, summary email should be sent
        Try
            _engine.PerformUpdateCheck("user", "pass", AddressOf TrackProgress).Wait()
        Catch ex As AggregateException
            ' Expected
        End Try

        Assert.IsTrue(_mockEmail.SendCalled, "Email should be sent even on VastNotFound")
        Assert.IsFalse(_mockEmail.LastSuccess, "Email should report failure")
    End Sub

    <TestMethod>
    Public Sub ProgressCallback_IsInvoked()
        ' Even when failing early (VastNotFound), the initial progress should fire
        Dim progressCalled As Boolean = False
        Dim progressAction As Action(Of Integer, String) = Sub(p, s) progressCalled = True

        Try
            _engine.PerformUpdateCheck("user", "pass", progressAction).Wait()
        Catch ex As AggregateException
            ' Expected
        End Try

        Assert.IsTrue(progressCalled, "Progress callback should be invoked at least once")
    End Sub

    <TestMethod>
    Public Sub SftpDisposed_AfterExecution()
        ' After the engine runs (even on failure), the SFTP service should be disposed
        Try
            _engine.PerformUpdateCheck("user", "pass", AddressOf TrackProgress).Wait()
        Catch ex As AggregateException
            ' Expected — VastNotFound
        End Try

        Assert.IsTrue(_mockSftp.DisconnectCalled, "SFTP should be disconnected/disposed after execution")
    End Sub

    <TestMethod>
    Public Sub MultipleRuns_EachGetsFreshSftpInstance()
        ' Verify the factory creates a new instance each time
        Dim instanceCount As Integer = 0
        Dim engine As New UpdaterEngine(_mockEmail, Function()
                                                         instanceCount += 1
                                                         Return New MockSftpService()
                                                     End Function)

        Try
            engine.PerformUpdateCheck("user", "pass", AddressOf TrackProgress).Wait()
        Catch
        End Try
        Try
            engine.PerformUpdateCheck("user", "pass", AddressOf TrackProgress).Wait()
        Catch
        End Try

        Assert.AreEqual(2, instanceCount, "Factory should be called once per PerformUpdateCheck")
    End Sub

End Class
