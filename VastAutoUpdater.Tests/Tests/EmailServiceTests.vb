Option Strict On

Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' Tests for EmailService validation and early-return paths.
''' These tests run without an SMTP server — they verify that
''' missing/invalid configuration is handled gracefully.
''' </summary>
<TestClass>
Public Class EmailServiceTests

    Private _service As EmailService

    <TestInitialize>
    Public Sub Setup()
        _service = New EmailService()
    End Sub

    <TestMethod>
    Public Sub SendSummary_WithIncompleteConfig_DoesNotThrow()
        ' App.config in test project has empty SMTP settings — should return silently
        _service.SendSummary(True, "Test details")
    End Sub

    <TestMethod>
    Public Sub SendSummary_WithFailure_DoesNotThrow()
        ' Even with failure + exception, should not throw when config is missing
        Dim ex As New InvalidOperationException("Test error")
        _service.SendSummary(False, "Something failed", ex)
    End Sub

    <TestMethod>
    Public Sub SendSummary_WithNullDetails_DoesNotThrow()
        _service.SendSummary(True, Nothing)
    End Sub

    <TestMethod>
    Public Sub SendSummary_WithEmptyDetails_DoesNotThrow()
        _service.SendSummary(False, String.Empty)
    End Sub

    <TestMethod>
    Public Sub ImplementsIEmailService()
        Assert.IsInstanceOfType(_service, GetType(IEmailService))
    End Sub

End Class
