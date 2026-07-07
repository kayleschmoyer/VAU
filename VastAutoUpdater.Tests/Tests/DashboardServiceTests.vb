Option Strict On

Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' Tests for DashboardService payload building and event mapping.
''' Network sends are not exercised here — the service is fire-and-forget
''' and skips sending entirely when the dashboard is not configured.
''' </summary>
<TestClass>
Public Class DashboardServiceTests

    <TestMethod>
    Public Sub BuildPayload_ContainsAllFields()
        Dim json As String = DashboardService.BuildPayload(
            "Acme", "Store 1", "HOST01", "abc-123", "update_start",
            "6.3.100", "6.3.200", "", "", "Windows 11")

        Assert.IsTrue(json.Contains("""customer"":""Acme"""))
        Assert.IsTrue(json.Contains("""site"":""Store 1"""))
        Assert.IsTrue(json.Contains("""hostname"":""HOST01"""))
        Assert.IsTrue(json.Contains("""machineKey"":""abc-123"""))
        Assert.IsTrue(json.Contains("""eventType"":""update_start"""))
        Assert.IsTrue(json.Contains("""version"":""6.3.100"""))
        Assert.IsTrue(json.Contains("""targetVersion"":""6.3.200"""))
        Assert.IsTrue(json.Contains("""result"":"""""))
        Assert.IsTrue(json.Contains("""message"":"""""))
        Assert.IsTrue(json.Contains("""osVersion"":""Windows 11"""))
        Assert.IsTrue(json.StartsWith("{") AndAlso json.EndsWith("}"))
    End Sub

    <TestMethod>
    Public Sub BuildPayload_HandlesNullValues()
        Dim json As String = DashboardService.BuildPayload(
            Nothing, Nothing, Nothing, Nothing, "heartbeat",
            Nothing, Nothing, Nothing, Nothing, Nothing)

        Assert.IsTrue(json.Contains("""customer"":"""""))
        Assert.IsTrue(json.Contains("""eventType"":""heartbeat"""))
    End Sub

    <TestMethod>
    Public Sub EscapeJson_EscapesSpecialCharacters()
        Assert.AreEqual("a\""b", DashboardService.EscapeJson("a""b"))
        Assert.AreEqual("a\\b", DashboardService.EscapeJson("a\b"))
        Assert.AreEqual("a\r\nb", DashboardService.EscapeJson("a" & vbCrLf & "b"))
        Assert.AreEqual("a\tb", DashboardService.EscapeJson("a" & vbTab & "b"))
        Assert.AreEqual(String.Empty, DashboardService.EscapeJson(Nothing))
    End Sub

    <TestMethod>
    Public Sub EscapeJson_EscapesControlCharacters()
        Dim escaped As String = DashboardService.EscapeJson(ChrW(1).ToString())
        Assert.AreEqual("\u0001", escaped)
    End Sub

    <TestMethod>
    Public Sub EventTypeName_MapsAllEventTypes()
        Assert.AreEqual("heartbeat", DashboardService.EventTypeName(DashboardEventType.Heartbeat))
        Assert.AreEqual("update_start", DashboardService.EventTypeName(DashboardEventType.UpdateStart))
        Assert.AreEqual("update_success", DashboardService.EventTypeName(DashboardEventType.UpdateSuccess))
        Assert.AreEqual("update_failure", DashboardService.EventTypeName(DashboardEventType.UpdateFailure))
    End Sub

    <TestMethod>
    Public Sub ResultName_OnlySetForSuccessAndFailure()
        Assert.AreEqual("success", DashboardService.ResultName(DashboardEventType.UpdateSuccess))
        Assert.AreEqual("failure", DashboardService.ResultName(DashboardEventType.UpdateFailure))
        Assert.AreEqual(String.Empty, DashboardService.ResultName(DashboardEventType.Heartbeat))
        Assert.AreEqual(String.Empty, DashboardService.ResultName(DashboardEventType.UpdateStart))
    End Sub

    <TestMethod>
    Public Sub ReportEvent_UnconfiguredDashboard_DoesNotThrow()
        ' With no DashboardApiUrl/DashboardApiKey configured in the test
        ' environment, ReportEvent must silently skip without throwing.
        Dim service As New DashboardService()
        service.ReportEvent(DashboardEventType.Heartbeat, "6.3.100")
    End Sub

End Class
