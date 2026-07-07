Option Strict On

Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass>
Public Class ConfigManagerTests

    <TestMethod>
    Public Sub GetSetting_MissingKey_ReturnsDefault()
        Dim result As String = ConfigManager.GetSetting("NonExistentKey_ABC123", "mydefault")
        Assert.AreEqual("mydefault", result)
    End Sub

    <TestMethod>
    Public Sub GetSetting_ExistingKey_ReturnsValue()
        System.Configuration.ConfigurationManager.AppSettings.Set("TestKey_Unit", "TestValue123")
        Try
            Dim result As String = ConfigManager.GetSetting("TestKey_Unit", "fallback")
            Assert.AreEqual("TestValue123", result)
        Finally
            System.Configuration.ConfigurationManager.AppSettings.Remove("TestKey_Unit")
        End Try
    End Sub

    <TestMethod>
    Public Sub GetSetting_EmptyValue_ReturnsDefault()
        System.Configuration.ConfigurationManager.AppSettings.Set("TestKey_Empty", "")
        Try
            Dim result As String = ConfigManager.GetSetting("TestKey_Empty", "fallback")
            Assert.AreEqual("fallback", result)
        Finally
            System.Configuration.ConfigurationManager.AppSettings.Remove("TestKey_Empty")
        End Try
    End Sub

    <TestMethod>
    Public Sub GetSetting_NullKey_ReturnsDefault()
        Dim result As String = ConfigManager.GetSetting(Nothing, "fallback")
        Assert.AreEqual("fallback", result)
    End Sub

    <TestMethod>
    Public Sub ValidateConfiguration_ReturnsFalse_WhenSftpNotConfigured()
        ' With no SFTP settings configured, validation should return False
        ' (default test app.config has empty values)
        Dim result As Boolean = ConfigManager.ValidateConfiguration()
        Assert.IsFalse(result, "Expected False when SFTP settings are not configured")
    End Sub

    <TestMethod>
    Public Sub SftpHost_ReturnsStringValue()
        ' SftpHost should return a string (empty or configured) without throwing
        Dim result As String = ConfigManager.SftpHost
        Assert.IsNotNull(result)
    End Sub

    <TestMethod>
    Public Sub SmtpPort_ReturnsPositiveInteger()
        ' SmtpPort should return the configured value or the default (587)
        Dim result As Integer = ConfigManager.SmtpPort
        Assert.IsTrue(result > 0, $"Expected positive port number, got {result}")
    End Sub

    <TestMethod>
    Public Sub SaveSetting_DoesNotThrow()
        ' SaveSetting should not throw even when config file is read-only in test
        ConfigManager.SaveSetting("TestSave_Unit", "value")
    End Sub

End Class
