Option Strict On

Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports System.Configuration

<TestClass>
Public Class ConfigManagerTests

    ''' <summary>
    ''' ConfigurationManager.AppSettings is a read-only collection, and the
    ''' backing config file can live in an unwritable location under the test
    ''' host, so tests mutate the cached in-memory collection directly after
    ''' clearing its read-only flag.
    ''' </summary>
    Private Shared Function WritableAppSettings() As Specialized.NameValueCollection
        Dim settings As Specialized.NameValueCollection = ConfigurationManager.AppSettings
        Dim readOnlyField As Reflection.FieldInfo =
            GetType(Specialized.NameObjectCollectionBase).GetField("_readOnly",
                Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
        readOnlyField.SetValue(settings, False)
        Return settings
    End Function

    Private Shared Sub SetAppSetting(key As String, value As String)
        ' Set is not overridden by the runtime collection, so it mutates in-memory only
        WritableAppSettings().Set(key, value)
    End Sub

    Private Shared Sub RemoveAppSetting(key As String)
        ' Remove is overridden to write through to the read-only config record,
        ' so call the in-memory BaseRemove instead
        Dim settings As Specialized.NameValueCollection = WritableAppSettings()
        Dim baseRemove As Reflection.MethodInfo =
            GetType(Specialized.NameObjectCollectionBase).GetMethod("BaseRemove",
                Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
        baseRemove.Invoke(settings, New Object() {key})
    End Sub

    <TestMethod>
    Public Sub GetSetting_MissingKey_ReturnsDefault()
        Dim result As String = ConfigManager.GetSetting("NonExistentKey_ABC123", "mydefault")
        Assert.AreEqual("mydefault", result)
    End Sub

    <TestMethod>
    Public Sub GetSetting_ExistingKey_ReturnsValue()
        SetAppSetting("TestKey_Unit", "TestValue123")
        Try
            Dim result As String = ConfigManager.GetSetting("TestKey_Unit", "fallback")
            Assert.AreEqual("TestValue123", result)
        Finally
            RemoveAppSetting("TestKey_Unit")
        End Try
    End Sub

    <TestMethod>
    Public Sub GetSetting_EmptyValue_ReturnsDefault()
        SetAppSetting("TestKey_Empty", "")
        Try
            Dim result As String = ConfigManager.GetSetting("TestKey_Empty", "fallback")
            Assert.AreEqual("fallback", result)
        Finally
            RemoveAppSetting("TestKey_Empty")
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
