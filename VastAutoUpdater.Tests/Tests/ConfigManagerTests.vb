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
        ' ConfigManager reads from ConfigurationManager.AppSettings.
        ' We can test by adding a value at runtime.
        System.Configuration.ConfigurationManager.AppSettings.Set("TestKey_Unit", "TestValue123")
        Try
            Dim result As String = ConfigManager.GetSetting("TestKey_Unit", "fallback")
            Assert.AreEqual("TestValue123", result)
        Finally
            System.Configuration.ConfigurationManager.AppSettings.Remove("TestKey_Unit")
        End Try
    End Sub

End Class
