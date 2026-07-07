Option Strict On

Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass>
Public Class VersionServiceTests

    <TestMethod>
    Public Sub GetFileVersion_NonExistentFile_ReturnsZeroVersion()
        Dim result As String = VersionService.GetFileVersion("C:\nonexistent\fake.exe")
        Assert.AreEqual("0.0.0", result)
    End Sub

    <TestMethod>
    Public Sub GetFileVersion_RealExecutable_ReturnsVersionString()
        Dim assemblyPath As String = System.Reflection.Assembly.GetExecutingAssembly().Location
        Dim result As String = VersionService.GetFileVersion(assemblyPath)
        Assert.IsNotNull(result)
        Assert.AreNotEqual(String.Empty, result)
    End Sub

    <TestMethod>
    Public Sub GetFileVersion_EmptyString_ReturnsZeroVersion()
        Dim result As String = VersionService.GetFileVersion("")
        Assert.AreEqual("0.0.0", result)
    End Sub

    <TestMethod>
    Public Sub GetFileVersion_NullPath_ReturnsZeroVersion()
        Dim result As String = VersionService.GetFileVersion(Nothing)
        Assert.AreEqual("0.0.0", result)
    End Sub

    <TestMethod>
    Public Sub FindVastExecutable_ReturnsStringResult()
        ' FindVastExecutable should return either a valid path or empty string
        ' On a dev machine without VAST installed, it returns empty
        Dim result As String = VersionService.FindVastExecutable()
        Assert.IsNotNull(result)
        ' Result is either empty or ends with VAST.exe
        If Not String.IsNullOrEmpty(result) Then
            Assert.IsTrue(result.EndsWith("VAST.exe", StringComparison.OrdinalIgnoreCase))
        End If
    End Sub

    <TestMethod>
    Public Sub VersionComparison_NewerVersionIsGreater()
        ' Test the Version comparison logic used by UpdaterEngine
        Dim current As New Version("1.2.3.0")
        Dim newer As New Version("1.2.4.0")
        Assert.IsTrue(newer.CompareTo(current) > 0)
    End Sub

    <TestMethod>
    Public Sub VersionComparison_SameVersionIsEqual()
        Dim v1 As New Version("1.2.3.0")
        Dim v2 As New Version("1.2.3.0")
        Assert.AreEqual(0, v1.CompareTo(v2))
    End Sub

    <TestMethod>
    Public Sub VersionComparison_OlderVersionIsLess()
        Dim current As New Version("2.0.0.0")
        Dim older As New Version("1.9.9.9")
        Assert.IsTrue(older.CompareTo(current) < 0)
    End Sub

    <TestMethod>
    Public Sub VersionTryParse_ValidThreePartVersion_Succeeds()
        Dim v As Version = Nothing
        Assert.IsTrue(Version.TryParse("1.2.3", v))
        Assert.AreEqual(1, v.Major)
        Assert.AreEqual(2, v.Minor)
        Assert.AreEqual(3, v.Build)
    End Sub

    <TestMethod>
    Public Sub VersionTryParse_InvalidString_ReturnsFalse()
        Dim v As Version = Nothing
        Assert.IsFalse(Version.TryParse("not.a.version", v))
    End Sub

End Class
