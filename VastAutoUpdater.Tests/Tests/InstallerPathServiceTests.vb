Option Strict On

Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass>
Public Class InstallerPathServiceTests

    <TestMethod>
    <ExpectedException(GetType(ArgumentException))>
    Public Sub GetInstallPath_EmptyString_ThrowsArgumentException()
        InstallerPathService.GetInstallPath("")
    End Sub

    <TestMethod>
    <ExpectedException(GetType(ArgumentException))>
    Public Sub GetInstallPath_Nothing_ThrowsArgumentException()
        InstallerPathService.GetInstallPath(Nothing)
    End Sub

    <TestMethod>
    <ExpectedException(GetType(ArgumentException))>
    Public Sub GetInstallPath_PathTraversalDotDot_ThrowsArgumentException()
        InstallerPathService.GetInstallPath("../../../etc")
    End Sub

    <TestMethod>
    <ExpectedException(GetType(ArgumentException))>
    Public Sub GetInstallPath_PathTraversalSemicolon_ThrowsArgumentException()
        InstallerPathService.GetInstallPath("1.0;rm -rf")
    End Sub

    <TestMethod>
    <ExpectedException(GetType(ArgumentException))>
    Public Sub GetInstallPath_PathTraversalBackslash_ThrowsArgumentException()
        InstallerPathService.GetInstallPath("..\windows\system32")
    End Sub

    <TestMethod>
    <ExpectedException(GetType(ArgumentException))>
    Public Sub GetInstallPath_LettersInVersion_ThrowsArgumentException()
        InstallerPathService.GetInstallPath("1.2.3a")
    End Sub

    <TestMethod>
    <ExpectedException(GetType(ArgumentException))>
    Public Sub GetInstallPath_SpacesInVersion_ThrowsArgumentException()
        InstallerPathService.GetInstallPath("1.2 .3")
    End Sub

    <TestMethod>
    <ExpectedException(GetType(ArgumentException))>
    Public Sub GetInstallPath_SingleNumber_ThrowsArgumentException()
        InstallerPathService.GetInstallPath("123")
    End Sub

    <TestMethod>
    Public Sub GetInstallPath_ValidThreePartVersion_ReturnsPathEndingWithExe()
        Dim result As String = InstallerPathService.GetInstallPath("1.2.3")
        Assert.IsTrue(result.EndsWith(".exe"), $"Expected path ending with .exe but got: {result}")
    End Sub

    <TestMethod>
    Public Sub GetInstallPath_ValidFourPartVersion_ReturnsPathEndingWithExe()
        Dim result As String = InstallerPathService.GetInstallPath("1.2.3.4")
        Assert.IsTrue(result.EndsWith(".exe"), $"Expected path ending with .exe but got: {result}")
    End Sub

    <TestMethod>
    Public Sub GetInstallPath_ValidTwoPartVersion_ReturnsPathEndingWithExe()
        Dim result As String = InstallerPathService.GetInstallPath("1.0")
        Assert.IsTrue(result.EndsWith(".exe"), $"Expected path ending with .exe but got: {result}")
    End Sub

    <TestMethod>
    Public Sub GetInstallPath_ValidVersion_ContainsVersionInPath()
        Dim result As String = InstallerPathService.GetInstallPath("5.6.7")
        Assert.IsTrue(result.Contains("5.6.7"), $"Expected version in path but got: {result}")
    End Sub

    <TestMethod>
    Public Sub GetInstallPath_ValidVersion_IsAbsolutePath()
        Dim result As String = InstallerPathService.GetInstallPath("1.0.0")
        Assert.IsTrue(System.IO.Path.IsPathRooted(result), $"Expected absolute path but got: {result}")
    End Sub

    <TestMethod>
    Public Sub CleanupOldInstallers_DoesNotThrow_WhenFolderMissing()
        ' Should handle missing folder gracefully
        InstallerPathService.CleanupOldInstallers("99.99.99")
    End Sub

End Class
