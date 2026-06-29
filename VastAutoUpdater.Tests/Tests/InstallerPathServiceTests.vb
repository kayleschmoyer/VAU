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
    Public Sub GetInstallPath_PathTraversalDotDot_ThrowsArgumentException()
        InstallerPathService.GetInstallPath("../../../etc")
    End Sub

    <TestMethod>
    <ExpectedException(GetType(ArgumentException))>
    Public Sub GetInstallPath_PathTraversalSemicolon_ThrowsArgumentException()
        InstallerPathService.GetInstallPath("1.0;rm -rf")
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

End Class
