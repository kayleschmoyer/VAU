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
        Inst