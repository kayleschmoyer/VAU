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
        