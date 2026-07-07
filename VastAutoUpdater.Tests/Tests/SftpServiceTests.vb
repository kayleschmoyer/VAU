Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports VastAutoUpdater

''' <summary>
''' Tests for SftpService.SelectLatestVersion — the pure version-selection
''' logic that decides which installer to download from /VASTAutoInstall/.
''' Locks in support for both 3-part (9.0.1653.exe) and 4-part
''' (9.0.1653.1.exe) file names observed on the SFTP server.
''' </summary>
<TestClass>
Public Class SftpServiceTests

    <TestMethod>
    Public Sub SelectLatest_ThreePartName_IsSelected()
        Dim result = SftpService.SelectLatestVersion({"9.0.1600.exe", "9.0.1653.exe"}, "9.0")
        Assert.AreEqual("9.0.1653", result)
    End Sub

    <TestMethod>
    Public Sub SelectLatest_FourPartBeatsSameThreePartBase()
        ' 9.0.1653.1 is newer than 9.0.1653
        Dim result = SftpService.SelectLatestVersion({"9.0.1653.exe", "9.0.1653.1.exe"}, "9.0")
        Assert.AreEqual("9.0.1653.1", result)
    End Sub

    <TestMethod>
    Public Sub SelectLatest_ResultRoundTripsToOriginalFileName()
        ' The engine downloads "{result}.exe", so the selected version string
        ' must reproduce the exact file name it was parsed from
        Dim files = {"9.0.1600.exe", "9.0.1653.exe", "9.0.1653.1.exe"}
        Dim result = SftpService.SelectLatestVersion(files, "9.0")
        CollectionAssert.Contains(files, result & ".exe")
    End Sub

    <TestMethod>
    Public Sub SelectLatest_FiltersToMajorMinorPrefix()
        Dim result = SftpService.SelectLatestVersion({"9.1.9999.exe", "9.0.50.exe"}, "9.0")
        Assert.AreEqual("9.0.50", result)
    End Sub

    <TestMethod>
    Public Sub SelectLatest_IgnoresNonVersionFileNames()
        Dim result = SftpService.SelectLatestVersion(
            {"VastPOS 901072 Patch71.exe", "readme.txt", "9.0.1653.exe.sha256"}, "9.0")
        Assert.AreEqual("0.0.0", result)
    End Sub

    <TestMethod>
    Public Sub SelectLatest_NoMatches_ReturnsSentinel()
        Dim result = SftpService.SelectLatestVersion(New String() {}, "9.0")
        Assert.AreEqual("0.0.0", result)
    End Sub

    <TestMethod>
    Public Sub SelectLatest_CaseInsensitiveExtension()
        Dim result = SftpService.SelectLatestVersion({"9.0.1653.EXE"}, "9.0")
        Assert.AreEqual("9.0.1653", result)
    End Sub

End Class
