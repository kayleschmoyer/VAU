Option Strict On

Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass>
Public Class UpdateExceptionTests

    <TestMethod>
    Public Sub Constructor_SetsErrorCodeAndMessage()
        Dim ex As New UpdateException(UpdateErrorCode.VastNotFound, "VAST.exe not found")
        Assert.AreEqual(UpdateErrorCode.VastNotFound, ex.ErrorCode)
        Assert.AreEqual("VAST.exe not found", ex.Message)
    End Sub

    <TestMethod>
    Public Sub Constructor_WithInnerException_PreservesChain()
        Dim inner As New System.IO.IOException("disk failure")
        Dim ex As New UpdateException(UpdateErrorCode.DownloadFailed, "Download failed", inner)
        Assert.AreEqual(UpdateErrorCode.DownloadFailed, ex.ErrorCode)
        Assert.AreEqual("Download failed", ex.Message)
        Assert.AreSame(inner, ex.InnerException)
    End Sub

    <TestMethod>
    Public Sub ErrorCode_AllValuesAreDefined()
        ' Verify all enum values are distinct and cover expected failure modes
        Dim values = System.Enum.GetValues(GetType(UpdateErrorCode))
        Assert.IsTrue(values.Length >= 10, $"Expected at least 10 error codes, got {values.Length}")

        ' Verify no duplicate integer values
        Dim intValues = New System.Collections.Generic.HashSet(Of Integer)()
        For Each v As UpdateErrorCode In values
            Assert.IsTrue(intValues.Add(CInt(v)), $"Duplicate error code value: {CInt(v)}")
        Next
    End Sub

    <TestMethod>
    Public Sub UpdateException_IsException()
        Dim ex As New UpdateException(UpdateErrorCode.Unknown, "test")
        Assert.IsInstanceOfType(ex, GetType(System.Exception))
    End Sub

    <TestMethod>
    Public Sub ErrorCode_Unknown_IsZero()
        Assert.AreEqual(0, CInt(UpdateErrorCode.Unknown))
    End Sub

    <TestMethod>
    Public Sub Constructor_EachErrorCode_CanBeAssigned()
        ' Verify each error code can be used in a real exception
        For Each code As UpdateErrorCode In System.Enum.GetValues(GetType(UpdateErrorCode))
            Dim ex As New UpdateException(code, $"Test for {code}")
            Assert.AreEqual(code, ex.ErrorCode)
        Next
    End Sub

End Class
