Option Strict On

Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass>
Public Class LoggerTests

    <TestMethod>
    Public Sub Log_InfoLevel_DoesNotThrow()
        Logger.Log("Unit test - Info level", Logger.LogLevel.Info)
    End Sub

    <TestMethod>
    Public Sub Log_WarningLevel_DoesNotThrow()
        Logger.Log("Unit test - Warning level", Logger.LogLevel.Warning)
    End Sub

    <TestMethod>
    Public Sub Log_ErrorLevel_DoesNotThrow()
        Logger.Log("Unit test - Error level", Logger.LogLevel.Error)
    End Sub

    <TestMethod>
    Public Sub Log_EmptyMessage_DoesNotThrow()
        Logger.Log("", Logger.LogLevel.Info)
    End Sub

    <TestMethod>
    Public Sub Log_LongMessage_DoesNotThrow()
        Dim longMsg As New String("X"c, 10000)
        Logger.Log(longMsg, Logger.LogLevel.Info)
    End Sub

    <TestMethod>
    Public Sub Log_SpecialCharacters_DoesNotThrow()
        Logger.Log("Test with special chars: <>&""'{}[]|\/", Logger.LogLevel.Info)
    End Sub

    <TestMethod>
    Public Sub TrimLogFile_FileDoesNotExist_DoesNotThrow()
        Logger.TrimLogFile()
    End Sub

    <TestMethod>
    Public Sub TrimLogFile_CustomMaxSize_DoesNotThrow()
        ' TrimLogFile with a very small max size should not throw
        Logger.TrimLogFile(1024)
    End Sub

    <TestMethod>
    Public Sub LogLevel_HasThreeValues()
        Dim values = System.Enum.GetValues(GetType(Logger.LogLevel))
        Assert.AreEqual(3, values.Length, "Expected exactly 3 log levels: Info, Warning, Error")
    End Sub

End Class
