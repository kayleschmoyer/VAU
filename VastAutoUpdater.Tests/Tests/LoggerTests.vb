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
 