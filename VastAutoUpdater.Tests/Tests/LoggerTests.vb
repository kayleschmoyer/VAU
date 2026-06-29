Option Strict On

Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass>
Public Class LoggerTests

    <TestMethod>
    Public Sub Log_AllLogLevels_DoNotThrow()
        ' Calling Log should not throw for any log level.
        ' It may fail to write to Event Log (no admin rights in test) but should
        ' fall back gracefully to file or Trace output.
        Logger.Log("Unit test - Info level", Logger.LogLevel.Info)
        Logger.Log("Unit test - Warning level", Logger.LogLevel.Warning)
        Logger.Log("Unit test - Error level", Logger.LogLevel.Error)
    End Sub

    <TestMethod>
    Public Sub TrimLogFile_FileDoesNotExist_DoesNotThrow()
        ' TrimLogFile should return silently when the log file does not exist.
        ' It checks File.Exists internally and returns early.
        Logger.TrimLogFile()
    End Sub

End Class
