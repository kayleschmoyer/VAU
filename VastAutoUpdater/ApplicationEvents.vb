Namespace My
    Partial Friend Class MyApplication

        ''' <summary>
        ''' Global handler for any unhandled exception in the application.
        ''' Logs the error to the file-based logger and exits gracefully
        ''' instead of crashing with a raw exit code.
        ''' </summary>
        Private Sub MyApplication_UnhandledException(
            sender As Object,
            e As Microsoft.VisualBasic.ApplicationServices.UnhandledExceptionEventArgs
        ) Handles Me.UnhandledException

            ' Prevent default crash dialog
            e.ExitApplication = True

            Try
                Logger.Log($"UNHANDLED EXCEPTION: {e.Exception.GetType().Name}: {e.Exception.Message}", Logger.LogLevel.Error)
                Logger.Log($"Stack Trace: {e.Exception.StackTrace}", Logger.LogLevel.Error)

                If e.Exception.InnerException IsNot Nothing Then
                    Logger.Log($"Inner Exception: {e.Exception.InnerException.Message}", Logger.LogLevel.Error)
                End If
            Catch
                ' Logger itself failed — nothing more we can do
            End Try
        End Sub

        ''' <summary>
        ''' Application startup — trim log file and log session start.
        ''' </summary>
        Private Sub MyApplication_Startup(
            sender As Object,
            e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs
        ) Handles Me.Startup

            Try
                Logger.TrimLogFile()
                Logger.Log("Application starting", Logger.LogLevel.Info)
            Catch
                ' Non-critical
            End Try
        End Sub

    End Class
End Namespace
