Imports System.Diagnostics
Imports System.IO

''' <summary>
''' Utility methods for locating VAST.exe and obtaining version information.
''' Extracted from <see cref="UpdaterEngine"/>.
''' </summary>
Public Module VersionService
    ''' <summary>
    ''' Search all drives for the VAST executable.
    ''' </summary>
    Public Function FindVastExecutable() As String
        Dim potential As String() = {
            "Program Files (x86)\MAM Software\VAST\VAST.exe",
            "Program Files\MAM Software\VAST\VAST.exe"
        }
        For Each drive In DriveInfo.GetDrives()
            If drive.IsReady Then
                For Each rel In potential
                    Dim full = Path.Combine(drive.Name, rel)
                    If File.Exists(full) Then
                        Logger.Log($"VAST executable found at: {full}", Logger.LogLevel.Info)
                        Return full
                    End If
                Next
            End If
        Next
        Logger.Log("No VAST.exe found", Logger.LogLevel.Warning)
        Return String.Empty
    End Function

    ''' <summary>
    ''' Return the file version of the provided path.
    ''' </summary>
    Public Function GetFileVersion(filePath As String) As String
        Try
            If File.Exists(filePath) Then
                Return FileVersionInfo.GetVersionInfo(filePath).ProductVersion
            End If
        Catch ex As Exception
            Logger.Log($"Error retrieving file version: {ex.Message}", Logger.LogLevel.Error)
        End Try
        Return "0.0.0"
    End Function
End Module
