Imports System.Diagnostics
Imports System.IO

''' <summary>
''' Utility methods for locating VAST.exe and obtaining version information.
''' Caches the discovered path to avoid repeated full-drive scans.
''' </summary>
Public Module VersionService
    Private _cachedVastPath As String = Nothing

    ''' <summary>
    ''' Search all drives for the VAST executable. Result is cached after first discovery.
    ''' </summary>
    Public Function FindVastExecutable() As String
        ' Return cached path if still valid
        If _cachedVastPath IsNot Nothing AndAlso File.Exists(_cachedVastPath) Then
            Return _cachedVastPath
        End If

        Dim potential As String() = {
            "Program Files (x86)\MAM Software\VAST\VAST.exe",
            "Program Files\MAM Software\VAST\VAST.exe"
        }

        For Each drive In DriveInfo.GetDrives()
            If drive.IsReady AndAlso drive.DriveType = DriveType.Fixed Then
                For Each rel In potential
                    Dim full As String = Path.Combine(drive.Name, rel)
                    If File.Exists(full) Then
                        Logger.Log($"VAST executable found at: {full}", Logger.LogLevel.Info)
                        _cachedVastPath = full
                        Return full
                    End If
                Next
            End If
        Next

        Logger.Log("No VAST.exe found on any fixed drive", Logger.LogLevel.Warning)
        Return String.Empty
    End Function

    ''' <summary>
    ''' Return the product version string of the provided file.
    ''' Returns "0.0.0" if the file is missing or version cannot be read.
    ''' </summary>
    Public Function GetFileVersion(filePath As String) As String
        Try
            If File.Exists(filePath) Then
                Dim ver As String = FileVersionInfo.GetVersionInfo(filePath).ProductVersion
                If Not String.IsNullOrEmpty(ver) Then Return ver
            End If
        Catch ex As Exception
            Logger.Log($"Error retrieving file version for '{filePath}': {ex.Message}", Logger.LogLevel.Error)
        End Try
        Return "0.0.0"
    End Function
End Module
