Imports System.IO

''' <summary>
''' Helper methods related to installer storage paths.
''' </summary>
Public Module InstallerPathService
    ''' <summary>
    ''' Ensure the update folder exists in ProgramData.
    ''' </summary>
    Public Sub EnsureUpdateFolderExists()
        Dim folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "VASTUpdater\NewPatchInstall")
        If Not Directory.Exists(folder) Then
            Directory.CreateDirectory(folder)
        End If
    End Sub

    ''' <summary>
    ''' Build the full path for a downloaded installer version.
    ''' </summary>
    Public Function GetInstallPath(version As String) As String
        Dim basePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "VASTUpdater\NewPatchInstall")
        If Not Directory.Exists(basePath) Then Directory.CreateDirectory(basePath)
        Return Path.Combine(basePath, $"{version}.exe")
    End Function
End Module
