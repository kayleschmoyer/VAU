Imports MaterialSkin
Imports MaterialSkin.Controls

Public Class VASTUpdater
    Inherits MaterialForm

    Private engine As New UpdaterEngine()

    Private txtSftpUsername As MaterialTextBox
    Private txtSftpPassword As MaterialTextBox
    Private btnCheckForUpdates As MaterialButton
    Private progressBar1 As MaterialProgressBar
    Private lblStatus As MaterialLabel
    Private runSilently As Boolean

    Public Sub New()
        InitializeComponent()
        InitializeUX()

        Dim args = Environment.GetCommandLineArgs()
        runSilently = args.Contains("silent")
        If runSilently Then
            Logger.Log("Starting in silent mode", Logger.LogLevel.Info)
            RunSilentUpdate()
        End If
    End Sub

    Private Sub InitializeUX()
        Me.Text = "VAST Updater"
        Me.Size = New Size(600, 400)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False

        Dim centerX As Integer = (Me.ClientSize.Width - 300) \ 2

        txtSftpUsername = New MaterialTextBox With {
            .Hint = "SFTP Username",
            .Size = New Size(300, 50),
            .Location = New Point(centerX, 100)
        }
        Me.Controls.Add(txtSftpUsername)

        txtSftpPassword = New MaterialTextBox With {
            .Hint = "SFTP Password",
            .Password = True,
            .Size = New Size(300, 50),
            .Location = New Point(centerX, 160)
        }
        Me.Controls.Add(txtSftpPassword)

        btnCheckForUpdates = New MaterialButton With {
            .Text = "Check",
            .Size = New Size(200, 40),
            .Location = New Point(centerX + 50, 230)
        }
        AddHandler btnCheckForUpdates.Click, AddressOf btnCheckForUpdates_Click
        Me.Controls.Add(btnCheckForUpdates)

        progressBar1 = New MaterialProgressBar With {
            .Size = New Size(300, 10),
            .Location = New Point(centerX, 280),
            .Visible = False
        }
        Me.Controls.Add(progressBar1)

        lblStatus = New MaterialLabel With {
            .Text = "Ready",
            .Location = New Point(centerX, 300),
            .Size = New Size(300, 30)
        }
        Me.Controls.Add(lblStatus)
    End Sub

    Private Async Sub btnCheckForUpdates_Click(sender As Object, e As EventArgs)
        Await RunUpdate()
    End Sub

    Private Async Sub RunSilentUpdate()
        Await RunUpdate()
        Me.Close()
    End Sub

    Private Async Function RunUpdate() As Task
        Dim user = txtSftpUsername.Text
        Dim pass = txtSftpPassword.Text
        If String.IsNullOrWhiteSpace(user) OrElse String.IsNullOrWhiteSpace(pass) Then
            Logger.Log("Credentials required", Logger.LogLevel.Warning)
            If Not runSilently Then MessageBox.Show("Enter SFTP credentials")
            Return
        End If
        progressBar1.Visible = True
        lblStatus.Text = "Checking..."
        Await engine.PerformUpdateCheck(user, pass, Sub(p, t)
                                                        progressBar1.Value = 100
                                                   End Sub)
        lblStatus.Text = "Finished"
        progressBar1.Visible = False
    End Function
End Class
