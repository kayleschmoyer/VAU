Imports System.Drawing.Drawing2D
Imports System.Threading

''' <summary>
''' Main application form for the VAST Auto Updater.
''' Supports both interactive (UI) and silent (headless) update modes.
''' Borderless window with rounded corners, drop shadow, gradient header,
''' card layout, and animated brand-styled controls.
''' </summary>
Public Class MainForm

    Private engine As New UpdaterEngine()
    Private dashboardService As IDashboardService = New DashboardService()
    Private runSilently As Boolean
    Private updateCts As CancellationTokenSource = Nothing
    Private fadeTimer As Windows.Forms.Timer = Nothing

    ''' <summary>Visual state of the status line in the activity card.</summary>
    Private Enum StatusKind
        Ready
        Working
        Success
        Warning
        Failure
    End Enum

    Private Const CS_DROPSHADOW As Integer = &H20000

    ''' <summary>
    ''' Add a native drop shadow to the borderless window.
    ''' </summary>
    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ClassStyle = cp.ClassStyle Or CS_DROPSHADOW
            Return cp
        End Get
    End Property

    ''' <summary>
    ''' Request native rounded corners from DWM (Windows 11).
    ''' </summary>
    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)
        UiTheme.ApplyRoundedCorners(Me.Handle)
    End Sub

    ''' <summary>
    ''' Initializes the form, applies brand styling, and configures silent mode if requested.
    ''' Pass "silent" as a command-line argument to run headless.
    ''' </summary>
    Public Sub New()
        InitializeComponent()

        If Me.DesignMode Then Return

        Me.DoubleBuffered = True
        ApplyBranding()

        Dim args As String() = Environment.GetCommandLineArgs()
        runSilently = args.Any(Function(a) a.Equals("silent", StringComparison.OrdinalIgnoreCase))

        AddHandler Me.Load, AddressOf MainForm_Load
        AddHandler Me.FormClosing, AddressOf MainForm_FormClosing

        If runSilently Then
            Me.WindowState = FormWindowState.Minimized
            Me.ShowInTaskbar = False
            Me.Opacity = 0
            Logger.Log("Starting in silent mode", Logger.LogLevel.Info)
        Else
            ' Start invisible; MainForm_Load fades the window in
            Me.Opacity = 0
        End If
    End Sub

    ''' <summary>
    ''' Apply brand styling: gradient header, typography, icon glyphs,
    ''' placeholders, and event wiring.
    ''' </summary>
    Private Sub ApplyBranding()
        ' Gradient header paint
        AddHandler pnlHeader.Paint, AddressOf PaintGradientHeader

        ' Window drag support on header
        UiTheme.AttachDrag(Me, pnlHeader)
        UiTheme.AttachDrag(Me, lblTitle)
        UiTheme.AttachDrag(Me, lblSubtitle)

        ' Typography hierarchy within the Segoe UI family
        lblTitle.Font = UiTheme.Semibold(14.0F)
        btnCheckForUpdates.Font = UiTheme.Semibold(10.5F)
        lblPercent.Font = UiTheme.Semibold(9.5F)

        ' Window-control glyphs: Segoe MDL2 Assets when available
        If UiTheme.Mdl2Available Then
            Dim glyphFont As Font = UiTheme.IconFont(9.0F)
            btnClose.Font = glyphFont
            btnMinimize.Font = glyphFont
            btnSettings.Font = glyphFont
            btnClose.Text = ChrW(&HE8BB)      ' ChromeClose
            btnMinimize.Text = ChrW(&HE921)   ' ChromeMinimize
            btnSettings.Text = ChrW(&HE713)   ' Settings gear
        End If
        lblStatusIcon.Font = UiTheme.IconFont(10.0F)

        ' Input placeholders
        txtSftpUsername.CueText = "Enter SFTP username"
        txtSftpPassword.CueText = "Enter SFTP password"

        ' Close / minimize / settings / action
        AddHandler btnClose.Click, AddressOf BtnClose_Click
        AddHandler btnMinimize.Click, AddressOf BtnMinimize_Click
        AddHandler btnCheckForUpdates.Click, AddressOf BtnCheckForUpdates_Click
        AddHandler btnSettings.Click, AddressOf BtnSettings_Click

        SetStatus("Ready for update check", StatusKind.Ready)
    End Sub

    ''' <summary>
    ''' Paint the header panel with a horizontal magenta gradient.
    ''' Guards against zero-size panels during form initialization.
    ''' </summary>
    Private Sub PaintGradientHeader(sender As Object, e As PaintEventArgs)
        Dim pnl As Panel = DirectCast(sender, Panel)
        If pnl.ClientRectangle.Width <= 0 OrElse pnl.ClientRectangle.Height <= 0 Then Return
        Using brush As New LinearGradientBrush(
            pnl.ClientRectangle,
            UiTheme.Magenta,
            UiTheme.MagentaDark,
            LinearGradientMode.Horizontal)
            e.Graphics.FillRectangle(brush, pnl.ClientRectangle)
        End Using
    End Sub

    ''' <summary>
    ''' Update the status line, its color, and its icon glyph in one place.
    ''' </summary>
    Private Sub SetStatus(text As String, kind As StatusKind)
        lblStatus.Text = text
        Select Case kind
            Case StatusKind.Working
                lblStatus.ForeColor = UiTheme.Charcoal
                lblStatusIcon.ForeColor = UiTheme.Magenta
                lblStatusIcon.Text = If(UiTheme.Mdl2Available, ChrW(&HE895), "↻") ' Sync
            Case StatusKind.Success
                lblStatus.ForeColor = UiTheme.SuccessGreen
                lblStatusIcon.ForeColor = UiTheme.SuccessGreen
                lblStatusIcon.Text = If(UiTheme.Mdl2Available, ChrW(&HE73E), "✓") ' CheckMark
            Case StatusKind.Warning
                lblStatus.ForeColor = UiTheme.MagentaDark
                lblStatusIcon.ForeColor = UiTheme.MagentaDark
                lblStatusIcon.Text = If(UiTheme.Mdl2Available, ChrW(&HE946), "•") ' Info
            Case StatusKind.Failure
                lblStatus.ForeColor = UiTheme.ErrorRed
                lblStatusIcon.ForeColor = UiTheme.ErrorRed
                lblStatusIcon.Text = If(UiTheme.Mdl2Available, ChrW(&HE783), "!") ' Error badge
            Case Else
                lblStatus.ForeColor = UiTheme.Charcoal
                lblStatusIcon.ForeColor = UiTheme.Magenta
                lblStatusIcon.Text = If(UiTheme.Mdl2Available, ChrW(&HE946), "•") ' Info
        End Select
    End Sub

    ''' <summary>
    ''' Close the application with a successful exit code.
    ''' </summary>
    Private Sub BtnClose_Click(sender As Object, e As EventArgs)
        ExitApplication(0)
    End Sub

    ''' <summary>
    ''' Minimize the application window.
    ''' </summary>
    Private Sub BtnMinimize_Click(sender As Object, e As EventArgs)
        Me.WindowState = FormWindowState.Minimized
    End Sub

    ''' <summary>
    ''' Open the settings dialog for configuring email recipients.
    ''' </summary>
    Private Sub BtnSettings_Click(sender As Object, e As EventArgs)
        Using settingsForm As New SettingsForm()
            settingsForm.ShowDialog(Me)
        End Using
    End Sub

    ''' <summary>
    ''' Form load handler. Encrypts credentials on first run, validates configuration,
    ''' discovers the current VAST version, fades the window in, and starts the
    ''' silent update if applicable.
    ''' </summary>
    Private Sub MainForm_Load(sender As Object, e As EventArgs)
        Try
            ' Encrypt plaintext credentials on first run
            ConfigManager.EncryptOnFirstRun()

            ' Validate configuration and log any missing settings
            ConfigManager.ValidateConfiguration()

            Dim exePath As String = VersionService.FindVastExecutable()
            If Not String.IsNullOrEmpty(exePath) Then
                Dim version As String = VersionService.GetFileVersion(exePath)
                lblCurrentVersion.Text = $"Current version: {version}"
            Else
                lblCurrentVersion.Text = "Current version: not found"
            End If
            SetStatus("Ready for update check", StatusKind.Ready)
        Catch ex As Exception
            Logger.Log($"Error on load: {ex.Message}", Logger.LogLevel.Error)
            SetStatus("Error during load", StatusKind.Failure)
        End Try

        If runSilently Then
            ' Start silent update after form handle exists and message loop is running
            RunSilentUpdate()
        Else
            BeginFadeIn()
        End If
    End Sub

    ''' <summary>
    ''' Fade the window in over ~150 ms for a soft entrance.
    ''' </summary>
    Private Sub BeginFadeIn()
        fadeTimer = New Windows.Forms.Timer() With {.Interval = 15}
        AddHandler fadeTimer.Tick,
            Sub()
                Me.Opacity = Math.Min(1.0, Me.Opacity + 0.1)
                If Me.Opacity >= 1.0 Then
                    fadeTimer.Stop()
                    fadeTimer.Dispose()
                    fadeTimer = Nothing
                End If
            End Sub
        fadeTimer.Start()
    End Sub

    ''' <summary>
    ''' Handle button click: toggles between starting and cancelling an update check.
    ''' </summary>
    Private Async Sub BtnCheckForUpdates_Click(sender As Object, e As EventArgs)
        If updateCts IsNot Nothing Then
            ' Cancel in progress -- disable button briefly to prevent double-cancel
            btnCheckForUpdates.Enabled = False
            updateCts.Cancel()
            btnCheckForUpdates.Enabled = True
            Return
        End If

        updateCts = New CancellationTokenSource()
        btnCheckForUpdates.Text = "Cancel"
        Try
            Await RunUpdate(updateCts.Token)
        Finally
            updateCts.Dispose()
            updateCts = Nothing
            btnCheckForUpdates.Text = "Check for updates"
        End Try
    End Sub

    ''' <summary>
    ''' Send a heartbeat event to the dashboard so the machine reports in
    ''' on every silent run, even when no update is needed.
    ''' </summary>
    Private Sub SendHeartbeat()
        Try
            Dim exePath As String = VersionService.FindVastExecutable()
            Dim version As String = If(String.IsNullOrEmpty(exePath), String.Empty, VersionService.GetFileVersion(exePath))
            dashboardService.ReportEvent(DashboardEventType.Heartbeat, version)
        Catch ex As Exception
            Logger.Log($"Failed to send heartbeat: {ex.Message}", Logger.LogLevel.Warning)
        End Try
    End Sub

    ''' <summary>
    ''' Execute the update in silent (headless) mode. Exits with code 0 on success, 1 on failure.
    ''' Always sends a dashboard heartbeat, even if no update is needed.
    ''' </summary>
    Private Async Sub RunSilentUpdate()
        SendHeartbeat()
        Try
            Await RunUpdate(CancellationToken.None)
            Logger.Log("Silent mode finished successfully, exiting", Logger.LogLevel.Info)
            ExitApplication(0)
        Catch ex As Exception
            Logger.Log($"Silent mode failed: {ex.Message}", Logger.LogLevel.Error)
            ExitApplication(1)
        End Try
    End Sub

    ''' <summary>
    ''' Route an engine progress callback into the activity card: status line,
    ''' percentage, detail line, progress bar, and update-availability badge.
    ''' </summary>
    Private Sub ApplyProgress(p As Integer, status As String)
        progressBar1.Value = Math.Min(Math.Max(p, 0), 100)
        lblPercent.Text = If(p > 0 AndAlso p < 100, $"{p}%", String.Empty)

        If status.StartsWith("Downloading... ", StringComparison.Ordinal) Then
            ' Byte-level download ticks go to the muted detail line
            lblProgressDetail.Text = status.Substring("Downloading... ".Length)
            Return
        End If

        SetStatus(status, StatusKind.Working)

        If status.StartsWith("Downloading version", StringComparison.OrdinalIgnoreCase) Then
            badgeUpdate.SetState("Update available", UiTheme.PinkPale, UiTheme.MagentaDark)
        ElseIf status.Contains("No update available") OrElse status.Contains("up-to-date") Then
            badgeUpdate.SetState("Up to date", UiTheme.SuccessPale, UiTheme.SuccessGreen)
        End If
    End Sub

    ''' <summary>
    ''' Core update workflow: validates credentials, invokes the engine, and updates UI with progress.
    ''' Handles <see cref="OperationCanceledException"/>, <see cref="UpdateException"/>, and general exceptions.
    ''' </summary>
    ''' <param name="cancelToken">Token to cancel the update mid-operation.</param>
    Private Async Function RunUpdate(cancelToken As CancellationToken) As Task
        Dim user As String
        Dim pass As String

        If runSilently Then
            user = ConfigManager.SftpUsername
            pass = ConfigManager.SftpPassword
            Logger.Log("Silent mode: using credentials from configuration", Logger.LogLevel.Info)
        Else
            user = txtSftpUsername.Text
            pass = txtSftpPassword.Text
        End If

        If String.IsNullOrWhiteSpace(user) OrElse String.IsNullOrWhiteSpace(pass) Then
            Logger.Log("Credentials required but not available", Logger.LogLevel.Warning)
            If runSilently Then
                Throw New InvalidOperationException("SFTP credentials are not configured")
            End If
            SetStatus("Enter SFTP credentials above", StatusKind.Warning)
            Return
        End If

        If Not runSilently Then
            progressBar1.Value = 0
            lblProgressDetail.Text = String.Empty
            SetStatus("Starting update check...", StatusKind.Working)
        End If

        Try
            Await engine.PerformUpdateCheck(user, pass,
                Sub(p As Integer, status As String)
                    If runSilently Then Return
                    Try
                        Me.Invoke(Sub() ApplyProgress(p, status))
                    Catch
                        ' Form may be disposed during exit
                    End Try
                End Sub, cancelToken)

            If Not runSilently Then
                Try
                    Me.Invoke(Sub()
                                  SetStatus("Update complete", StatusKind.Success)
                                  lblPercent.Text = String.Empty
                                  lblProgressDetail.Text = String.Empty
                              End Sub)
                Catch
                    ' Form may be disposed during exit
                End Try
                Logger.Log("Update completed in UI mode.", Logger.LogLevel.Info)
            End If

        Catch ex As OperationCanceledException
            Logger.Log("Update check cancelled by user", Logger.LogLevel.Info)
            If Not runSilently Then
                Try
                    Me.Invoke(Sub()
                                  SetStatus("Update check cancelled", StatusKind.Warning)
                                  lblPercent.Text = String.Empty
                                  lblProgressDetail.Text = String.Empty
                                  progressBar1.Value = 0
                              End Sub)
                Catch
                End Try
            End If

        Catch ex As UpdateException
            Logger.Log($"Update failed [{ex.ErrorCode}]: {ex.Message}", Logger.LogLevel.Error)

            If Not runSilently Then
                Try
                    Me.Invoke(Sub()
                                  SetStatus($"Error [{ex.ErrorCode}]: {ex.Message}", StatusKind.Failure)
                                  lblPercent.Text = String.Empty
                                  lblProgressDetail.Text = String.Empty
                                  progressBar1.Value = 0
                              End Sub)
                Catch
                End Try
            End If

            If runSilently Then Throw

        Catch ex As Exception
            Logger.Log($"Update failed: {ex.Message}", Logger.LogLevel.Error)

            If Not runSilently Then
                Try
                    Me.Invoke(Sub()
                                  SetStatus($"Error: {ex.Message}", StatusKind.Failure)
                                  lblPercent.Text = String.Empty
                                  lblProgressDetail.Text = String.Empty
                                  progressBar1.Value = 0
                              End Sub)
                Catch
                End Try
            End If

            If runSilently Then Throw
        End Try
    End Function

    ''' <summary>
    ''' Cancel any in-progress update and log application closure.
    ''' </summary>
    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs)
        If updateCts IsNot Nothing Then
            updateCts.Cancel()
            Logger.Log("Cancelled in-progress update due to form closing", Logger.LogLevel.Info)
        End If
        Logger.Log("Application closing", Logger.LogLevel.Info)
    End Sub

    ''' <summary>
    ''' Log and exit the application with the specified exit code.
    ''' </summary>
    ''' <param name="exitCode">Process exit code (0 = success, non-zero = failure).</param>
    Private Sub ExitApplication(exitCode As Integer)
        Logger.Log($"Application exiting with code {exitCode}", Logger.LogLevel.Info)
        If exitCode <> 0 Then
            Environment.ExitCode = exitCode
        End If
        Application.Exit()
    End Sub

End Class
