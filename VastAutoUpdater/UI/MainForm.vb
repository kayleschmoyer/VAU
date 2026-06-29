Imports System.Drawing.Drawing2D

Public Class MainForm

    Private engine As New UpdaterEngine()
    Private runSilently As Boolean
    Private isDragging As Boolean = False
    Private dragStart As Point

    ' Brand colors
    Private Shared ReadOnly Magenta As Color = Color.FromArgb(237, 1, 127)
    Private Shared ReadOnly MagentaDark As Color = Color.FromArgb(180, 0, 96)
    Private Shared ReadOnly Charcoal As Color = Color.FromArgb(51, 51, 51)


    Public Sub New()
        InitializeComponent()
        ApplyBranding()

        Dim args As String() = Environment.GetCommandLineArgs()
        runSilently = args.Contains("silent")

        AddHandler Me.Load, AddressOf MainForm_Load
        AddHandler Me.FormClosing, AddressOf MainForm_FormClosing

        If runSilently Then
            Me.WindowState = FormWindowState.Minimized
            Me.ShowInTaskbar = False
            Me.Opacity = 0
            Logger.Log("Starting in silent mode", Logger.LogLevel.Info)
        End If
    End Sub

    ''' <summary>
    ''' Apply brand styling and gradient painting.
    ''' </summary>
    Private Sub ApplyBranding()
        ' Gradient header paint
        AddHandler pnlHeader.Paint, AddressOf PaintGradientHeader

        ' Window drag support on header
        AddHandler pnlHeader.MouseDown, AddressOf Header_MouseDown
        AddHandler pnlHeader.MouseMove, AddressOf Header_MouseMove
        AddHandler pnlHeader.MouseUp, AddressOf Header_MouseUp
        AddHandler lblTitle.MouseDown, AddressOf Header_MouseDown
        AddHandler lblTitle.MouseMove, AddressOf Header_MouseMove
        AddHandler lblTitle.MouseUp, AddressOf Header_MouseUp
        AddHandler lblSubtitle.MouseDown, AddressOf Header_MouseDown
        AddHandler lblSubtitle.MouseMove, AddressOf Header_MouseMove
        AddHandler lblSubtitle.MouseUp, AddressOf Header_MouseUp

        ' Close / minimize
        AddHandler btnClose.Click, AddressOf BtnClose_Click
        AddHandler btnMinimize.Click, AddressOf BtnMinimize_Click

        ' Credential panel no longer needs custom paint

        ' Custom progress bar color
        SetProgressBarColor()

        ' Button handler
        AddHandler btnCheckForUpdates.Click, AddressOf BtnCheckForUpdates_Click
    End Sub

    ' ── Gradient header ──────────────────────────────────────────────

    Private Sub PaintGradientHeader(sender As Object, e As PaintEventArgs)
        Dim pnl As Panel = DirectCast(sender, Panel)
        If pnl.ClientRectangle.Width <= 0 OrElse pnl.ClientRectangle.Height <= 0 Then Return
        Using brush As New LinearGradientBrush(
            pnl.ClientRectangle,
            Magenta,
            MagentaDark,
            LinearGradientMode.Horizontal)
            e.Graphics.FillRectangle(brush, pnl.ClientRectangle)
        End Using
    End Sub

    ' ── Window drag ──────────────────────────────────────────────────

    Private Sub Header_MouseDown(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            isDragging = True
            dragStart = e.Location
        End If
    End Sub

    Private Sub Header_MouseMove(sender As Object, e As MouseEventArgs)
        If isDragging Then
            Dim ctrl As Control = DirectCast(sender, Control)
            Dim screenPoint As Point = ctrl.PointToScreen(e.Location)
            Me.Location = New Point(screenPoint.X - dragStart.X, screenPoint.Y - dragStart.Y)
        End If
    End Sub

    Private Sub Header_MouseUp(sender As Object, e As MouseEventArgs)
        isDragging = False
    End Sub

    ' ── Window controls ──────────────────────────────────────────────

    Private Sub BtnClose_Click(sender As Object, e As EventArgs)
        ExitApplication(0)
    End Sub

    Private Sub BtnMinimize_Click(sender As Object, e As EventArgs)
        Me.WindowState = FormWindowState.Minimized
    End Sub

    ' ── Progress bar color hack ──────────────────────────────────────
    ''' <summary>
    ''' Apply magenta color to the progress bar via SendMessage PBM_SETBARCOLOR.
    ''' </summary>
    Private Sub SetProgressBarColor()
        ' Use Win32 message to set progress bar color to magenta
        SendMessage(progressBar1.Handle, &H409, IntPtr.Zero, New IntPtr(ColorTranslator.ToWin32(Magenta)))
    End Sub

    <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Auto)>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function

    ' ── Load ─────────────────────────────────────────────────────────

    Private Sub MainForm_Load(sender As Object, e As EventArgs)
        Try
            ' Encrypt plaintext credentials on first run
            ConfigManager.EncryptOnFirstRun()

            Dim exePath As String = VersionService.FindVastExecutable()
            If Not String.IsNullOrEmpty(exePath) Then
                Dim version As String = VersionService.GetFileVersion(exePath)
                lblCurrentVersion.Text = $"Current Version: {version}"
            Else
                lblCurrentVersion.Text = "Current Version: Not Found"
            End If
            lblStatus.Text = "Ready for update check..."
        Catch ex As Exception
            Logger.Log($"Error on load: {ex.Message}", Logger.LogLevel.Error)
            lblStatus.Text = "Error during load"
        End Try

        ' Start silent update after form handle exists and message loop is running
        If runSilently Then
            RunSilentUpdate()
        End If
    End Sub

    ' ── Update logic ─────────────────────────────────────────────────

    Private Async Sub BtnCheckForUpdates_Click(sender As Object, e As EventArgs)
        btnCheckForUpdates.Enabled = False
        btnCheckForUpdates.Text = "CHECKING..."
        Await RunUpdate()
        btnCheckForUpdates.Text = "CHECK FOR UPDATES"
        btnCheckForUpdates.Enabled = True
    End Sub

    Private Async Sub RunSilentUpdate()
        Try
            Await RunUpdate()
            Logger.Log("Silent mode finished successfully, exiting", Logger.LogLevel.Info)
            ExitApplication(0)
        Catch ex As Exception
            Logger.Log($"Silent mode failed: {ex.Message}", Logger.LogLevel.Error)
            ExitApplication(1)
        End Try
    End Sub

    Private Async Function RunUpdate() As Task
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
            lblStatus.Text = "Enter SFTP credentials above."
            lblStatus.ForeColor = Color.FromArgb(200, 0, 80)
            Return
        End If

        If Not runSilently Then
            lblStatus.ForeColor = Charcoal
            pnlProgress.Visible = True
            lblStatus.Text = "Starting update check..."
        End If

        Try
            Await engine.PerformUpdateCheck(user, pass,
                Sub(p As Integer, status As String)
                    If runSilently Then Return
                    Try
                        Me.Invoke(Sub()
                                      lblStatus.Text = status
                                      progressBar1.Value = Math.Min(Math.Max(p, 0), 100)
                                      progressBar1.Invalidate()
                                  End Sub)
                    Catch
                        ' Form may be disposed during exit
                    End Try
                End Sub)

            If Not runSilently Then
                Try
                    Me.Invoke(Sub()
                                  lblStatus.Text = "Update complete."
                                  lblStatus.ForeColor = Color.FromArgb(0, 150, 80)
                                  pnlProgress.Visible = False
                              End Sub)
                Catch
                    ' Form may be disposed during exit
                End Try
                Logger.Log("Update completed in UI mode.", Logger.LogLevel.Info)
            End If

        Catch ex As Exception
            Logger.Log($"Update failed: {ex.Message}", Logger.LogLevel.Error)

            If Not runSilently Then
                Try
                    Me.Invoke(Sub()
                                  lblStatus.Text = $"Error: {ex.Message}"
                                  lblStatus.ForeColor = Color.FromArgb(200, 0, 0)
                                  pnlProgress.Visible = False
                              End Sub)
                Catch
                End Try
            End If

            If runSilently Then Throw
        End Try
    End Function

    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs)
        Logger.Log("Application closing", Logger.LogLevel.Info)
    End Sub

    Private Sub ExitApplication(exitCode As Integer)
        Logger.Log($"Application exiting with code {exitCode}", Logger.LogLevel.Info)
        If exitCode <> 0 Then
            Environment.ExitCode = exitCode
     