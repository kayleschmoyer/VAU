Imports System.Drawing.Drawing2D
Imports System.Threading

''' <summary>
''' Main application form for the VAST Auto Updater.
''' Supports both interactive (UI) and silent (headless) update modes.
''' Uses a borderless window with custom gradient header and brand styling.
''' </summary>
Public Class MainForm

    Private engine As New UpdaterEngine()
    Private runSilently As Boolean
    Private isDragging As Boolean = False
    Private dragStart As Point
    Private updateCts As CancellationTokenSource = Nothing

    ' Brand colors
    Private Shared ReadOnly Magenta As Color = Color.FromArgb(237, 1, 127)
    Private Shared ReadOnly MagentaDark As Color = Color.FromArgb(180, 0, 96)
    Private Shared ReadOnly Charcoal As Color = Color.FromArgb(51, 51, 51)

    ''' <summary>
    ''' Initializes the form, applies brand styling, and configures silent mode if requested.
    ''' Pass "silent" as a command-line argument to run headless.
    ''' </summary>
    Public Sub New()
        InitializeComponent()
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

    ''' <summary>
    ''' Paint the header panel with a horizontal magenta gradient.
    ''' Guards against zero-size panels during form initialization.
    ''' </summary>
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

    ''' <summary>
    ''' Begin window drag when the user presses the left mouse button on the header.
    ''' </summary>
    Private Sub Header_MouseDown(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            isDragging = True
            dragStart = e.Location
        End If
    End Sub

    ''' <summary>
    ''' Move the window to follow the mouse cursor during a drag operation.
    ''' </summary>
    Private Sub Header_MouseMove(sender As Object, e As MouseEventArgs)
        If isDragging Then
            Dim ctrl As Control = DirectCast(sender, Control)
            Dim screenPoint As Point = ctrl.PointToScreen(e.Location)
            Me.Location = New Point(screenPoint.X - dragStart.X, screenPoint.Y - dragStart.Y)
        End If
    End Sub

    ''' <summary>
    ''' End the window drag operation on mouse button release.
    ''' </summary>
    Private Sub Header_MouseUp(sender As Object, e As MouseEventArgs)
        isDragging = False
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

    ' ── Progress bar color hack ──────────────────────────────────────
    ''' <summary>
    ''' Apply magenta color to the progress bar via SendMessage PBM_SETBARCOLOR.
    ''' </summary>
    Private Sub SetProgressBarColor()
        ' Use Win32 message to set progress bar color to magenta
        SendMessage(progressBar1.Handle, &H409, IntPtr.Zero, New IntPtr(ColorTranslator.ToWin32(Magenta)))
    End Sub

    ''' <summary>
    ''' Win32 SendMessage for setting progress bar color (PBM_SETBARCOLOR).
    ''' </summary>
    <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Auto)>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function

    ''' <summary>
    ''' Form load handler. Encrypts credentials on first run, validates configuration,
    ''' discovers the current VAST version, and starts the silent update if applicable.
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

    ''' <summary>
    ''' Handle button click: toggles between starting and cancelling an update check.
    ''' </summary>
    Private Async Sub BtnCheckForUpdates_Click(sender As Object, e As EventArgs)
        If updateCts IsNot Nothing Then
            ' Cancel in progress — disable button briefly to prevent double-cancel
            btnCheckForUpdates.Enabled = False
            updateCts.Cancel()
            btnCheckForUpdates.Enabled = True
            Return
        End If

        updateCts = New CancellationTokenSource()
        btnCheckForUpdates.Text = "CANCEL"
        Try
            Await RunUpdate(updateCts.Token)
        Finally
            updateCts.Dispose()
            updateCts = Nothing
            btnCheckForUpdates.Text = "CHECK FOR UPDATES"
        End Try
    End Sub

    ''' <summary>
    ''' Execute the update in silent (headless) mode. Exits with code 0 on success, 1 on failure.
    ''' </summary>
    Private Async Sub RunSilentUpdate()
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
                                      lblStatus.Text = st