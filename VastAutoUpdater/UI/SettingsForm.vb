Option Strict On

Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Windows.Forms

''' <summary>
''' Settings dialog for configuring the SQL Server instance, customer identity,
''' and site number. The Test button queries VastOffice and auto-fills
''' Customer Name and Site Number from the COMPANY table.
''' Styled to match the main window: borderless, rounded, card layout.
''' </summary>
Public Class SettingsForm
    Inherits Form

    Private txtDatabaseServer As ModernTextBox
    Private txtCustomerName As ModernTextBox
    Private txtSiteNumber As ModernTextBox
    Private btnTestConnection As RoundedButton
    Private btnSave As RoundedButton
    Private btnCancel As RoundedButton
    Private btnClose As Button
    Private pnlHeader As Panel
    Private lblFormTitle As Label
    Private card As CardPanel
    Private lblDatabaseServer As Label
    Private lblDbHint As Label
    Private lblCustomerName As Label
    Private lblSiteNumber As Label

    Private Const CS_DROPSHADOW As Integer = &H20000

    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ClassStyle = cp.ClassStyle Or CS_DROPSHADOW
            Return cp
        End Get
    End Property

    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)
        UiTheme.ApplyRoundedCorners(Me.Handle)
    End Sub

    Public Sub New()
        InitializeSettingsForm()
        LoadSettings()
    End Sub

    Private Sub InitializeSettingsForm()
        Me.SuspendLayout()
        Me.Text = "Settings"
        Me.ClientSize = New Size(440, 392)
        Me.FormBorderStyle = FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = UiTheme.Canvas
        Me.Font = New Font("Segoe UI", 9.0!)

        ' --- Header ---
        ' Size must be set before anchored children are added, or the
        ' right-anchored close button ends up positioned off-screen
        pnlHeader = New Panel() With {
            .BackColor = UiTheme.Magenta,
            .Size = New Size(440, 48),
            .Dock = DockStyle.Top
        }

        lblFormTitle = New Label() With {
            .Text = "Settings",
            .Font = UiTheme.Semibold(12.0F),
            .ForeColor = Color.White,
            .BackColor = Color.Transparent,
            .Location = New Point(18, 12),
            .AutoSize = True
        }

        btnClose = New Button() With {
            .FlatStyle = FlatStyle.Flat,
            .BackColor = Color.Transparent,
            .ForeColor = Color.White,
            .Font = If(UiTheme.Mdl2Available, UiTheme.IconFont(9.0F), New Font("Segoe UI", 11.0!)),
            .Text = If(UiTheme.Mdl2Available, ChrW(&HE8BB), "X"),
            .Location = New Point(398, 9),
            .Size = New Size(32, 30),
            .TabStop = False,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right
        }
        btnClose.FlatAppearance.BorderSize = 0
        btnClose.FlatAppearance.MouseOverBackColor = UiTheme.MagentaDark
        btnClose.FlatAppearance.MouseDownBackColor = UiTheme.MagentaDeep

        pnlHeader.Controls.Add(lblFormTitle)
        pnlHeader.Controls.Add(btnClose)

        ' --- Card with the form fields ---
        card = New CardPanel() With {
            .Location = New Point(20, 68),
            .Size = New Size(400, 250)
        }

        lblDatabaseServer = MakeFieldLabel("DATABASE SERVER", New Point(18, 16))

        txtDatabaseServer = New ModernTextBox() With {
            .Location = New Point(18, 36),
            .Size = New Size(268, 36),
            .CueText = "e.g. .\SQLEXPRESS"
        }

        btnTestConnection = New RoundedButton() With {
            .Text = "Test",
            .Location = New Point(294, 36),
            .Size = New Size(88, 36),
            .CornerRadius = 8,
            .AccentColor = UiTheme.SuccessGreen,
            .AccentDarkColor = Color.FromArgb(0, 120, 64),
            .Font = UiTheme.Semibold(9.0F)
        }

        lblDbHint = New Label() With {
            .Text = "Examples: .\SQLEXPRESS or SERVERNAME\SQLEXPRESS",
            .Location = New Point(18, 80),
            .Size = New Size(364, 16),
            .Font = New Font("Segoe UI", 8.0!, FontStyle.Italic),
            .ForeColor = UiTheme.TextMuted,
            .BackColor = Color.White
        }

        lblCustomerName = MakeFieldLabel("CUSTOMER NAME", New Point(18, 106))

        txtCustomerName = New ModernTextBox() With {
            .Location = New Point(18, 126),
            .Size = New Size(364, 36)
        }

        lblSiteNumber = MakeFieldLabel("SITE NUMBER", New Point(18, 176))

        txtSiteNumber = New ModernTextBox() With {
            .Location = New Point(18, 196),
            .Size = New Size(364, 36)
        }

        card.Controls.AddRange({lblDatabaseServer, txtDatabaseServer, btnTestConnection, lblDbHint,
                                lblCustomerName, txtCustomerName,
                                lblSiteNumber, txtSiteNumber})

        ' --- Action buttons ---
        btnSave = New RoundedButton() With {
            .Text = "Save",
            .Location = New Point(230, 334),
            .Size = New Size(92, 40),
            .Font = UiTheme.Semibold(9.5F)
        }

        btnCancel = New RoundedButton() With {
            .Text = "Cancel",
            .Location = New Point(330, 334),
            .Size = New Size(90, 40),
            .AccentColor = Color.FromArgb(228, 228, 234),
            .AccentDarkColor = Color.FromArgb(206, 206, 214),
            .ForeColor = UiTheme.Charcoal,
            .Font = UiTheme.Semibold(9.5F)
        }

        ' --- Footer accent ---
        Dim pnlFooter As New Panel() With {
            .BackColor = UiTheme.Magenta,
            .Dock = DockStyle.Bottom,
            .Height = 3
        }

        AddHandler btnTestConnection.Click, AddressOf BtnTestConnection_Click
        AddHandler btnSave.Click, AddressOf BtnSave_Click
        AddHandler btnCancel.Click, AddressOf BtnCancel_Click
        AddHandler btnClose.Click, AddressOf BtnCancel_Click

        UiTheme.AttachDrag(Me, pnlHeader)
        UiTheme.AttachDrag(Me, lblFormTitle)

        Me.Controls.AddRange({card, btnSave, btnCancel, pnlFooter, pnlHeader})
        Me.AcceptButton = btnSave
        Me.CancelButton = btnCancel

        ' AutoScale properties must be set after controls are added (designer
        ' ordering) for per-monitor DPI scaling to apply at ResumeLayout
        Me.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
        Me.AutoScaleMode = AutoScaleMode.Dpi
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    ''' <summary>
    ''' Small muted field label used above each input.
    ''' </summary>
    Private Shared Function MakeFieldLabel(text As String, location As Point) As Label
        Return New Label() With {
            .Text = text,
            .Location = location,
            .AutoSize = True,
            .Font = New Font("Segoe UI", 8.0!, FontStyle.Bold),
            .ForeColor = UiTheme.TextMuted,
            .BackColor = Color.White
        }
    End Function

    Private Sub LoadSettings()
        Try
            txtDatabaseServer.Text = ConfigManager.DatabaseServer
            txtCustomerName.Text = ConfigurationManager.AppSettings("CustomerName")
            txtSiteNumber.Text = ConfigurationManager.AppSettings("SiteName")
        Catch ex As Exception
            Logger.Log("Error loading settings: " & ex.Message, Logger.LogLevel.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Test the SQL connection and auto-fill Customer Name and Site Number from the query result.
    ''' </summary>
    Private Sub BtnTestConnection_Click(sender As Object, e As EventArgs)
        Dim server As String = txtDatabaseServer.Text.Trim()
        If String.IsNullOrWhiteSpace(server) Then
            ShowConnectionHelp("No Server Entered",
                "Please type a database server name in the field above before testing.",
                "This is the name of the computer running SQL Server." & vbCrLf & vbCrLf &
                "Common examples:" & vbCrLf &
                "   .\SQLEXPRESS   (this computer, SQL Express)" & vbCrLf &
                "   MYSERVER\SQLEXPRESS   (another computer)" & vbCrLf &
                "   MYSERVER   (full SQL Server)")
            Return
        End If

        btnTestConnection.Enabled = False
        btnTestConnection.Text = "..."
        Application.DoEvents()

        Try
            Dim connStr As String = "Server=" & server & ";Database=VastOffice;User Id=vastnr;Password=snowdrift;Connection Timeout=5;ApplicationIntent=ReadOnly;"
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using cmd As New SqlCommand("SELECT TOP 1 NAME, COMPANY_NUMBER FROM COMPANY", conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Dim compName As String = If(reader.IsDBNull(0), String.Empty, reader.GetString(0).Trim())
                            Dim compNum As String = If(reader.IsDBNull(1), String.Empty, reader(1).ToString().Trim())

                            ' Auto-fill the fields
                            txtCustomerName.Text = compName
                            txtSiteNumber.Text = compNum

                            ShowConnectionSuccess(compName, compNum)
                        Else
                            ShowConnectionHelp("No Company Data Found",
                                "Connected to the database, but the COMPANY table has no records.",
                                "The VastOffice database exists but hasn't been set up with company information yet." & vbCrLf & vbCrLf &
                                "You can type the Customer Name and Site Number manually, or contact your administrator to verify the database is configured.")
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ShowConnectionError(server, ex)
        Finally
            btnTestConnection.Enabled = True
            btnTestConnection.Text = "Test"
        End Try
    End Sub

    ''' <summary>
    ''' Show a branded success dialog with the retrieved company info.
    ''' </summary>
    Private Sub ShowConnectionSuccess(compName As String, compNum As String)
        Const CONTENT_WIDTH As Integer = 380
        Const MARGIN As Integer = 24
        Const PADDING As Integer = 12

        Using dlg As New Form()
            dlg.Text = "Connection Successful"
            dlg.FormBorderStyle = FormBorderStyle.FixedDialog
            dlg.MaximizeBox = False
            dlg.MinimizeBox = False
            dlg.StartPosition = FormStartPosition.CenterParent
            dlg.BackColor = Color.White
            dlg.Font = New Font("Segoe UI", 9.0!)

            ' Green accent bar
            Dim pnlAccent As New Panel() With {
                .BackColor = UiTheme.SuccessGreen,
                .Dock = DockStyle.Top,
                .Height = 4
            }

            Dim lblTitle As New Label() With {
                .Text = "Connection Successful",
                .Font = UiTheme.Semibold(14.0F),
                .ForeColor = UiTheme.SuccessGreen,
                .MaximumSize = New Size(CONTENT_WIDTH, 0),
                .AutoSize = True
            }

            Dim lblDetails As New Label() With {
                .Text = "Customer Name:  " & compName & vbCrLf & "Site Number:  " & compNum,
                .Font = New Font("Segoe UI", 11.0!),
                .ForeColor = UiTheme.Charcoal,
                .MaximumSize = New Size(CONTENT_WIDTH, 0),
                .AutoSize = True
            }

            Dim lblNote As New Label() With {
                .Text = "These values have been filled in below. Click Save to keep them.",
                .Font = New Font("Segoe UI", 9.0!, FontStyle.Italic),
                .ForeColor = Color.FromArgb(100, 100, 100),
                .MaximumSize = New Size(CONTENT_WIDTH, 0),
                .AutoSize = True
            }

            Dim btnOk As New RoundedButton() With {
                .Text = "OK",
                .Size = New Size(84, 36),
                .AccentColor = UiTheme.SuccessGreen,
                .AccentDarkColor = Color.FromArgb(0, 120, 64),
                .Font = UiTheme.Semibold(9.0F),
                .DialogResult = DialogResult.OK
            }

            dlg.Controls.AddRange({pnlAccent, lblTitle, lblDetails, lblNote, btnOk})

            dlg.SuspendLayout()
            Dim y As Integer = MARGIN
            lblTitle.Location = New Point(MARGIN, y)
            y += lblTitle.Height + PADDING

            lblDetails.Location = New Point(MARGIN, y)
            y += lblDetails.Height + PADDING

            lblNote.Location = New Point(MARGIN, y)
            y += lblNote.Height + PADDING + 8

            btnOk.Location = New Point(MARGIN + CONTENT_WIDTH - btnOk.Width, y)
            y += btnOk.Height + MARGIN

            dlg.ClientSize = New Size(MARGIN * 2 + CONTENT_WIDTH, y)
            dlg.AcceptButton = btnOk
            dlg.ResumeLayout(True)

            dlg.ShowDialog(Me)
        End Using
    End Sub

    ''' <summary>
    ''' Show a friendly, branded error dialog with guidance based on the exception type.
    ''' </summary>
    Private Sub ShowConnectionError(server As String, ex As Exception)
        Dim title As String
        Dim whatHappened As String
        Dim whatToDo As String

        Dim msg As String = ex.Message.ToLowerInvariant()

        If msg.Contains("network-related") OrElse msg.Contains("server was not found") OrElse
           msg.Contains("could not open a connection") OrElse msg.Contains("network path was not found") Then
            title = "Could Not Reach Server"
            whatHappened = "The computer """ & server & """ could not be found on the network, or SQL Server is not running on it."
            whatToDo = "Things to check:" & vbCrLf & vbCrLf &
                       "1. Make sure the server name is spelled correctly" & vbCrLf &
                       "2. If SQL Server is on THIS computer, try:  .\SQLEXPRESS" & vbCrLf &
                       "3. If it's on another computer, use:  COMPUTERNAME\SQLEXPRESS" & vbCrLf &
                       "4. Make sure SQL Server is running on that computer" & vbCrLf &
                       "5. Check that this computer can reach the server (same network)"

        ElseIf msg.Contains("login failed") OrElse msg.Contains("not trusted") OrElse msg.Contains("password") Then
            title = "Login Was Refused"
            whatHappened = "Connected to the server, but the database login was not accepted."
            whatToDo = "This usually means the VastOffice database hasn't been set up with the expected login yet." & vbCrLf & vbCrLf &
                       "Contact your administrator and let them know the 'vastnr' database login needs to be configured on this SQL Server instance."

        ElseIf msg.Contains("cannot open database") OrElse msg.Contains("does not exist") Then
            title = "Database Not Found"
            whatHappened = "Connected to """ & server & """, but the VastOffice database does not exist there."
            whatToDo = "This server doesn't have the VastOffice database installed." & vbCrLf & vbCrLf &
                       "If this is a workstation, you'll need to point to your main server instead." & vbCrLf & vbCrLf &
                       "Try entering the name of your VAST server (e.g. VASTSERVER\SQLEXPRESS)."

        ElseIf msg.Contains("timeout") OrElse msg.Contains("timed out") Then
            title = "Connection Timed Out"
            whatHappened = "Waited too long trying to reach """ & server & """."
            whatToDo = "The server might be busy, turned off, or blocked by a firewall." & vbCrLf & vbCrLf &
                       "Things to check:" & vbCrLf & vbCrLf &
                       "1. Is the server computer turned on?" & vbCrLf &
                       "2. Is SQL Server running on it?" & vbCrLf &
                       "3. Is a firewall blocking port 1433?"

        Else
            title = "Connection Problem"
            whatHappened = "Something unexpected went wrong while trying to connect."
            whatToDo = "Please check the server name and try again. If the problem continues, " &
                       "contact your administrator with this information:" & vbCrLf & vbCrLf &
                       ex.Message
        End If

        ShowConnectionHelp(title, whatHappened, whatToDo)
    End Sub

    ''' <summary>
    ''' Show a branded help/error dialog with a title, explanation, and guidance.
    ''' Dialog auto-sizes vertically to fit all content regardless of text length.
    ''' </summary>
    Private Sub ShowConnectionHelp(title As String, whatHappened As String, whatToDo As String)
        Const CONTENT_WIDTH As Integer = 440
        Const MARGIN As Integer = 24
        Const PADDING As Integer = 12

        Using dlg As New Form()
            dlg.Text = title
            dlg.FormBorderStyle = FormBorderStyle.FixedDialog
            dlg.MaximizeBox = False
            dlg.MinimizeBox = False
            dlg.StartPosition = FormStartPosition.CenterParent
            dlg.BackColor = Color.White
            dlg.Font = New Font("Segoe UI", 9.0!)

            ' Magenta accent bar
            Dim pnlAccent As New Panel() With {
                .BackColor = UiTheme.Magenta,
                .Dock = DockStyle.Top,
                .Height = 4
            }

            ' Title
            Dim lblTitle As New Label() With {
                .Text = title,
                .Font = UiTheme.Semibold(14.0F),
                .ForeColor = UiTheme.Magenta,
                .MaximumSize = New Size(CONTENT_WIDTH, 0),
                .AutoSize = True
            }

            ' What happened
            Dim lblWhat As New Label() With {
                .Text = whatHappened,
                .Font = New Font("Segoe UI", 10.0!),
                .ForeColor = Color.FromArgb(80, 80, 80),
                .MaximumSize = New Size(CONTENT_WIDTH, 0),
                .AutoSize = True
            }

            ' Separator
            Dim pnlSep As New Panel() With {
                .BackColor = Color.FromArgb(230, 230, 230),
                .Size = New Size(CONTENT_WIDTH, 1)
            }

            ' How to fix label
            Dim lblHowLabel As New Label() With {
                .Text = "How to fix this:",
                .Font = UiTheme.Semibold(9.5F),
                .ForeColor = UiTheme.Charcoal,
                .AutoSize = True
            }

            ' How to fix content
            Dim lblHow As New Label() With {
                .Text = whatToDo,
                .Font = New Font("Segoe UI", 9.5!),
                .ForeColor = Color.FromArgb(60, 60, 60),
                .MaximumSize = New Size(CONTENT_WIDTH, 0),
                .AutoSize = True
            }

            ' Got it button
            Dim btnOk As New RoundedButton() With {
                .Text = "Got it",
                .Size = New Size(100, 38),
                .Font = UiTheme.Semibold(9.0F),
                .DialogResult = DialogResult.OK
            }

            ' Add controls so they measure correctly
            dlg.Controls.AddRange({pnlAccent, lblTitle, lblWhat, pnlSep, lblHowLabel, lblHow, btnOk})

            ' Stack elements vertically with padding - dialog sizes to fit
            dlg.SuspendLayout()

            Dim y As Integer = MARGIN
            lblTitle.Location = New Point(MARGIN, y)
            y += lblTitle.Height + PADDING

            lblWhat.Location = New Point(MARGIN, y)
            y += lblWhat.Height + PADDING

            pnlSep.Location = New Point(MARGIN, y)
            y += 1 + PADDING

            lblHowLabel.Location = New Point(MARGIN, y)
            y += lblHowLabel.Height + 6

            lblHow.Location = New Point(MARGIN, y)
            y += lblHow.Height + PADDING + 8

            btnOk.Location = New Point(MARGIN + CONTENT_WIDTH - btnOk.Width, y)
            y += btnOk.Height + MARGIN

            ' Size the dialog to fit all content
            dlg.ClientSize = New Size(MARGIN * 2 + CONTENT_WIDTH, y)
            dlg.AcceptButton = btnOk
            dlg.ResumeLayout(True)

            dlg.ShowDialog(Me)
        End Using
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As EventArgs)
        Try
            Dim config As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)

            SetOrAddSetting(config, "DatabaseServer", txtDatabaseServer.Text.Trim())
            SetOrAddSetting(config, "CustomerName", txtCustomerName.Text.Trim())
            SetOrAddSetting(config, "SiteName", txtSiteNumber.Text.Trim())

            config.Save(ConfigurationSaveMode.Modified)
            ConfigurationManager.RefreshSection("appSettings")

            Logger.Log("Settings saved successfully", Logger.LogLevel.Info)
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Catch ex As Exception
            Logger.Log("Error saving settings: " & ex.Message, Logger.LogLevel.Error)
            MessageBox.Show("Failed to save settings: " & ex.Message, "Error",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    ''' <summary>
    ''' Sets an appSettings key, creating it if it doesn't exist.
    ''' </summary>
    Private Shared Sub SetOrAddSetting(config As Configuration, key As String, value As String)
        If config.AppSettings.Settings(key) IsNot Nothing Then
            config.AppSettings.Settings(key).Value = value
        Else
            config.AppSettings.Settings.Add(key, value)
        End If
    End Sub

End Class
