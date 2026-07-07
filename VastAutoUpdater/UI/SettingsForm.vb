Option Strict On

Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Windows.Forms

''' <summary>
''' Settings dialog for configuring the SQL Server instance, customer identity,
''' and site number. The TEST button queries VastOffice and auto-fills
''' Customer Name and Site Number from the COMPANY table.
''' </summary>
Public Class SettingsForm
    Inherits Form

    Private txtDatabaseServer As TextBox
    Private txtCustomerName As TextBox
    Private txtSiteNumber As TextBox
    Private btnTestConnection As Button
    Private btnSave As Button
    Private btnCancel As Button
    Private lblDatabaseServer As Label
    Private lblDbHint As Label
    Private lblCustomerName As Label
    Private lblSiteNumber As Label

    Private Shared ReadOnly Magenta As Color = Color.FromArgb(237, 1, 127)
    Private Shared ReadOnly Charcoal As Color = Color.FromArgb(51, 51, 51)

    Public Sub New()
        InitializeSettingsForm()
        LoadSettings()
    End Sub

    Private Sub InitializeSettingsForm()
        Me.Text = "Settings"
        Me.Size = New Size(420, 320)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.White
        Me.Font = New Font("Segoe UI", 9.0!)

        ' --- Database Server (top) ---
        lblDatabaseServer = New Label() With {
            .Text = "DATABASE SERVER",
            .Location = New Point(20, 20),
            .Size = New Size(360, 18),
            .Font = New Font("Segoe UI", 9.0!, FontStyle.Bold),
            .ForeColor = Charcoal
        }

        txtDatabaseServer = New TextBox() With {
            .Location = New Point(20, 42),
            .Size = New Size(260, 26),
            .Font = New Font("Segoe UI", 11.0!),
            .BorderStyle = BorderStyle.FixedSingle
        }

        btnTestConnection = New Button() With {
            .Text = "TEST",
            .Location = New Point(290, 42),
            .Size = New Size(90, 26),
            .FlatStyle = FlatStyle.Flat,
            .BackColor = Color.FromArgb(0, 150, 80),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 8.0!, FontStyle.Bold),
            .Cursor = Cursors.Hand
        }
        btnTestConnection.FlatAppearance.BorderSize = 0

        lblDbHint = New Label() With {
            .Text = "e.g. .\SQLEXPRESS  or  SERVERNAME\SQLEXPRESS  or  SERVERNAME",
            .Location = New Point(20, 72),
            .Size = New Size(360, 16),
            .Font = New Font("Segoe UI", 8.0!, FontStyle.Italic),
            .ForeColor = Color.FromArgb(140, 140, 140)
        }

        ' --- Customer Name ---
        lblCustomerName = New Label() With {
            .Text = "CUSTOMER NAME",
            .Location = New Point(20, 100),
            .Size = New Size(360, 18),
            .Font = New Font("Segoe UI", 9.0!, FontStyle.Bold),
            .ForeColor = Charcoal
        }

        txtCustomerName = New TextBox() With {
            .Location = New Point(20, 122),
            .Size = New Size(360, 26),
            .Font = New Font("Segoe UI", 11.0!),
            .BorderStyle = BorderStyle.FixedSingle
        }

        ' --- Site Number ---
        lblSiteNumber = New Label() With {
            .Text = "SITE NUMBER",
            .Location = New Point(20, 160),
            .Size = New Size(360, 18),
            .Font = New Font("Segoe UI", 9.0!, FontStyle.Bold),
            .ForeColor = Charcoal
        }

        txtSiteNumber = New TextBox() With {
            .Location = New Point(20, 182),
            .Size = New Size(360, 26),
            .Font = New Font("Segoe UI", 11.0!),
            .BorderStyle = BorderStyle.FixedSingle
        }

        ' --- Buttons ---
        btnSave = New Button() With {
            .Text = "SAVE",
            .Location = New Point(210, 230),
            .Size = New Size(80, 36),
            .FlatStyle = FlatStyle.Flat,
            .BackColor = Magenta,
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 9.0!, FontStyle.Bold),
            .Cursor = Cursors.Hand
        }
        btnSave.FlatAppearance.BorderSize = 0

        btnCancel = New Button() With {
            .Text = "CANCEL",
            .Location = New Point(300, 230),
            .Size = New Size(80, 36),
            .FlatStyle = FlatStyle.Flat,
            .BackColor = Color.FromArgb(200, 200, 200),
            .ForeColor = Charcoal,
            .Font = New Font("Segoe UI", 9.0!, FontStyle.Bold),
            .Cursor = Cursors.Hand
        }
        btnCancel.FlatAppearance.BorderSize = 0

        AddHandler btnTestConnection.Click, AddressOf BtnTestConnection_Click
        AddHandler btnSave.Click, AddressOf BtnSave_Click
        AddHandler btnCancel.Click, AddressOf BtnCancel_Click

        Me.Controls.AddRange({lblDatabaseServer, txtDatabaseServer, btnTestConnection, lblDbHint,
                              lblCustomerName, txtCustomerName,
                              lblSiteNumber, txtSiteNumber,
                              btnSave, btnCancel})
        Me.AcceptButton = btnSave
        Me.CancelButton = btnCancel
    End Sub

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
            btnTestConnection.Text = "TEST"
        End Try
    End Sub

    ''' <summary>
    ''' Show a branded success dialog with the retrieved company info.
    ''' </summary>
    Private Sub ShowConnectionSuccess(compName As String, compNum As String)
        Const CONTENT_WIDTH As Integer = 380
        Const MARGIN As Integer = 24
        Const PADDING As Integer = 12
        Dim SuccessGreen As Color = Color.FromArgb(0, 150, 80)

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
                .BackColor = SuccessGreen,
                .Dock = DockStyle.Top,
                .Height = 4
            }

            Dim lblTitle As New Label() With {
                .Text = "Connection Successful",
                .Font = New Font("Segoe UI", 14.0!, FontStyle.Bold),
                .ForeColor = SuccessGreen,
                .MaximumSize = New Size(CONTENT_WIDTH, 0),
                .AutoSize = True
            }

            Dim lblDetails As New Label() With {
                .Text = "Customer Name:  " & compName & vbCrLf & "Site Number:  " & compNum,
                .Font = New Font("Segoe UI", 11.0!),
                .ForeColor = Charcoal,
                .MaximumSize = New Size(CONTENT_WIDTH, 0),
                .AutoSize = True
            }

            Dim lblNote As New Label() With {
                .Text = "These values have been filled in below. Click SAVE to keep them.",
                .Font = New Font("Segoe UI", 9.0!, FontStyle.Italic),
                .ForeColor = Color.FromArgb(100, 100, 100),
                .MaximumSize = New Size(CONTENT_WIDTH, 0),
                .AutoSize = True
            }

            Dim btnOk As New Button() With {
                .Text = "OK",
                .Size = New Size(80, 34),
                .FlatStyle = FlatStyle.Flat,
                .BackColor = SuccessGreen,
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI", 9.0!, FontStyle.Bold),
                .Cursor = Cursors.Hand,
                .DialogResult = DialogResult.OK
            }
            btnOk.FlatAppearance.BorderSize = 0

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
                .BackColor = Magenta,
                .Dock = DockStyle.Top,
                .Height = 4
            }

            ' Title
            Dim lblTitle As New Label() With {
                .Text = title,
                .Font = New Font("Segoe UI", 14.0!, FontStyle.Bold),
                .ForeColor = Magenta,
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
                .Font = New Font("Segoe UI", 9.5!, FontStyle.Bold),
                .ForeColor = Charcoal,
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

            ' GOT IT button
            Dim btnOk As New Button() With {
                .Text = "GOT IT",
                .Size = New Size(100, 36),
                .FlatStyle = FlatStyle.Flat,
                .BackColor = Magenta,
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI", 9.0!, FontStyle.Bold),
                .Cursor = Cursors.Hand,
                .DialogResult = DialogResult.OK
            }
            btnOk.FlatAppearance.BorderSize = 0

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
