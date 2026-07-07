Option Strict On

Imports System.Configuration
Imports System.Windows.Forms

''' <summary>
''' Simple settings dialog for configuring email recipients and the
''' customer/site identity reported to the VAU-Dashboard.
''' Reads and writes directly to the application's .exe.config file.
''' </summary>
Public Class SettingsForm
    Inherits Form

    Private txtEmailTo As TextBox
    Private txtCustomerName As TextBox
    Private txtSiteName As TextBox
    Private btnSave As Button
    Private btnCancel As Button
    Private lblEmailTo As Label
    Private lblCustomerName As Label
    Private lblSiteName As Label

    Private Shared ReadOnly Magenta As Color = Color.FromArgb(237, 1, 127)

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

        lblEmailTo = New Label() With {
            .Text = "EMAIL RECIPIENTS (comma-separated)",
            .Location = New Point(20, 20),
            .Size = New Size(360, 18),
            .Font = New Font("Segoe UI", 9.0!, FontStyle.Bold),
            .ForeColor = Color.FromArgb(51, 51, 51)
        }

        txtEmailTo = New TextBox() With {
            .Location = New Point(20, 42),
            .Size = New Size(360, 26),
            .Font = New Font("Segoe UI", 11.0!),
            .BorderStyle = BorderStyle.FixedSingle
        }

        lblCustomerName = New Label() With {
            .Text = "CUSTOMER NAME",
            .Location = New Point(20, 80),
            .Size = New Size(360, 18),
            .Font = New Font("Segoe UI", 9.0!, FontStyle.Bold),
            .ForeColor = Color.FromArgb(51, 51, 51)
        }

        txtCustomerName = New TextBox() With {
            .Location = New Point(20, 102),
            .Size = New Size(360, 26),
            .Font = New Font("Segoe UI", 11.0!),
            .BorderStyle = BorderStyle.FixedSingle
        }

        lblSiteName = New Label() With {
            .Text = "SITE NAME",
            .Location = New Point(20, 140),
            .Size = New Size(360, 18),
            .Font = New Font("Segoe UI", 9.0!, FontStyle.Bold),
            .ForeColor = Color.FromArgb(51, 51, 51)
        }

        txtSiteName = New TextBox() With {
            .Location = New Point(20, 162),
            .Size = New Size(360, 26),
            .Font = New Font("Segoe UI", 11.0!),
            .BorderStyle = BorderStyle.FixedSingle
        }

        btnSave = New Button() With {
            .Text = "SAVE",
            .Location = New Point(210, 220),
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
            .Location = New Point(300, 220),
            .Size = New Size(80, 36),
            .FlatStyle = FlatStyle.Flat,
            .BackColor = Color.FromArgb(200, 200, 200),
            .ForeColor = Color.FromArgb(51, 51, 51),
            .Font = New Font("Segoe UI", 9.0!, FontStyle.Bold),
            .Cursor = Cursors.Hand
        }
        btnCancel.FlatAppearance.BorderSize = 0

        AddHandler btnSave.Click, AddressOf BtnSave_Click
        AddHandler btnCancel.Click, AddressOf BtnCancel_Click

        Me.Controls.AddRange({lblEmailTo, txtEmailTo, lblCustomerName, txtCustomerName,
                              lblSiteName, txtSiteName, btnSave, btnCancel})
        Me.AcceptButton = btnSave
        Me.CancelButton = btnCancel
    End Sub

    Private Sub LoadSettings()
        Try
            Dim emailTo As String = ConfigurationManager.AppSettings("EmailTo")
            If Not String.IsNullOrEmpty(emailTo) Then
                txtEmailTo.Text = emailTo
            End If

            txtCustomerName.Text = ConfigManager.CustomerName
            txtSiteName.Text = ConfigManager.SiteName
        Catch ex As Exception
            Logger.Log($"Error loading settings: {ex.Message}", Logger.LogLevel.Error)
        End Try
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As EventArgs)
        Try
            Dim config As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)

            SetOrAddSetting(config, "EmailTo", txtEmailTo.Text.Trim())
            SetOrAddSetting(config, "CustomerName", txtCustomerName.Text.Trim())
            SetOrAddSetting(config, "SiteName", txtSiteName.Text.Trim())

            config.Save(ConfigurationSaveMode.Modified)
            ConfigurationManager.RefreshSection("appSettings")

            Logger.Log("Settings saved successfully", Logger.LogLevel.Info)
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Catch ex As Exception
            Logger.Log($"Error saving settings: {ex.Message}", Logger.LogLevel.Error)
            MessageBox.Show($"Failed to save settings: {ex.Message}", "Error",
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
