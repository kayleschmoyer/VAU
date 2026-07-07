Option Strict On

Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Windows.Forms

''' <summary>
''' Settings dialog for configuring email recipients, customer/site identity,
''' and the SQL Server instance used for VastOffice database lookups.
''' Reads and writes directly to the application's .exe.config file.
''' </summary>
Public Class SettingsForm
    Inherits Form

    Private txtEmailTo As TextBox
    Private txtCustomerName As TextBox
    Private txtSiteName As TextBox
    Private txtDatabaseServer As TextBox
    Private btnTestConnection As Button
    Private btnSave As Button
    Private btnCancel As Button
    Private lblEmailTo As Label
    Private lblCustomerName As Label
    Private lblSiteName As Label
    Private lblDatabaseServer As Label
    Private lblDbHint As Label

    Private Shared ReadOnly Magenta As Color = Color.FromArgb(237, 1, 127)
    Private Shared ReadOnly Charcoal As Color = Color.FromArgb(51, 51, 51)

    Public Sub New()
        InitializeSettingsForm()
        LoadSettings()
    End Sub

    Private Sub InitializeSettingsForm()
        Me.Text = "Settings"
        Me.Size = New Size(420, 420)
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
            .ForeColor = Charcoal
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
            .ForeColor = Charcoal
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
            .ForeColor = Charcoal
        }

        txtSiteName = New TextBox() With {
            .Location = New Point(20, 162),
            .Size = New Size(360, 26),
            .Font = New Font("Segoe UI", 11.0!),
            .BorderStyle = BorderStyle.FixedSingle
        }

        lblDatabaseServer = New Label() With {
            .Text = "DATABASE SERVER",
            .Location = New Point(20, 200),
            .Size = New Size(360, 18),
            .Font = New Font("Segoe UI", 9.0!, FontStyle.Bold),
            .ForeColor = Charcoal
        }

        txtDatabaseServer = New TextBox() With {
            .Location = New Point(20, 222),
            .Size = New Size(260, 26),
            .Font = New Font("Segoe UI", 11.0!),
            .BorderStyle = BorderStyle.FixedSingle
        }

        btnTestConnection = New Button() With {
            .Text = "TEST",
            .Location = New Point(290, 222),
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
            .Location = New Point(20, 252),
            .Size = New Size(360, 16),
            .Font = New Font("Segoe UI", 8.0!, FontStyle.Italic),
            .ForeColor = Color.FromArgb(140, 140, 140)
        }

        btnSave = New Button() With {
            .Text = "SAVE",
            .Location = New Point(210, 310),
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
            .Location = New Point(300, 310),
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

        Me.Controls.AddRange({lblEmailTo, txtEmailTo, lblCustomerName, txtCustomerName,
                              lblSiteName, txtSiteName, lblDatabaseServer, txtDatabaseServer,
                              btnTestConnection, lblDbHint, btnSave, btnCancel})
        Me.AcceptButton = btnSave
        Me.CancelButton = btnCancel
    End Sub

    Private Sub LoadSettings()
        Try
            Dim emailTo As String = ConfigurationManager.AppSettings("EmailTo")
            If Not String.IsNullOrEmpty(emailTo) Then
                txtEmailTo.Text = emailTo
            End If

            txtCustomerName.Text = ConfigurationManager.AppSettings("CustomerName")
            txtSiteName.Text