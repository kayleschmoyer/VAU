Option Strict On

Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Drawing.Drawing2D
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
            Logger.Log($"Error loading settings: {ex.Message}", Logger.LogLevel.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Test the SQL connection and auto-fill Customer Name and Site Number from the query result.
    ''' </summary>
    Private Sub BtnTestConnection_Click(sender As Object