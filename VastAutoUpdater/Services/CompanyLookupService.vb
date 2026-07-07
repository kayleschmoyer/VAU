Option Strict On

Imports System.Data.SqlClient

''' <summary>
''' Queries the VastOffice database to retrieve company information (read-only).
''' Used to dynamically populate CustomerName and SiteName for dashboard reporting.
''' The database server instance is configurable via the DatabaseServer App.config key.
''' Credentials are always vastnr/snowdrift (SQL authentication).
''' </summary>
Public Class CompanyLookupService

    ''' <summary>
    ''' Build a read-only connection string for the configured database server.
    ''' Uses SQL authentication with the standard vastnr account.
    ''' ApplicationIntent=ReadOnly signals read-only intent to the server.
    ''' </summary>
    Private Shared Function BuildConnectionString() As String
        Dim server As String = ConfigManager.DatabaseServer
        If String.IsNullOrWhiteSpace(server) Then
            Logger.Log("DatabaseServer not configured - skipping company lookup", Logger.LogLevel.Warning)
            Return String.Empty
        End If
        Return $"Server={server};Database=VastOffice;User Id=vastnr;Password=snowdrift;Connection Timeout=5;ApplicationIntent=ReadOnly;"
    End Function

    ''' <summary>
    ''' Look up the company name and number from the COMPANY table.
    ''' Returns a tuple of (Name, CompanyNumber). Both empty if lookup fails or server not configured.
    ''' </summary>
    Public Shared Function GetCompanyInfo() As (Name As String, CompanyNumber As String)
        Try
            Dim connStr As String = BuildConnectionString()
            If String.IsNullOrEmpty(connStr) Then Return (String.Empty, String.Empty)

            Using conn As New SqlConnection(connStr)
                conn.Open()
             