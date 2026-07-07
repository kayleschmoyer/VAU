Option Strict On

Imports System.Data.SqlClient

''' <summary>
''' Queries the local VastOffice database to retrieve company information.
''' Used to dynamically populate CustomerName and SiteName for dashboard reporting.
''' </summary>
Public Class CompanyLookupService

    Private Const CONNECTION_STRING As String = "Server=.\SQLEXPRESS;Database=VastOffice;Integrated Security=True;Connection Timeout=5;"

    ''' <summary>
    ''' Look up the company name and number from the COMPANY table.
    ''' Returns a tuple of (Name, CompanyNumber). Both empty if lookup fails.
    ''' </summary>
    Public Shared Function GetCompanyInfo() As (Name As String, CompanyNumber As String)
        Try
            Using conn As New SqlConnection(CONNECTION_STRING)
                conn.Open()
                Using cmd As New SqlCommand("SELECT TOP 1 NAME, COMPANY_NUMBER FROM COMPANY", conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Dim name As String = If(reader.IsDBNull(0), String.Empty, reader.GetString(0).Trim())
                            Dim companyNumber As String = If(reader.IsDBNull(1), String.Empty, reader(1).ToString().Trim())
                            Logger.Log($"Company lookup: {name} (#{companyNumber})", Logger.LogLevel.Info)
                            Return (name, companyNumber)
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Logger.Log($"Company lookup failed: {ex.Message}", Logger.LogLevel.Warning)
        End Try

        Return (String.Empty, String.Empty)
    End Function

End Class
