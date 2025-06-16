<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class VASTUpdater
    Inherits MaterialSkin.Controls.MaterialForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Private components As System.ComponentModel.IContainer
    Private WithEvents lblStatus As MaterialSkin.Controls.MaterialLabel
    Private WithEvents lblCurrentVersion As MaterialSkin.Controls.MaterialLabel

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.lblStatus = New MaterialSkin.Controls.MaterialLabel()
        Me.lblCurrentVersion = New MaterialSkin.Controls.MaterialLabel()

        ' 
        ' lblStatus
        ' 
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Depth = 0
        Me.lblStatus.Font = New System.Drawing.Font("Roboto", 11.0!)
        Me.lblStatus.ForeColor = System.Drawing.Color.Black
        Me.lblStatus.Location = New System.Drawing.Point(30, 300)
        Me.lblStatus.MouseState = MaterialSkin.MouseState.HOVER
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(240, 19)
        Me.lblStatus.TabIndex = 0
        Me.lblStatus.Text = "Status: Ready for update check..."

        ' 
        ' lblCurrentVersion
        ' 
        Me.lblCurrentVersion.AutoSize = True
        Me.lblCurrentVersion.Depth = 0
        Me.lblCurrentVersion.Font = New System.Drawing.Font("Roboto", 11.0!)
        Me.lblCurrentVersion.ForeColor = System.Drawing.Color.Black
        Me.lblCurrentVersion.Location = New System.Drawing.Point(30, 330)
        Me.lblCurrentVersion.MouseState = MaterialSkin.MouseState.HOVER
        Me.lblCurrentVersion.Name = "lblCurrentVersion"
        Me.lblCurrentVersion.Size = New System.Drawing.Size(160, 19)
        Me.lblCurrentVersion.TabIndex = 1
        Me.lblCurrentVersion.Text = "Current Version: ..."

        ' 
        ' VASTUpdater
        ' 
        Me.ClientSize = New System.Drawing.Size(600, 400)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.lblCurrentVersion)
        Me.Name = "VASTUpdater"
        Me.Text = "VAST Updater"
    End Sub

End Class
