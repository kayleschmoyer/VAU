<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainForm
    Inherits System.Windows.Forms.Form

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

    ' Header
    Private WithEvents pnlHeader As System.Windows.Forms.Panel
    Private WithEvents lblTitle As System.Windows.Forms.Label
    Private WithEvents lblSubtitle As System.Windows.Forms.Label
    Private WithEvents btnClose As System.Windows.Forms.Button
    Private WithEvents btnMinimize As System.Windows.Forms.Button
    Private WithEvents btnSettings As System.Windows.Forms.Button

    ' Credential inputs
    Private WithEvents pnlCredentials As System.Windows.Forms.Panel
    Private WithEvents lblUsername As System.Windows.Forms.Label
    Private WithEvents txtSftpUsername As System.Windows.Forms.TextBox
    Private WithEvents lblPassword As System.Windows.Forms.Label
    Private WithEvents txtSftpPassword As System.Windows.Forms.TextBox

    ' Action
    Private WithEvents btnCheckForUpdates As System.Windows.Forms.Button

    ' Progress
    Private WithEvents pnlProgress As System.Windows.Forms.Panel
    Private WithEvents progressBar1 As System.Windows.Forms.ProgressBar

    ' Status
    Private WithEvents lblStatus As System.Windows.Forms.Label
    Private WithEvents lblCurrentVersion As System.Windows.Forms.Label

    ' Footer accent
    Private WithEvents pnlFooter As System.Windows.Forms.Panel

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.pnlHeader = New System.Windows.Forms.Panel()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.lblSubtitle = New System.Windows.Forms.Label()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnMinimize = New System.Windows.Forms.Button()
        Me.btnSettings = New System.Windows.Forms.Button()
        Me.pnlCredentials = New System.Windows.Forms.Panel()
        Me.lblUsername = New System.Windows.Forms.Label()
        Me.txtSftpUsername = New System.Windows.Forms.TextBox()
        Me.lblPassword = New System.Windows.Forms.Label()
        Me.txtSftpPassword = New System.Windows.Forms.TextBox()
        Me.btnCheckForUpdates = New System.Windows.Forms.Button()
        Me.pnlProgress = New System.Windows.Forms.Panel()
        Me.progressBar1 = New System.Windows.Forms.ProgressBar()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.lblCurrentVersion = New System.Windows.Forms.Label()
        Me.pnlFooter = New System.Windows.Forms.Panel()
        Me.pnlHeader.SuspendLayout()
        Me.pnlCredentials.SuspendLayout()
        Me.pnlProgress.SuspendLayout()
        Me.SuspendLayout()
        '
        ' pnlHeader — gradient painted in code-behind
        '
        Me.pnlHeader.BackColor = System.Drawing.Color.FromArgb(237, 1, 127)
        Me.pnlHeader.Controls.Add(Me.btnClose)
        Me.pnlHeader.Controls.Add(Me.btnMinimize)
        Me.pnlHeader.Controls.Add(Me.btnSettings)
        Me.pnlHeader.Controls.Add(Me.lblTitle)
        Me.pnlHeader.Controls.Add(Me.lblSubtitle)
        Me.pnlHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlHeader.Name = "pnlHeader"
        Me.pnlHeader.Size = New System.Drawing.Size(520, 55)
        Me.pnlHeader.TabIndex = 0
        '
        ' lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.BackColor = System.Drawing.Color.Transparent
        Me.lblTitle.Font = New System.Drawing.Font("Segoe UI", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.lblTitle.ForeColor = System.Drawing.Color.White
        Me.lblTitle.Location = New System.Drawing.Point(16, 8)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(200, 24)
        Me.lblTitle.TabIndex = 0
        Me.lblTitle.Text = "VAST Updater"
        '
        ' lblSubtitle
        '
        Me.lblSubtitle.AutoSize = True
        Me.lblSubtitle.BackColor = System.Drawing.Color.Transparent
        Me.lblSubtitle.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.lblSubtitle.ForeColor = System.Drawing.Color.FromArgb(255, 200, 230)
        Me.lblSubtitle.Location = New System.Drawing.Point(18, 32)
        Me.lblSubtitle.Name = "lblSubtitle"
        Me.lblSubtitle.Size = New System.Drawing.Size(180, 16)
        Me.lblSubtitle.TabIndex = 1
        Me.lblSubtitle.Text = "Automated patch management"
        '
        ' btnClose
        '
        Me.btnClose.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.btnClose.BackColor = System.Drawing.Color.Transparent
        Me.btnClose.FlatAppearance.BorderSize = 0
        Me.btnClose.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(200, 0, 100)
        Me.btnClose.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(220, 0, 110)
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnClose.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.btnClose.ForeColor = System.Drawing.Color.White
        Me.btnClose.Location = New System.Drawing.Point(484, 4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(32, 28)
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "X"
        Me.btnClose.UseVisualStyleBackColor = False
        '
        ' btnMinimize
        '
        Me.btnMinimize.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.btnMinimize.BackColor = System.Drawing.Color.Transparent
        Me.btnMinimize.FlatAppearance.BorderSize = 0
        Me.btnMinimize.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(200, 0, 100)
        Me.btnMinimize.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(220, 0, 110)
        Me.btnMinimize.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnMinimize.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.btnMinimize.ForeColor = System.Drawing.Color.White
        Me.btnMinimize.Location = New System.Drawing.Point(450, 4)
        Me.btnMinimize.Name = "btnMinimize"
        Me.btnMinimize.Size = New System.Drawing.Size(32, 28)
        Me.btnMinimize.TabIndex = 3
        Me.btnMinimize.Text = "_"
        Me.btnMinimize.UseVisualStyleBackColor = False
        '
        ' btnSettings
        '
        Me.btnSettings.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.btnSettings.BackColor = System.Drawing.Color.Transparent
        Me.btnSettings.FlatAppearance.BorderSize = 0
        Me.btnSettings.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(200, 0, 100)
        Me.btnSettings.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(220, 0, 110)
        Me.btnSettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSettings.Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Regular)
        Me.btnSettings.ForeColor = System.Drawing.Color.White
        Me.btnSettings.Location = New System.Drawing.Point(416, 4)
        Me.btnSettings.Name = "btnSettings"
        Me.btnSettings.Size = New System.Drawing.Size(32, 28)
        Me.btnSettings.TabIndex = 7
        Me.btnSettings.Text = ChrW(&H2699)
        Me.btnSettings.UseVisualStyleBackColor = False
        '
        ' pnlCredentials
        '
        Me.pnlCredentials.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.pnlCredentials.BackColor = System.Drawing.Color.White
        Me.pnlCredentials.Controls.Add(Me.lblUsername)
        Me.pnlCredentials.Controls.Add(Me.txtSftpUsername)
        Me.pnlCredentials.Controls.Add(Me.lblPassword)
        Me.pnlCredentials.Controls.Add(Me.txtSftpPassword)
        Me.pnlCredentials.Location = New System.Drawing.Point(24, 70)
        Me.pnlCredentials.Name = "pnlCredentials"
        Me.pnlCredentials.Padding = New System.Windows.Forms.Padding(16, 12, 16, 12)
        Me.pnlCredentials.Size = New System.Drawing.Size(472, 130)
        Me.pnlCredentials.TabIndex = 1
        '
        ' lblUsername
        '
        Me.lblUsername.AutoSize = True
        Me.lblUsername.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblUsername.ForeColor = System.Drawing.Color.FromArgb(51, 51, 51)
        Me.lblUsername.Location = New System.Drawing.Point(16, 12)
        Me.lblUsername.Name = "lblUsername"
        Me.lblUsername.Size = New System.Drawing.Size(100, 16)
        Me.lblUsername.TabIndex = 0
        Me.lblUsername.Text = "SFTP USERNAME"
        '
        ' txtSftpUsername
        '
        Me.txtSftpUsername.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.txtSftpUsername.BackColor = System.Drawing.Color.White
        Me.txtSftpUsername.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSftpUsername.Font = New System.Drawing.Font("Segoe UI", 11.0!)
        Me.txtSftpUsername.ForeColor = System.Drawing.Color.FromArgb(51, 51, 51)
        Me.txtSftpUsername.Location = New System.Drawing.Point(16, 32)
        Me.txtSftpUsername.Name = "txtSftpUsername"
        Me.txtSftpUsername.Size = New System.Drawing.Size(440, 26)
        Me.txtSftpUsername.TabIndex = 1
        '
        ' lblPassword
        '
        Me.lblPassword.AutoSize = True
        Me.lblPassword.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblPassword.ForeColor = System.Drawing.Color.FromArgb(51, 51, 51)
        Me.lblPassword.Location = New System.Drawing.Point(16, 68)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(100, 16)
        Me.lblPassword.TabIndex = 2
        Me.lblPassword.Text = "SFTP PASSWORD"
        '
        ' txtSftpPassword
        '
        Me.txtSftpPassword.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.txtSftpPassword.BackColor = System.Drawing.Color.White
        Me.txtSftpPassword.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSftpPassword.Font = New System.Drawing.Font("Segoe UI", 11.0!)
        Me.txtSftpPassword.ForeColor = System.Drawing.Color.FromArgb(51, 51, 51)
        Me.txtSftpPassword.Location = New System.Drawing.Point(16, 88)
        Me.txtSftpPassword.Name = "txtSftpPassword"
        Me.txtSftpPassword.Size = New System.Drawing.Size(440, 26)
        Me.txtSftpPassword.TabIndex = 3
        Me.txtSftpPassword.UseSystemPasswordChar = True
        '
        ' btnCheckForUpdates
        '
        Me.btnCheckForUpdates.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.btnCheckForUpdates.BackColor = System.Drawing.Color.FromArgb(237, 1, 127)
        Me.btnCheckForUpdates.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnCheckForUpdates.FlatAppearance.BorderSize = 0
        Me.btnCheckForUpdates.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(190, 0, 100)
        Me.btnCheckForUpdates.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(210, 0, 115)
        Me.btnCheckForUpdates.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCheckForUpdates.Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Bold)
        Me.btnCheckForUpdates.ForeColor = System.Drawing.Color.White
        Me.btnCheckForUpdates.Location = New System.Drawing.Point(24, 214)
        Me.btnCheckForUpdates.Name = "btnCheckForUpdates"
        Me.btnCheckForUpdates.Size = New System.Drawing.Size(472, 44)
        Me.btnCheckForUpdates.TabIndex = 2
        Me.btnCheckForUpdates.Text = "CHECK FOR UPDATES"
        Me.btnCheckForUpdates.UseVisualStyleBackColor = False
        '
        ' pnlProgress
        '
        Me.pnlProgress.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.pnlProgress.BackColor = System.Drawing.Color.White
        Me.pnlProgress.Controls.Add(Me.progressBar1)
        Me.pnlProgress.Location = New System.Drawing.Point(24, 272)
        Me.pnlProgress.Name = "pnlProgress"
        Me.pnlProgress.Padding = New System.Windows.Forms.Padding(0, 6, 0, 6)
        Me.pnlProgress.Size = New System.Drawing.Size(472, 20)
        Me.pnlProgress.TabIndex = 3
        Me.pnlProgress.Visible = False
        '
        ' progressBar1
        '
        Me.progressBar1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.progressBar1.Location = New System.Drawing.Point(0, 6)
        Me.progressBar1.Name = "progressBar1"
        Me.progressBar1.Size = New System.Drawing.Size(472, 8)
        Me.progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.progressBar1.TabIndex = 0
        '
        ' lblStatus
        '
        Me.lblStatus.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.lblStatus.Font = New System.Drawing.Font("Segoe UI", 9.5!, System.Drawing.FontStyle.Regular)
        Me.lblStatus.ForeColor = System.Drawing.Color.FromArgb(51, 51, 51)
        Me.lblStatus.Location = New System.Drawing.Point(24, 302)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(472, 20)
        Me.lblStatus.TabIndex = 4
        Me.lblStatus.Text = "Ready for update check..."
        '
        ' lblCurrentVersion
        '
        Me.lblCurrentVersion.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.lblCurrentVersion.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular)
        Me.lblCurrentVersion.ForeColor = System.Drawing.Color.FromArgb(140, 140, 140)
        Me.lblCurrentVersion.Location = New System.Drawing.Point(24, 326)
        Me.lblCurrentVersion.Name = "lblCurrentVersion"
        Me.lblCurrentVersion.Size = New System.Drawing.Size(472, 18)
        Me.lblCurrentVersion.TabIndex = 5
        Me.lblCurrentVersion.Text = "Current Version: ..."
        '
        ' pnlFooter — thin magenta accent bar at the bottom
        '
        Me.pnlFooter.BackColor = System.Drawing.Color.FromArgb(237, 1, 127)
        Me.pnlFooter.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlFooter.Location = New System.Drawing.Point(0, 356)
        Me.pnlFooter.Name = "pnlFooter"
        Me.pnlFooter.Size = New System.Drawing.Size(520, 4)
        Me.pnlFooter.TabIndex = 6
        '
        ' MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(520, 360)
        Me.Controls.Add(Me.pnlFooter)
        Me.Controls.Add(Me.lblCurrentVersion)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.pnlProgress)
        Me.Controls.Add(Me.btnCheckForUpdates)
        Me.Controls.Add(Me.pnlCredentials)
        Me.Controls.Add(Me.pnlHeader)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.Name = "MainForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "VAST Updater"
        Me.pnlHeader.ResumeLayout(False)
        Me.pnlHeader.PerformLayout()
        Me.pnlCredentials.ResumeLayout(False)
        Me.pnlCredentials.PerformLayout()
        Me.pnlProgress.ResumeLayout(False)
        Me.ResumeLayout(False)
    End Sub

End Class
