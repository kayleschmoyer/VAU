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

    ' Credential card
    Private WithEvents cardCredentials As CardPanel
    Private WithEvents lblUsername As System.Windows.Forms.Label
    Private WithEvents txtSftpUsername As ModernTextBox
    Private WithEvents lblPassword As System.Windows.Forms.Label
    Private WithEvents txtSftpPassword As ModernTextBox

    ' Action
    Private WithEvents btnCheckForUpdates As RoundedButton

    ' Activity card (status + progress)
    Private WithEvents cardActivity As CardPanel
    Private WithEvents lblStatusIcon As System.Windows.Forms.Label
    Private WithEvents lblStatus As System.Windows.Forms.Label
    Private WithEvents lblPercent As System.Windows.Forms.Label
    Private WithEvents progressBar1 As SmoothProgressBar
    Private WithEvents lblProgressDetail As System.Windows.Forms.Label

    ' Footer row
    Private WithEvents lblCurrentVersion As System.Windows.Forms.Label
    Private WithEvents badgeUpdate As PillBadge
    Private WithEvents pnlFooter As System.Windows.Forms.Panel

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.pnlHeader = New System.Windows.Forms.Panel()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.lblSubtitle = New System.Windows.Forms.Label()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnMinimize = New System.Windows.Forms.Button()
        Me.btnSettings = New System.Windows.Forms.Button()
        Me.cardCredentials = New CardPanel()
        Me.lblUsername = New System.Windows.Forms.Label()
        Me.txtSftpUsername = New ModernTextBox()
        Me.lblPassword = New System.Windows.Forms.Label()
        Me.txtSftpPassword = New ModernTextBox()
        Me.btnCheckForUpdates = New RoundedButton()
        Me.cardActivity = New CardPanel()
        Me.lblStatusIcon = New System.Windows.Forms.Label()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.lblPercent = New System.Windows.Forms.Label()
        Me.progressBar1 = New SmoothProgressBar()
        Me.lblProgressDetail = New System.Windows.Forms.Label()
        Me.lblCurrentVersion = New System.Windows.Forms.Label()
        Me.badgeUpdate = New PillBadge()
        Me.pnlFooter = New System.Windows.Forms.Panel()
        Me.pnlHeader.SuspendLayout()
        Me.cardCredentials.SuspendLayout()
        Me.cardActivity.SuspendLayout()
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
        Me.pnlHeader.Size = New System.Drawing.Size(520, 64)
        Me.pnlHeader.TabIndex = 0
        '
        ' lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.BackColor = System.Drawing.Color.Transparent
        Me.lblTitle.Font = New System.Drawing.Font("Segoe UI", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.lblTitle.ForeColor = System.Drawing.Color.White
        Me.lblTitle.Location = New System.Drawing.Point(20, 10)
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
        Me.lblSubtitle.Location = New System.Drawing.Point(22, 38)
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
        Me.btnClose.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(150, 0, 80)
        Me.btnClose.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(180, 0, 96)
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnClose.Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Regular)
        Me.btnClose.ForeColor = System.Drawing.Color.White
        Me.btnClose.Location = New System.Drawing.Point(478, 15)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(32, 30)
        Me.btnClose.TabIndex = 2
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "X"
        Me.btnClose.UseVisualStyleBackColor = False
        '
        ' btnMinimize
        '
        Me.btnMinimize.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.btnMinimize.BackColor = System.Drawing.Color.Transparent
        Me.btnMinimize.FlatAppearance.BorderSize = 0
        Me.btnMinimize.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(150, 0, 80)
        Me.btnMinimize.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(180, 0, 96)
        Me.btnMinimize.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnMinimize.Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Regular)
        Me.btnMinimize.ForeColor = System.Drawing.Color.White
        Me.btnMinimize.Location = New System.Drawing.Point(444, 15)
        Me.btnMinimize.Name = "btnMinimize"
        Me.btnMinimize.Size = New System.Drawing.Size(32, 30)
        Me.btnMinimize.TabIndex = 3
        Me.btnMinimize.TabStop = False
        Me.btnMinimize.Text = "_"
        Me.btnMinimize.UseVisualStyleBackColor = False
        '
        ' btnSettings
        '
        Me.btnSettings.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.btnSettings.BackColor = System.Drawing.Color.Transparent
        Me.btnSettings.FlatAppearance.BorderSize = 0
        Me.btnSettings.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(150, 0, 80)
        Me.btnSettings.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(180, 0, 96)
        Me.btnSettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSettings.Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Regular)
        Me.btnSettings.ForeColor = System.Drawing.Color.White
        Me.btnSettings.Location = New System.Drawing.Point(410, 15)
        Me.btnSettings.Name = "btnSettings"
        Me.btnSettings.Size = New System.Drawing.Size(32, 30)
        Me.btnSettings.TabIndex = 7
        Me.btnSettings.TabStop = False
        Me.btnSettings.Text = Global.Microsoft.VisualBasic.ChrW(&H2699)
        Me.btnSettings.UseVisualStyleBackColor = False
        '
        ' cardCredentials
        '
        Me.cardCredentials.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.cardCredentials.BackColor = System.Drawing.Color.White
        Me.cardCredentials.Controls.Add(Me.lblUsername)
        Me.cardCredentials.Controls.Add(Me.txtSftpUsername)
        Me.cardCredentials.Controls.Add(Me.lblPassword)
        Me.cardCredentials.Controls.Add(Me.txtSftpPassword)
        Me.cardCredentials.Location = New System.Drawing.Point(24, 88)
        Me.cardCredentials.Name = "cardCredentials"
        Me.cardCredentials.Size = New System.Drawing.Size(472, 156)
        Me.cardCredentials.TabIndex = 1
        '
        ' lblUsername
        '
        Me.lblUsername.AutoSize = True
        Me.lblUsername.BackColor = System.Drawing.Color.White
        Me.lblUsername.Font = New System.Drawing.Font("Segoe UI", 8.0!, System.Drawing.FontStyle.Bold)
        Me.lblUsername.ForeColor = System.Drawing.Color.FromArgb(138, 138, 144)
        Me.lblUsername.Location = New System.Drawing.Point(18, 15)
        Me.lblUsername.Name = "lblUsername"
        Me.lblUsername.Size = New System.Drawing.Size(100, 14)
        Me.lblUsername.TabIndex = 0
        Me.lblUsername.Text = "SFTP USERNAME"
        '
        ' txtSftpUsername
        '
        Me.txtSftpUsername.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.txtSftpUsername.Location = New System.Drawing.Point(18, 35)
        Me.txtSftpUsername.Name = "txtSftpUsername"
        Me.txtSftpUsername.Size = New System.Drawing.Size(436, 36)
        Me.txtSftpUsername.TabIndex = 1
        '
        ' lblPassword
        '
        Me.lblPassword.AutoSize = True
        Me.lblPassword.BackColor = System.Drawing.Color.White
        Me.lblPassword.Font = New System.Drawing.Font("Segoe UI", 8.0!, System.Drawing.FontStyle.Bold)
        Me.lblPassword.ForeColor = System.Drawing.Color.FromArgb(138, 138, 144)
        Me.lblPassword.Location = New System.Drawing.Point(18, 83)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(100, 14)
        Me.lblPassword.TabIndex = 2
        Me.lblPassword.Text = "SFTP PASSWORD"
        '
        ' txtSftpPassword
        '
        Me.txtSftpPassword.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.txtSftpPassword.IsPassword = True
        Me.txtSftpPassword.Location = New System.Drawing.Point(18, 103)
        Me.txtSftpPassword.Name = "txtSftpPassword"
        Me.txtSftpPassword.Size = New System.Drawing.Size(436, 36)
        Me.txtSftpPassword.TabIndex = 3
        '
        ' btnCheckForUpdates
        '
        Me.btnCheckForUpdates.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.btnCheckForUpdates.Location = New System.Drawing.Point(24, 260)
        Me.btnCheckForUpdates.Name = "btnCheckForUpdates"
        Me.btnCheckForUpdates.Size = New System.Drawing.Size(472, 46)
        Me.btnCheckForUpdates.TabIndex = 2
        Me.btnCheckForUpdates.Text = "Check for updates"
        '
        ' cardActivity
        '
        Me.cardActivity.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.cardActivity.BackColor = System.Drawing.Color.White
        Me.cardActivity.Controls.Add(Me.lblStatusIcon)
        Me.cardActivity.Controls.Add(Me.lblStatus)
        Me.cardActivity.Controls.Add(Me.lblPercent)
        Me.cardActivity.Controls.Add(Me.progressBar1)
        Me.cardActivity.Controls.Add(Me.lblProgressDetail)
        Me.cardActivity.Location = New System.Drawing.Point(24, 322)
        Me.cardActivity.Name = "cardActivity"
        Me.cardActivity.Size = New System.Drawing.Size(472, 96)
        Me.cardActivity.TabIndex = 3
        '
        ' lblStatusIcon
        '
        Me.lblStatusIcon.BackColor = System.Drawing.Color.White
        Me.lblStatusIcon.Font = New System.Drawing.Font("Segoe UI", 10.0!)
        Me.lblStatusIcon.ForeColor = System.Drawing.Color.FromArgb(237, 1, 127)
        Me.lblStatusIcon.Location = New System.Drawing.Point(16, 16)
        Me.lblStatusIcon.Name = "lblStatusIcon"
        Me.lblStatusIcon.Size = New System.Drawing.Size(24, 22)
        Me.lblStatusIcon.TabIndex = 0
        Me.lblStatusIcon.Text = ""
        Me.lblStatusIcon.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        ' lblStatus
        '
        Me.lblStatus.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.lblStatus.BackColor = System.Drawing.Color.White
        Me.lblStatus.Font = New System.Drawing.Font("Segoe UI", 9.5!, System.Drawing.FontStyle.Regular)
        Me.lblStatus.ForeColor = System.Drawing.Color.FromArgb(51, 51, 51)
        Me.lblStatus.Location = New System.Drawing.Point(44, 18)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(336, 20)
        Me.lblStatus.TabIndex = 1
        Me.lblStatus.Text = "Ready for update check"
        '
        ' lblPercent
        '
        Me.lblPercent.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.lblPercent.BackColor = System.Drawing.Color.White
        Me.lblPercent.Font = New System.Drawing.Font("Segoe UI", 9.5!, System.Drawing.FontStyle.Bold)
        Me.lblPercent.ForeColor = System.Drawing.Color.FromArgb(237, 1, 127)
        Me.lblPercent.Location = New System.Drawing.Point(388, 18)
        Me.lblPercent.Name = "lblPercent"
        Me.lblPercent.Size = New System.Drawing.Size(66, 20)
        Me.lblPercent.TabIndex = 2
        Me.lblPercent.Text = ""
        Me.lblPercent.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        ' progressBar1
        '
        Me.progressBar1.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.progressBar1.Location = New System.Drawing.Point(18, 50)
        Me.progressBar1.Name = "progressBar1"
        Me.progressBar1.Size = New System.Drawing.Size(436, 6)
        Me.progressBar1.TabIndex = 3
        '
        ' lblProgressDetail
        '
        Me.lblProgressDetail.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.lblProgressDetail.BackColor = System.Drawing.Color.White
        Me.lblProgressDetail.Font = New System.Drawing.Font("Segoe UI", 8.5!)
        Me.lblProgressDetail.ForeColor = System.Drawing.Color.FromArgb(138, 138, 144)
        Me.lblProgressDetail.Location = New System.Drawing.Point(18, 64)
        Me.lblProgressDetail.Name = "lblProgressDetail"
        Me.lblProgressDetail.Size = New System.Drawing.Size(436, 18)
        Me.lblProgressDetail.TabIndex = 4
        Me.lblProgressDetail.Text = ""
        '
        ' lblCurrentVersion
        '
        Me.lblCurrentVersion.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left, System.Windows.Forms.AnchorStyles)
        Me.lblCurrentVersion.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular)
        Me.lblCurrentVersion.ForeColor = System.Drawing.Color.FromArgb(138, 138, 144)
        Me.lblCurrentVersion.Location = New System.Drawing.Point(24, 432)
        Me.lblCurrentVersion.Name = "lblCurrentVersion"
        Me.lblCurrentVersion.Size = New System.Drawing.Size(300, 18)
        Me.lblCurrentVersion.TabIndex = 4
        Me.lblCurrentVersion.Text = "Current version: ..."
        '
        ' badgeUpdate
        '
        Me.badgeUpdate.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.badgeUpdate.Location = New System.Drawing.Point(396, 430)
        Me.badgeUpdate.Name = "badgeUpdate"
        Me.badgeUpdate.Size = New System.Drawing.Size(100, 22)
        Me.badgeUpdate.TabIndex = 5
        Me.badgeUpdate.Text = ""
        '
        ' pnlFooter — thin magenta accent bar at the bottom
        '
        Me.pnlFooter.BackColor = System.Drawing.Color.FromArgb(237, 1, 127)
        Me.pnlFooter.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlFooter.Location = New System.Drawing.Point(0, 459)
        Me.pnlFooter.Name = "pnlFooter"
        Me.pnlFooter.Size = New System.Drawing.Size(520, 3)
        Me.pnlFooter.TabIndex = 6
        '
        ' MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.Color.FromArgb(246, 246, 248)
        Me.ClientSize = New System.Drawing.Size(520, 462)
        Me.Controls.Add(Me.pnlFooter)
        Me.Controls.Add(Me.badgeUpdate)
        Me.Controls.Add(Me.lblCurrentVersion)
        Me.Controls.Add(Me.cardActivity)
        Me.Controls.Add(Me.btnCheckForUpdates)
        Me.Controls.Add(Me.cardCredentials)
        Me.Controls.Add(Me.pnlHeader)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.Name = "MainForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "VAST Updater"
        Me.pnlHeader.ResumeLayout(False)
        Me.pnlHeader.PerformLayout()
        Me.cardCredentials.ResumeLayout(False)
        Me.cardCredentials.PerformLayout()
        Me.cardActivity.ResumeLayout(False)
        Me.ResumeLayout(False)
    End Sub

End Class
