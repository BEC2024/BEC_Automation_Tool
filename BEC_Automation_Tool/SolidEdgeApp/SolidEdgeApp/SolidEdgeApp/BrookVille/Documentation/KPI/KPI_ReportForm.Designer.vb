<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class KPI_ReportForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.radMTR = New System.Windows.Forms.RadioButton()
        Me.radMTC = New System.Windows.Forms.RadioButton()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.chkMTR = New System.Windows.Forms.CheckBox()
        Me.chkMTC = New System.Windows.Forms.CheckBox()
        Me.btnGenerateReport = New System.Windows.Forms.Button()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtFolder = New System.Windows.Forms.TextBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.btnCreateReport = New System.Windows.Forms.Button()
        Me.btnFolderBrowse = New System.Windows.Forms.Button()
        Me.lblMTCReportDirPath = New System.Windows.Forms.Label()
        Me.txtFilename = New System.Windows.Forms.TextBox()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.AccessibleRole = System.Windows.Forms.AccessibleRole.None
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.CheckBox2)
        Me.Panel1.Controls.Add(Me.CheckBox1)
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.btnCreateReport)
        Me.Panel1.Controls.Add(Me.btnFolderBrowse)
        Me.Panel1.Controls.Add(Me.lblMTCReportDirPath)
        Me.Panel1.Controls.Add(Me.txtFilename)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1062, 591)
        Me.Panel1.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.AccessibleRole = System.Windows.Forms.AccessibleRole.None
        Me.Panel2.BackColor = System.Drawing.SystemColors.Control
        Me.Panel2.Controls.Add(Me.radMTR)
        Me.Panel2.Controls.Add(Me.radMTC)
        Me.Panel2.Controls.Add(Me.btnClose)
        Me.Panel2.Controls.Add(Me.chkMTR)
        Me.Panel2.Controls.Add(Me.chkMTC)
        Me.Panel2.Controls.Add(Me.btnGenerateReport)
        Me.Panel2.Controls.Add(Me.btnBrowse)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.txtFolder)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1062, 591)
        Me.Panel2.TabIndex = 9
        '
        'radMTR
        '
        Me.radMTR.AutoSize = True
        Me.radMTR.Location = New System.Drawing.Point(238, 25)
        Me.radMTR.Name = "radMTR"
        Me.radMTR.Size = New System.Drawing.Size(49, 19)
        Me.radMTR.TabIndex = 11
        Me.radMTR.TabStop = True
        Me.radMTR.Text = "MTR"
        Me.radMTR.UseVisualStyleBackColor = True
        Me.radMTR.Visible = False
        '
        'radMTC
        '
        Me.radMTC.AutoSize = True
        Me.radMTC.Location = New System.Drawing.Point(182, 24)
        Me.radMTC.Name = "radMTC"
        Me.radMTC.Size = New System.Drawing.Size(49, 19)
        Me.radMTC.TabIndex = 10
        Me.radMTC.TabStop = True
        Me.radMTC.Text = "MTC"
        Me.radMTC.UseVisualStyleBackColor = True
        Me.radMTC.Visible = False
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(809, 549)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(100, 30)
        Me.btnClose.TabIndex = 9
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'chkMTR
        '
        Me.chkMTR.AutoSize = True
        Me.chkMTR.Location = New System.Drawing.Point(238, 87)
        Me.chkMTR.Name = "chkMTR"
        Me.chkMTR.Size = New System.Drawing.Size(50, 19)
        Me.chkMTR.TabIndex = 8
        Me.chkMTR.Text = "MTR"
        Me.chkMTR.UseVisualStyleBackColor = True
        Me.chkMTR.Visible = False
        '
        'chkMTC
        '
        Me.chkMTC.AutoSize = True
        Me.chkMTC.Location = New System.Drawing.Point(182, 87)
        Me.chkMTC.Name = "chkMTC"
        Me.chkMTC.Size = New System.Drawing.Size(50, 19)
        Me.chkMTC.TabIndex = 7
        Me.chkMTC.Text = "MTC"
        Me.chkMTC.UseVisualStyleBackColor = True
        Me.chkMTC.Visible = False
        '
        'btnGenerateReport
        '
        Me.btnGenerateReport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGenerateReport.Location = New System.Drawing.Point(915, 549)
        Me.btnGenerateReport.Name = "btnGenerateReport"
        Me.btnGenerateReport.Size = New System.Drawing.Size(135, 30)
        Me.btnGenerateReport.TabIndex = 6
        Me.btnGenerateReport.Text = "Generate Report"
        Me.btnGenerateReport.UseVisualStyleBackColor = True
        '
        'btnBrowse
        '
        Me.btnBrowse.Location = New System.Drawing.Point(751, 46)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(100, 30)
        Me.btnBrowse.TabIndex = 2
        Me.btnBrowse.Text = "Browser"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(14, 54)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(119, 15)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "MTC Reports Dir Path"
        '
        'txtFolder
        '
        Me.txtFolder.Location = New System.Drawing.Point(182, 50)
        Me.txtFolder.Name = "txtFolder"
        Me.txtFolder.Size = New System.Drawing.Size(562, 23)
        Me.txtFolder.TabIndex = 0
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(246, 24)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(50, 19)
        Me.CheckBox2.TabIndex = 8
        Me.CheckBox2.Text = "MTR"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(182, 24)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(50, 19)
        Me.CheckBox1.TabIndex = 7
        Me.CheckBox1.Text = "MTC"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(751, 82)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(87, 27)
        Me.Button2.TabIndex = 6
        Me.Button2.Text = "Total"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(337, 309)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(87, 27)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "MTR Report"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'btnCreateReport
        '
        Me.btnCreateReport.Location = New System.Drawing.Point(243, 309)
        Me.btnCreateReport.Name = "btnCreateReport"
        Me.btnCreateReport.Size = New System.Drawing.Size(87, 27)
        Me.btnCreateReport.TabIndex = 3
        Me.btnCreateReport.Text = "MTC Report"
        Me.btnCreateReport.UseVisualStyleBackColor = True
        '
        'btnFolderBrowse
        '
        Me.btnFolderBrowse.Location = New System.Drawing.Point(751, 48)
        Me.btnFolderBrowse.Name = "btnFolderBrowse"
        Me.btnFolderBrowse.Size = New System.Drawing.Size(87, 27)
        Me.btnFolderBrowse.TabIndex = 2
        Me.btnFolderBrowse.Text = "Browser"
        Me.btnFolderBrowse.UseVisualStyleBackColor = True
        '
        'lblMTCReportDirPath
        '
        Me.lblMTCReportDirPath.AutoSize = True
        Me.lblMTCReportDirPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMTCReportDirPath.Location = New System.Drawing.Point(14, 54)
        Me.lblMTCReportDirPath.Name = "lblMTCReportDirPath"
        Me.lblMTCReportDirPath.Size = New System.Drawing.Size(138, 13)
        Me.lblMTCReportDirPath.TabIndex = 1
        Me.lblMTCReportDirPath.Text = "MTC-MTR Reports Dir Path"
        '
        'txtFilename
        '
        Me.txtFilename.Location = New System.Drawing.Point(182, 50)
        Me.txtFilename.Name = "txtFilename"
        Me.txtFilename.Size = New System.Drawing.Size(562, 23)
        Me.txtFilename.TabIndex = 0
        '
        'KPI_ReportForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1062, 591)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "KPI_ReportForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "KPI Report"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As Panel
    Friend WithEvents btnFolderBrowse As Button
    Friend WithEvents lblMTCReportDirPath As Label
    Friend WithEvents txtFilename As TextBox
    Friend WithEvents btnCreateReport As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents Panel2 As Panel
    Friend WithEvents btnClose As Button
    Friend WithEvents chkMTR As CheckBox
    Friend WithEvents chkMTC As CheckBox
    Friend WithEvents btnGenerateReport As Button
    Friend WithEvents btnBrowse As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents txtFolder As TextBox
    Friend WithEvents CheckBox2 As CheckBox
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents radMTR As RadioButton
    Friend WithEvents radMTC As RadioButton
End Class
