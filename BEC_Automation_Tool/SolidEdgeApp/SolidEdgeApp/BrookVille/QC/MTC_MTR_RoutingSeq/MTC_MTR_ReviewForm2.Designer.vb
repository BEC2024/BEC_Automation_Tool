<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MTC_MTR_ReviewForm2
    Inherits System.Windows.Forms.Form

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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.chkAll = New System.Windows.Forms.CheckBox()
        Me.chkAssembly = New System.Windows.Forms.CheckBox()
        Me.chkPart = New System.Windows.Forms.CheckBox()
        Me.chkSheetMetal = New System.Windows.Forms.CheckBox()
        Me.lblBaseLineDirectoryPath = New System.Windows.Forms.Label()
        Me.txtBaseLineDirectoryPath = New System.Windows.Forms.TextBox()
        Me.BtnBrowseBaselineDirPath = New System.Windows.Forms.Button()
        Me.BtnExportMTC_MTR_RoutingSeq_Report = New System.Windows.Forms.Button()
        Me.BtnClose = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.dgvDocumentDetails = New System.Windows.Forms.DataGridView()
        Me.BtnBrowseExportDirMTR = New System.Windows.Forms.Button()
        Me.txtExportDirLocationMTR = New System.Windows.Forms.TextBox()
        Me.lblExportDirectoryLocationMTC = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.BtnBrowseExportDirRouting = New System.Windows.Forms.Button()
        Me.txtExportDirLocationRouting = New System.Windows.Forms.TextBox()
        Me.lblExportDirectoryLocationRouting = New System.Windows.Forms.Label()
        CType(Me.dgvDocumentDetails, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkAll
        '
        Me.chkAll.AutoSize = True
        Me.chkAll.Location = New System.Drawing.Point(60, 13)
        Me.chkAll.Name = "chkAll"
        Me.chkAll.Size = New System.Drawing.Size(40, 19)
        Me.chkAll.TabIndex = 0
        Me.chkAll.Text = "All"
        Me.chkAll.UseVisualStyleBackColor = True
        '
        'chkAssembly
        '
        Me.chkAssembly.AutoSize = True
        Me.chkAssembly.Location = New System.Drawing.Point(22, 53)
        Me.chkAssembly.Name = "chkAssembly"
        Me.chkAssembly.Size = New System.Drawing.Size(77, 19)
        Me.chkAssembly.TabIndex = 1
        Me.chkAssembly.Text = "Assembly"
        Me.chkAssembly.UseVisualStyleBackColor = True
        '
        'chkPart
        '
        Me.chkPart.AutoSize = True
        Me.chkPart.Location = New System.Drawing.Point(117, 53)
        Me.chkPart.Name = "chkPart"
        Me.chkPart.Size = New System.Drawing.Size(47, 19)
        Me.chkPart.TabIndex = 2
        Me.chkPart.Text = "Part"
        Me.chkPart.UseVisualStyleBackColor = True
        '
        'chkSheetMetal
        '
        Me.chkSheetMetal.AutoSize = True
        Me.chkSheetMetal.Location = New System.Drawing.Point(179, 53)
        Me.chkSheetMetal.Name = "chkSheetMetal"
        Me.chkSheetMetal.Size = New System.Drawing.Size(88, 19)
        Me.chkSheetMetal.TabIndex = 3
        Me.chkSheetMetal.Text = "Sheet Metal"
        Me.chkSheetMetal.UseVisualStyleBackColor = True
        '
        'lblBaseLineDirectoryPath
        '
        Me.lblBaseLineDirectoryPath.AutoSize = True
        Me.lblBaseLineDirectoryPath.Location = New System.Drawing.Point(25, 27)
        Me.lblBaseLineDirectoryPath.Name = "lblBaseLineDirectoryPath"
        Me.lblBaseLineDirectoryPath.Size = New System.Drawing.Size(128, 15)
        Me.lblBaseLineDirectoryPath.TabIndex = 4
        Me.lblBaseLineDirectoryPath.Text = "Baseline Directory Path"
        '
        'txtBaseLineDirectoryPath
        '
        Me.txtBaseLineDirectoryPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtBaseLineDirectoryPath.Enabled = False
        Me.txtBaseLineDirectoryPath.Location = New System.Drawing.Point(214, 23)
        Me.txtBaseLineDirectoryPath.Name = "txtBaseLineDirectoryPath"
        Me.txtBaseLineDirectoryPath.Size = New System.Drawing.Size(753, 23)
        Me.txtBaseLineDirectoryPath.TabIndex = 5
        '
        'BtnBrowseBaselineDirPath
        '
        Me.BtnBrowseBaselineDirPath.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnBrowseBaselineDirPath.Location = New System.Drawing.Point(973, 22)
        Me.BtnBrowseBaselineDirPath.Name = "BtnBrowseBaselineDirPath"
        Me.BtnBrowseBaselineDirPath.Size = New System.Drawing.Size(77, 24)
        Me.BtnBrowseBaselineDirPath.TabIndex = 6
        Me.BtnBrowseBaselineDirPath.Text = "Browse"
        Me.BtnBrowseBaselineDirPath.UseVisualStyleBackColor = True
        Me.BtnBrowseBaselineDirPath.Visible = False
        '
        'BtnExportMTC_MTR_RoutingSeq_Report
        '
        Me.BtnExportMTC_MTR_RoutingSeq_Report.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnExportMTC_MTR_RoutingSeq_Report.Location = New System.Drawing.Point(942, 558)
        Me.BtnExportMTC_MTR_RoutingSeq_Report.Name = "BtnExportMTC_MTR_RoutingSeq_Report"
        Me.BtnExportMTC_MTR_RoutingSeq_Report.Size = New System.Drawing.Size(101, 24)
        Me.BtnExportMTC_MTR_RoutingSeq_Report.TabIndex = 7
        Me.BtnExportMTC_MTR_RoutingSeq_Report.Text = "Export Reports"
        Me.BtnExportMTC_MTR_RoutingSeq_Report.UseVisualStyleBackColor = True
        '
        'BtnClose
        '
        Me.BtnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnClose.Location = New System.Drawing.Point(859, 558)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(77, 24)
        Me.BtnClose.TabIndex = 8
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'dgvDocumentDetails
        '
        Me.dgvDocumentDetails.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvDocumentDetails.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDocumentDetails.Location = New System.Drawing.Point(25, 85)
        Me.dgvDocumentDetails.Name = "dgvDocumentDetails"
        Me.dgvDocumentDetails.Size = New System.Drawing.Size(1018, 467)
        Me.dgvDocumentDetails.TabIndex = 9
        '
        'BtnBrowseExportDirMTR
        '
        Me.BtnBrowseExportDirMTR.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnBrowseExportDirMTR.Location = New System.Drawing.Point(973, 52)
        Me.BtnBrowseExportDirMTR.Name = "BtnBrowseExportDirMTR"
        Me.BtnBrowseExportDirMTR.Size = New System.Drawing.Size(77, 24)
        Me.BtnBrowseExportDirMTR.TabIndex = 12
        Me.BtnBrowseExportDirMTR.Text = "Browse"
        Me.BtnBrowseExportDirMTR.UseVisualStyleBackColor = True
        Me.BtnBrowseExportDirMTR.Visible = False
        '
        'txtExportDirLocationMTR
        '
        Me.txtExportDirLocationMTR.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtExportDirLocationMTR.Enabled = False
        Me.txtExportDirLocationMTR.Location = New System.Drawing.Point(214, 53)
        Me.txtExportDirLocationMTR.Name = "txtExportDirLocationMTR"
        Me.txtExportDirLocationMTR.Size = New System.Drawing.Size(753, 23)
        Me.txtExportDirLocationMTR.TabIndex = 11
        '
        'lblExportDirectoryLocationMTC
        '
        Me.lblExportDirectoryLocationMTC.AutoSize = True
        Me.lblExportDirectoryLocationMTC.Location = New System.Drawing.Point(25, 57)
        Me.lblExportDirectoryLocationMTC.Name = "lblExportDirectoryLocationMTC"
        Me.lblExportDirectoryLocationMTC.Size = New System.Drawing.Size(123, 15)
        Me.lblExportDirectoryLocationMTC.TabIndex = 10
        Me.lblExportDirectoryLocationMTC.Text = "MTC Output Directory"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.chkAll)
        Me.Panel1.Controls.Add(Me.chkSheetMetal)
        Me.Panel1.Controls.Add(Me.chkAssembly)
        Me.Panel1.Controls.Add(Me.chkPart)
        Me.Panel1.Location = New System.Drawing.Point(44, 169)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(418, 100)
        Me.Panel1.TabIndex = 13
        Me.Panel1.Visible = False
        '
        'BtnBrowseExportDirRouting
        '
        Me.BtnBrowseExportDirRouting.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnBrowseExportDirRouting.Location = New System.Drawing.Point(973, 77)
        Me.BtnBrowseExportDirRouting.Name = "BtnBrowseExportDirRouting"
        Me.BtnBrowseExportDirRouting.Size = New System.Drawing.Size(77, 24)
        Me.BtnBrowseExportDirRouting.TabIndex = 16
        Me.BtnBrowseExportDirRouting.Text = "Browse"
        Me.BtnBrowseExportDirRouting.UseVisualStyleBackColor = True
        Me.BtnBrowseExportDirRouting.Visible = False
        '
        'txtExportDirLocationRouting
        '
        Me.txtExportDirLocationRouting.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtExportDirLocationRouting.Enabled = False
        Me.txtExportDirLocationRouting.Location = New System.Drawing.Point(214, 78)
        Me.txtExportDirLocationRouting.Name = "txtExportDirLocationRouting"
        Me.txtExportDirLocationRouting.Size = New System.Drawing.Size(753, 23)
        Me.txtExportDirLocationRouting.TabIndex = 15
        Me.txtExportDirLocationRouting.Visible = False
        '
        'lblExportDirectoryLocationRouting
        '
        Me.lblExportDirectoryLocationRouting.AutoSize = True
        Me.lblExportDirectoryLocationRouting.Location = New System.Drawing.Point(25, 82)
        Me.lblExportDirectoryLocationRouting.Name = "lblExportDirectoryLocationRouting"
        Me.lblExportDirectoryLocationRouting.Size = New System.Drawing.Size(185, 15)
        Me.lblExportDirectoryLocationRouting.TabIndex = 14
        Me.lblExportDirectoryLocationRouting.Text = "Export Routing Sequence Dir Path"
        Me.lblExportDirectoryLocationRouting.Visible = False
        '
        'MTC_MTR_ReviewForm2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1062, 591)
        Me.Controls.Add(Me.BtnBrowseExportDirRouting)
        Me.Controls.Add(Me.txtExportDirLocationRouting)
        Me.Controls.Add(Me.lblExportDirectoryLocationRouting)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.BtnBrowseExportDirMTR)
        Me.Controls.Add(Me.txtExportDirLocationMTR)
        Me.Controls.Add(Me.lblExportDirectoryLocationMTC)
        Me.Controls.Add(Me.dgvDocumentDetails)
        Me.Controls.Add(Me.BtnClose)
        Me.Controls.Add(Me.BtnExportMTC_MTR_RoutingSeq_Report)
        Me.Controls.Add(Me.BtnBrowseBaselineDirPath)
        Me.Controls.Add(Me.txtBaseLineDirectoryPath)
        Me.Controls.Add(Me.lblBaseLineDirectoryPath)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "MTC_MTR_ReviewForm2"
        Me.Text = "MTC MTR Review"
        CType(Me.dgvDocumentDetails, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents chkAll As CheckBox
    Friend WithEvents chkAssembly As CheckBox
    Friend WithEvents chkPart As CheckBox
    Friend WithEvents chkSheetMetal As CheckBox
    Friend WithEvents lblBaseLineDirectoryPath As Label
    Friend WithEvents txtBaseLineDirectoryPath As TextBox
    Friend WithEvents BtnBrowseBaselineDirPath As Button
    Friend WithEvents BtnExportMTC_MTR_RoutingSeq_Report As Button
    Friend WithEvents BtnClose As Button
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents dgvDocumentDetails As DataGridView
    Friend WithEvents BtnBrowseExportDirMTR As Button
    Friend WithEvents txtExportDirLocationMTR As TextBox
    Friend WithEvents lblExportDirectoryLocationMTC As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents BtnBrowseExportDirRouting As Button
    Friend WithEvents txtExportDirLocationRouting As TextBox
    Friend WithEvents lblExportDirectoryLocationRouting As Label
End Class
