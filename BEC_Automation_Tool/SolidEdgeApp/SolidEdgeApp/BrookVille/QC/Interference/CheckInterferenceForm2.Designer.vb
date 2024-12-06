<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CheckInterferenceForm2
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ComboBoxFields = New System.Windows.Forms.ComboBox()
        Me.txtSearch = New System.Windows.Forms.TextBox()
        Me.btnSearchFile = New System.Windows.Forms.Button()
        Me.btnGenerateInterferenceReport = New System.Windows.Forms.Button()
        Me.dgvDocumentDetails = New System.Windows.Forms.DataGridView()
        Me.lblExcelPath = New System.Windows.Forms.Label()
        Me.btnBrowseExcel = New System.Windows.Forms.Button()
        Me.txtExcelPath = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.btnInterferenceExcludeMaterial = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        CType(Me.dgvDocumentDetails, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 565)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(381, 17)
        Me.Label1.TabIndex = 47
        Me.Label1.Text = "Note: Selected Materials will not be considered for Interference."
        '
        'ComboBoxFields
        '
        Me.ComboBoxFields.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ComboBoxFields.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxFields.FormattingEnabled = True
        Me.ComboBoxFields.Location = New System.Drawing.Point(15, 525)
        Me.ComboBoxFields.Name = "ComboBoxFields"
        Me.ComboBoxFields.Size = New System.Drawing.Size(140, 23)
        Me.ComboBoxFields.TabIndex = 46
        '
        'txtSearch
        '
        Me.txtSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtSearch.Location = New System.Drawing.Point(161, 525)
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Size = New System.Drawing.Size(140, 23)
        Me.txtSearch.TabIndex = 44
        '
        'btnSearchFile
        '
        Me.btnSearchFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSearchFile.Image = Global.SolidEdgeApp.My.Resources.Resources.search_16px
        Me.btnSearchFile.Location = New System.Drawing.Point(307, 523)
        Me.btnSearchFile.Name = "btnSearchFile"
        Me.btnSearchFile.Size = New System.Drawing.Size(43, 26)
        Me.btnSearchFile.TabIndex = 45
        Me.btnSearchFile.UseVisualStyleBackColor = True
        '
        'btnGenerateInterferenceReport
        '
        Me.btnGenerateInterferenceReport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGenerateInterferenceReport.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnGenerateInterferenceReport.Location = New System.Drawing.Point(552, 521)
        Me.btnGenerateInterferenceReport.Name = "btnGenerateInterferenceReport"
        Me.btnGenerateInterferenceReport.Size = New System.Drawing.Size(245, 30)
        Me.btnGenerateInterferenceReport.TabIndex = 43
        Me.btnGenerateInterferenceReport.Text = "Generate Interference Report"
        Me.btnGenerateInterferenceReport.UseVisualStyleBackColor = True
        Me.btnGenerateInterferenceReport.Visible = False
        '
        'dgvDocumentDetails
        '
        Me.dgvDocumentDetails.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvDocumentDetails.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDocumentDetails.Location = New System.Drawing.Point(15, 78)
        Me.dgvDocumentDetails.Name = "dgvDocumentDetails"
        Me.dgvDocumentDetails.Size = New System.Drawing.Size(1033, 437)
        Me.dgvDocumentDetails.TabIndex = 42
        '
        'lblExcelPath
        '
        Me.lblExcelPath.AutoSize = True
        Me.lblExcelPath.Location = New System.Drawing.Point(15, 17)
        Me.lblExcelPath.Name = "lblExcelPath"
        Me.lblExcelPath.Size = New System.Drawing.Size(61, 15)
        Me.lblExcelPath.TabIndex = 39
        Me.lblExcelPath.Text = "Excel Path"
        '
        'btnBrowseExcel
        '
        Me.btnBrowseExcel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrowseExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnBrowseExcel.Location = New System.Drawing.Point(948, 9)
        Me.btnBrowseExcel.Name = "btnBrowseExcel"
        Me.btnBrowseExcel.Size = New System.Drawing.Size(100, 30)
        Me.btnBrowseExcel.TabIndex = 41
        Me.btnBrowseExcel.Text = "Browse"
        Me.btnBrowseExcel.UseVisualStyleBackColor = True
        Me.btnBrowseExcel.Visible = False
        '
        'txtExcelPath
        '
        Me.txtExcelPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtExcelPath.Enabled = False
        Me.txtExcelPath.Location = New System.Drawing.Point(79, 13)
        Me.txtExcelPath.Name = "txtExcelPath"
        Me.txtExcelPath.Size = New System.Drawing.Size(863, 23)
        Me.txtExcelPath.TabIndex = 40
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(803, 521)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(245, 30)
        Me.Button1.TabIndex = 48
        Me.Button1.Text = "To False Interferance Occurance Properties"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button2.Location = New System.Drawing.Point(803, 558)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(245, 30)
        Me.Button2.TabIndex = 49
        Me.Button2.Text = "Check Interference TopLevel"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button3.Location = New System.Drawing.Point(552, 558)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(245, 30)
        Me.Button3.TabIndex = 50
        Me.Button3.Text = "Check child interferences"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'btnInterferenceExcludeMaterial
        '
        Me.btnInterferenceExcludeMaterial.Location = New System.Drawing.Point(15, 44)
        Me.btnInterferenceExcludeMaterial.Name = "btnInterferenceExcludeMaterial"
        Me.btnInterferenceExcludeMaterial.Size = New System.Drawing.Size(229, 23)
        Me.btnInterferenceExcludeMaterial.TabIndex = 51
        Me.btnInterferenceExcludeMaterial.Text = "Get Interference Exclude Material"
        Me.btnInterferenceExcludeMaterial.UseVisualStyleBackColor = True
        '
        'CheckInterferenceForm2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1062, 591)
        Me.Controls.Add(Me.btnInterferenceExcludeMaterial)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboBoxFields)
        Me.Controls.Add(Me.txtSearch)
        Me.Controls.Add(Me.btnSearchFile)
        Me.Controls.Add(Me.btnGenerateInterferenceReport)
        Me.Controls.Add(Me.dgvDocumentDetails)
        Me.Controls.Add(Me.lblExcelPath)
        Me.Controls.Add(Me.btnBrowseExcel)
        Me.Controls.Add(Me.txtExcelPath)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "CheckInterferenceForm2"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CheckInterferenceForm2"
        CType(Me.dgvDocumentDetails, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents ComboBoxFields As ComboBox
    Friend WithEvents txtSearch As TextBox
    Friend WithEvents btnSearchFile As Button
    Friend WithEvents btnGenerateInterferenceReport As Button
    Friend WithEvents dgvDocumentDetails As DataGridView
    Friend WithEvents lblExcelPath As Label
    Friend WithEvents btnBrowseExcel As Button
    Friend WithEvents txtExcelPath As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents btnInterferenceExcludeMaterial As Button
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
End Class
