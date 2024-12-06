<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class CheckIntereferenceForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CheckIntereferenceForm))
        Me.lblExcelPath = New System.Windows.Forms.Label()
        Me.btnBrowseExcel = New System.Windows.Forms.Button()
        Me.txtExcelPath = New System.Windows.Forms.TextBox()
        Me.dgvDocumentDetails = New System.Windows.Forms.DataGridView()
        Me.btnGenerateInterferenceReport = New System.Windows.Forms.Button()
        Me.ComboBoxFields = New System.Windows.Forms.ComboBox()
        Me.txtSearch = New System.Windows.Forms.TextBox()
        Me.btnSearchFile = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.dgvDocumentDetails, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblExcelPath
        '
        Me.lblExcelPath.AutoSize = True
        Me.lblExcelPath.Location = New System.Drawing.Point(17, 17)
        Me.lblExcelPath.Name = "lblExcelPath"
        Me.lblExcelPath.Size = New System.Drawing.Size(61, 15)
        Me.lblExcelPath.TabIndex = 21
        Me.lblExcelPath.Text = "Excel Path"
        '
        'btnBrowseExcel
        '
        Me.btnBrowseExcel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrowseExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnBrowseExcel.Location = New System.Drawing.Point(558, 10)
        Me.btnBrowseExcel.Name = "btnBrowseExcel"
        Me.btnBrowseExcel.Size = New System.Drawing.Size(78, 29)
        Me.btnBrowseExcel.TabIndex = 23
        Me.btnBrowseExcel.Text = "Browse"
        Me.btnBrowseExcel.UseVisualStyleBackColor = True
        '
        'txtExcelPath
        '
        Me.txtExcelPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtExcelPath.Location = New System.Drawing.Point(81, 13)
        Me.txtExcelPath.Name = "txtExcelPath"
        Me.txtExcelPath.Size = New System.Drawing.Size(469, 23)
        Me.txtExcelPath.TabIndex = 22
        '
        'dgvDocumentDetails
        '
        Me.dgvDocumentDetails.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvDocumentDetails.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDocumentDetails.Location = New System.Drawing.Point(17, 45)
        Me.dgvDocumentDetails.Name = "dgvDocumentDetails"
        Me.dgvDocumentDetails.Size = New System.Drawing.Size(619, 386)
        Me.dgvDocumentDetails.TabIndex = 24
        '
        'btnGenerateInterferenceReport
        '
        Me.btnGenerateInterferenceReport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGenerateInterferenceReport.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnGenerateInterferenceReport.Location = New System.Drawing.Point(469, 437)
        Me.btnGenerateInterferenceReport.Name = "btnGenerateInterferenceReport"
        Me.btnGenerateInterferenceReport.Size = New System.Drawing.Size(167, 29)
        Me.btnGenerateInterferenceReport.TabIndex = 26
        Me.btnGenerateInterferenceReport.Text = "Generate Interference Report"
        Me.btnGenerateInterferenceReport.UseVisualStyleBackColor = True
        '
        'ComboBoxFields
        '
        Me.ComboBoxFields.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ComboBoxFields.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxFields.FormattingEnabled = True
        Me.ComboBoxFields.Location = New System.Drawing.Point(17, 440)
        Me.ComboBoxFields.Name = "ComboBoxFields"
        Me.ComboBoxFields.Size = New System.Drawing.Size(140, 23)
        Me.ComboBoxFields.TabIndex = 37
        '
        'txtSearch
        '
        Me.txtSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtSearch.Location = New System.Drawing.Point(163, 440)
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Size = New System.Drawing.Size(191, 23)
        Me.txtSearch.TabIndex = 35
        '
        'btnSearchFile
        '
        Me.btnSearchFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSearchFile.Image = Global.SolidEdgeApp.My.Resources.Resources.search_16px
        Me.btnSearchFile.Location = New System.Drawing.Point(360, 438)
        Me.btnSearchFile.Name = "btnSearchFile"
        Me.btnSearchFile.Size = New System.Drawing.Size(43, 26)
        Me.btnSearchFile.TabIndex = 36
        Me.btnSearchFile.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(17, 479)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(381, 17)
        Me.Label1.TabIndex = 38
        Me.Label1.Text = "Note: Selected Materials will not be considered for Interference."
        '
        'CheckIntereferenceForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(648, 507)
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
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "CheckIntereferenceForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " Generate Interference Report"
        CType(Me.dgvDocumentDetails, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblExcelPath As Label
    Friend WithEvents btnBrowseExcel As Button
    Friend WithEvents txtExcelPath As TextBox
    Friend WithEvents dgvDocumentDetails As DataGridView
    Friend WithEvents btnGenerateInterferenceReport As Button
    Friend WithEvents ComboBoxFields As ComboBox
    Friend WithEvents txtSearch As TextBox
    Friend WithEvents btnSearchFile As Button
    Friend WithEvents Label1 As Label
End Class
