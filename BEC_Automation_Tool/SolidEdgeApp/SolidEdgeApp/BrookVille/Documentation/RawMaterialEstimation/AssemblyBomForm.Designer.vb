<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class AssemblyBomForm
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
        Me.btnGetData = New System.Windows.Forms.Button()
        Me.lblAssemblyName = New System.Windows.Forms.Label()
        Me.dgvDocumentDetails = New System.Windows.Forms.DataGridView()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnExportExcel = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.btnTemplateLocation = New System.Windows.Forms.Button()
        Me.txtRawMaterialEstimationReportDirPath = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnBrowseRawMaterialBOM = New System.Windows.Forms.Button()
        Me.txtBecMaterialExcel = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        CType(Me.dgvDocumentDetails, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnGetData
        '
        Me.btnGetData.Location = New System.Drawing.Point(12, 60)
        Me.btnGetData.Name = "btnGetData"
        Me.btnGetData.Size = New System.Drawing.Size(190, 30)
        Me.btnGetData.TabIndex = 5
        Me.btnGetData.Text = "Get Current Assembly Data"
        Me.btnGetData.UseVisualStyleBackColor = True
        '
        'lblAssemblyName
        '
        Me.lblAssemblyName.AutoSize = True
        Me.lblAssemblyName.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAssemblyName.Location = New System.Drawing.Point(213, 67)
        Me.lblAssemblyName.Name = "lblAssemblyName"
        Me.lblAssemblyName.Size = New System.Drawing.Size(102, 17)
        Me.lblAssemblyName.TabIndex = 4
        Me.lblAssemblyName.Text = "Assembly Name"
        '
        'dgvDocumentDetails
        '
        Me.dgvDocumentDetails.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvDocumentDetails.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDocumentDetails.Location = New System.Drawing.Point(12, 135)
        Me.dgvDocumentDetails.Name = "dgvDocumentDetails"
        Me.dgvDocumentDetails.Size = New System.Drawing.Size(1038, 410)
        Me.dgvDocumentDetails.TabIndex = 6
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(827, 554)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(100, 30)
        Me.btnClose.TabIndex = 7
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnExportExcel
        '
        Me.btnExportExcel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportExcel.Location = New System.Drawing.Point(933, 554)
        Me.btnExportExcel.Name = "btnExportExcel"
        Me.btnExportExcel.Size = New System.Drawing.Size(117, 30)
        Me.btnExportExcel.TabIndex = 8
        Me.btnExportExcel.Text = "Export Excel"
        Me.btnExportExcel.UseVisualStyleBackColor = True
        '
        'btnTemplateLocation
        '
        Me.btnTemplateLocation.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnTemplateLocation.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnTemplateLocation.Location = New System.Drawing.Point(933, 96)
        Me.btnTemplateLocation.Name = "btnTemplateLocation"
        Me.btnTemplateLocation.Size = New System.Drawing.Size(117, 30)
        Me.btnTemplateLocation.TabIndex = 24
        Me.btnTemplateLocation.Text = "Browse"
        Me.btnTemplateLocation.UseVisualStyleBackColor = True
        '
        'txtRawMaterialEstimationReportDirPath
        '
        Me.txtRawMaterialEstimationReportDirPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRawMaterialEstimationReportDirPath.Location = New System.Drawing.Point(177, 100)
        Me.txtRawMaterialEstimationReportDirPath.Name = "txtRawMaterialEstimationReportDirPath"
        Me.txtRawMaterialEstimationReportDirPath.Size = New System.Drawing.Size(750, 23)
        Me.txtRawMaterialEstimationReportDirPath.TabIndex = 23
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 104)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 15)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "Output Directory"
        '
        'btnBrowseRawMaterialBOM
        '
        Me.btnBrowseRawMaterialBOM.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrowseRawMaterialBOM.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnBrowseRawMaterialBOM.Location = New System.Drawing.Point(933, 18)
        Me.btnBrowseRawMaterialBOM.Name = "btnBrowseRawMaterialBOM"
        Me.btnBrowseRawMaterialBOM.Size = New System.Drawing.Size(117, 30)
        Me.btnBrowseRawMaterialBOM.TabIndex = 27
        Me.btnBrowseRawMaterialBOM.Text = "Browse"
        Me.btnBrowseRawMaterialBOM.UseVisualStyleBackColor = True
        Me.btnBrowseRawMaterialBOM.Visible = False
        '
        'txtBecMaterialExcel
        '
        Me.txtBecMaterialExcel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtBecMaterialExcel.Enabled = False
        Me.txtBecMaterialExcel.Location = New System.Drawing.Point(177, 22)
        Me.txtBecMaterialExcel.Name = "txtBecMaterialExcel"
        Me.txtBecMaterialExcel.Size = New System.Drawing.Size(750, 23)
        Me.txtBecMaterialExcel.TabIndex = 26
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(15, 26)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(131, 15)
        Me.Label2.TabIndex = 25
        Me.Label2.Text = "BEC Material Excel Path"
        '
        'AssemblyBomForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1062, 591)
        Me.Controls.Add(Me.btnBrowseRawMaterialBOM)
        Me.Controls.Add(Me.txtBecMaterialExcel)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnTemplateLocation)
        Me.Controls.Add(Me.txtRawMaterialEstimationReportDirPath)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnExportExcel)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.dgvDocumentDetails)
        Me.Controls.Add(Me.btnGetData)
        Me.Controls.Add(Me.lblAssemblyName)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "AssemblyBomForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "RawMaterial"
        CType(Me.dgvDocumentDetails, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnGetData As Button
    Friend WithEvents lblAssemblyName As Label
    Friend WithEvents dgvDocumentDetails As DataGridView
    Friend WithEvents btnClose As Button
    Friend WithEvents btnExportExcel As Button
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents btnTemplateLocation As Button
    Friend WithEvents txtRawMaterialEstimationReportDirPath As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents btnBrowseRawMaterialBOM As Button
    Friend WithEvents txtBecMaterialExcel As TextBox
    Friend WithEvents Label2 As Label
End Class
