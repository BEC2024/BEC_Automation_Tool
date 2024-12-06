<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class CreateNewPartForm
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
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmbBECMaterial = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmbThickness = New System.Windows.Forms.ComboBox()
        Me.lblTemplateName = New System.Windows.Forms.Label()
        Me.cmbMaterialUsed = New System.Windows.Forms.ComboBox()
        Me.lblType = New System.Windows.Forms.Label()
        Me.cmbType = New System.Windows.Forms.ComboBox()
        Me.lblCategory = New System.Windows.Forms.Label()
        Me.cmbCategory = New System.Windows.Forms.ComboBox()
        Me.lblHeight = New System.Windows.Forms.Label()
        Me.lblWidth = New System.Windows.Forms.Label()
        Me.txtHeight = New System.Windows.Forms.TextBox()
        Me.txtWidth = New System.Windows.Forms.TextBox()
        Me.btnCreatePart = New System.Windows.Forms.Button()
        Me.lblGageTable = New System.Windows.Forms.Label()
        Me.cmbGageTable = New System.Windows.Forms.ComboBox()
        Me.lblGageName = New System.Windows.Forms.Label()
        Me.cmbGageName = New System.Windows.Forms.ComboBox()
        Me.txtBendRadius = New System.Windows.Forms.TextBox()
        Me.txtBendType = New System.Windows.Forms.TextBox()
        Me.lblBendRadius = New System.Windows.Forms.Label()
        Me.lblBendType = New System.Windows.Forms.Label()
        Me.txtMaterialLibrary = New System.Windows.Forms.TextBox()
        Me.lblMaterialLibrary = New System.Windows.Forms.Label()
        Me.txtMaterialSpec = New System.Windows.Forms.TextBox()
        Me.lblMaterialSpec = New System.Windows.Forms.Label()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.txtTemplate = New System.Windows.Forms.TextBox()
        Me.lblTemplate = New System.Windows.Forms.Label()
        Me.txtDiameter = New System.Windows.Forms.TextBox()
        Me.lblDiameter = New System.Windows.Forms.Label()
        Me.PanelBody = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.BtnBrowseSolidEdgePartsTemplateDir = New System.Windows.Forms.Button()
        Me.TxtSolidEdgePartsTemplateDirectory = New System.Windows.Forms.TextBox()
        Me.lblExcelPath = New System.Windows.Forms.Label()
        Me.btnBrowseExcel = New System.Windows.Forms.Button()
        Me.txtExcelPath = New System.Windows.Forms.TextBox()
        Me.btnGetData = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.txtFileName = New System.Windows.Forms.TextBox()
        Me.lblFileName = New System.Windows.Forms.Label()
        Me.txtLinearLength = New System.Windows.Forms.TextBox()
        Me.lblLinearLength = New System.Windows.Forms.Label()
        Me.PanelBody.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(21, 295)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(74, 15)
        Me.Label3.TabIndex = 73
        Me.Label3.Text = "BEC Material"
        '
        'cmbBECMaterial
        '
        Me.cmbBECMaterial.FormattingEnabled = True
        Me.cmbBECMaterial.Location = New System.Drawing.Point(164, 291)
        Me.cmbBECMaterial.Name = "cmbBECMaterial"
        Me.cmbBECMaterial.Size = New System.Drawing.Size(475, 23)
        Me.cmbBECMaterial.TabIndex = 72
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(21, 331)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(58, 15)
        Me.Label2.TabIndex = 71
        Me.Label2.Text = "Thickness"
        '
        'cmbThickness
        '
        Me.cmbThickness.FormattingEnabled = True
        Me.cmbThickness.Location = New System.Drawing.Point(164, 328)
        Me.cmbThickness.Name = "cmbThickness"
        Me.cmbThickness.Size = New System.Drawing.Size(475, 23)
        Me.cmbThickness.TabIndex = 70
        '
        'lblTemplateName
        '
        Me.lblTemplateName.AutoSize = True
        Me.lblTemplateName.Location = New System.Drawing.Point(21, 222)
        Me.lblTemplateName.Name = "lblTemplateName"
        Me.lblTemplateName.Size = New System.Drawing.Size(136, 15)
        Me.lblTemplateName.TabIndex = 69
        Me.lblTemplateName.Text = "BEC Code/Material Used"
        '
        'cmbMaterialUsed
        '
        Me.cmbMaterialUsed.FormattingEnabled = True
        Me.cmbMaterialUsed.Location = New System.Drawing.Point(164, 218)
        Me.cmbMaterialUsed.Name = "cmbMaterialUsed"
        Me.cmbMaterialUsed.Size = New System.Drawing.Size(475, 23)
        Me.cmbMaterialUsed.TabIndex = 68
        '
        'lblType
        '
        Me.lblType.AutoSize = True
        Me.lblType.Location = New System.Drawing.Point(21, 185)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(31, 15)
        Me.lblType.TabIndex = 67
        Me.lblType.Text = "Type"
        '
        'cmbType
        '
        Me.cmbType.FormattingEnabled = True
        Me.cmbType.Location = New System.Drawing.Point(164, 181)
        Me.cmbType.Name = "cmbType"
        Me.cmbType.Size = New System.Drawing.Size(475, 23)
        Me.cmbType.TabIndex = 66
        '
        'lblCategory
        '
        Me.lblCategory.AutoSize = True
        Me.lblCategory.Location = New System.Drawing.Point(21, 150)
        Me.lblCategory.Name = "lblCategory"
        Me.lblCategory.Size = New System.Drawing.Size(55, 15)
        Me.lblCategory.TabIndex = 65
        Me.lblCategory.Text = "Category"
        '
        'cmbCategory
        '
        Me.cmbCategory.FormattingEnabled = True
        Me.cmbCategory.Location = New System.Drawing.Point(164, 146)
        Me.cmbCategory.Name = "cmbCategory"
        Me.cmbCategory.Size = New System.Drawing.Size(475, 23)
        Me.cmbCategory.TabIndex = 64
        '
        'lblHeight
        '
        Me.lblHeight.AutoSize = True
        Me.lblHeight.Location = New System.Drawing.Point(21, 369)
        Me.lblHeight.Name = "lblHeight"
        Me.lblHeight.Size = New System.Drawing.Size(43, 15)
        Me.lblHeight.TabIndex = 74
        Me.lblHeight.Text = "Height"
        '
        'lblWidth
        '
        Me.lblWidth.AutoSize = True
        Me.lblWidth.Location = New System.Drawing.Point(21, 434)
        Me.lblWidth.Name = "lblWidth"
        Me.lblWidth.Size = New System.Drawing.Size(39, 15)
        Me.lblWidth.TabIndex = 75
        Me.lblWidth.Text = "Width"
        '
        'txtHeight
        '
        Me.txtHeight.Location = New System.Drawing.Point(164, 366)
        Me.txtHeight.Name = "txtHeight"
        Me.txtHeight.Size = New System.Drawing.Size(475, 23)
        Me.txtHeight.TabIndex = 76
        '
        'txtWidth
        '
        Me.txtWidth.Location = New System.Drawing.Point(164, 431)
        Me.txtWidth.Name = "txtWidth"
        Me.txtWidth.Size = New System.Drawing.Size(475, 23)
        Me.txtWidth.TabIndex = 77
        '
        'btnCreatePart
        '
        Me.btnCreatePart.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCreatePart.Location = New System.Drawing.Point(951, 549)
        Me.btnCreatePart.Name = "btnCreatePart"
        Me.btnCreatePart.Size = New System.Drawing.Size(100, 30)
        Me.btnCreatePart.TabIndex = 78
        Me.btnCreatePart.Text = "Create Part"
        Me.btnCreatePart.UseVisualStyleBackColor = True
        '
        'lblGageTable
        '
        Me.lblGageTable.AutoSize = True
        Me.lblGageTable.Location = New System.Drawing.Point(21, 369)
        Me.lblGageTable.Name = "lblGageTable"
        Me.lblGageTable.Size = New System.Drawing.Size(64, 15)
        Me.lblGageTable.TabIndex = 80
        Me.lblGageTable.Text = "Gage Table"
        '
        'cmbGageTable
        '
        Me.cmbGageTable.FormattingEnabled = True
        Me.cmbGageTable.Location = New System.Drawing.Point(164, 366)
        Me.cmbGageTable.Name = "cmbGageTable"
        Me.cmbGageTable.Size = New System.Drawing.Size(475, 23)
        Me.cmbGageTable.TabIndex = 79
        '
        'lblGageName
        '
        Me.lblGageName.AutoSize = True
        Me.lblGageName.Location = New System.Drawing.Point(21, 402)
        Me.lblGageName.Name = "lblGageName"
        Me.lblGageName.Size = New System.Drawing.Size(69, 15)
        Me.lblGageName.TabIndex = 82
        Me.lblGageName.Text = "Gage Name"
        '
        'cmbGageName
        '
        Me.cmbGageName.FormattingEnabled = True
        Me.cmbGageName.Location = New System.Drawing.Point(164, 399)
        Me.cmbGageName.Name = "cmbGageName"
        Me.cmbGageName.Size = New System.Drawing.Size(475, 23)
        Me.cmbGageName.TabIndex = 81
        '
        'txtBendRadius
        '
        Me.txtBendRadius.Location = New System.Drawing.Point(164, 469)
        Me.txtBendRadius.Name = "txtBendRadius"
        Me.txtBendRadius.Size = New System.Drawing.Size(475, 23)
        Me.txtBendRadius.TabIndex = 86
        '
        'txtBendType
        '
        Me.txtBendType.Location = New System.Drawing.Point(164, 431)
        Me.txtBendType.Name = "txtBendType"
        Me.txtBendType.Size = New System.Drawing.Size(475, 23)
        Me.txtBendType.TabIndex = 85
        '
        'lblBendRadius
        '
        Me.lblBendRadius.AutoSize = True
        Me.lblBendRadius.Location = New System.Drawing.Point(23, 469)
        Me.lblBendRadius.Name = "lblBendRadius"
        Me.lblBendRadius.Size = New System.Drawing.Size(72, 15)
        Me.lblBendRadius.TabIndex = 84
        Me.lblBendRadius.Text = "Bend Radius"
        '
        'lblBendType
        '
        Me.lblBendType.AutoSize = True
        Me.lblBendType.Location = New System.Drawing.Point(21, 434)
        Me.lblBendType.Name = "lblBendType"
        Me.lblBendType.Size = New System.Drawing.Size(61, 15)
        Me.lblBendType.TabIndex = 83
        Me.lblBendType.Text = "Bend Type"
        '
        'txtMaterialLibrary
        '
        Me.txtMaterialLibrary.Location = New System.Drawing.Point(164, 532)
        Me.txtMaterialLibrary.Name = "txtMaterialLibrary"
        Me.txtMaterialLibrary.Size = New System.Drawing.Size(475, 23)
        Me.txtMaterialLibrary.TabIndex = 88
        Me.txtMaterialLibrary.Text = "BEC MATERIAL LIBRARY"
        Me.txtMaterialLibrary.Visible = False
        '
        'lblMaterialLibrary
        '
        Me.lblMaterialLibrary.AutoSize = True
        Me.lblMaterialLibrary.Location = New System.Drawing.Point(21, 532)
        Me.lblMaterialLibrary.Name = "lblMaterialLibrary"
        Me.lblMaterialLibrary.Size = New System.Drawing.Size(67, 15)
        Me.lblMaterialLibrary.TabIndex = 87
        Me.lblMaterialLibrary.Text = "Mat Library"
        Me.lblMaterialLibrary.Visible = False
        '
        'txtMaterialSpec
        '
        Me.txtMaterialSpec.Location = New System.Drawing.Point(164, 254)
        Me.txtMaterialSpec.Name = "txtMaterialSpec"
        Me.txtMaterialSpec.Size = New System.Drawing.Size(475, 23)
        Me.txtMaterialSpec.TabIndex = 90
        '
        'lblMaterialSpec
        '
        Me.lblMaterialSpec.AutoSize = True
        Me.lblMaterialSpec.Location = New System.Drawing.Point(21, 258)
        Me.lblMaterialSpec.Name = "lblMaterialSpec"
        Me.lblMaterialSpec.Size = New System.Drawing.Size(81, 15)
        Me.lblMaterialSpec.TabIndex = 89
        Me.lblMaterialSpec.Text = "Material Spec."
        '
        'txtTemplate
        '
        Me.txtTemplate.Location = New System.Drawing.Point(164, 503)
        Me.txtTemplate.Name = "txtTemplate"
        Me.txtTemplate.Size = New System.Drawing.Size(475, 23)
        Me.txtTemplate.TabIndex = 92
        '
        'lblTemplate
        '
        Me.lblTemplate.AutoSize = True
        Me.lblTemplate.Location = New System.Drawing.Point(21, 506)
        Me.lblTemplate.Name = "lblTemplate"
        Me.lblTemplate.Size = New System.Drawing.Size(55, 15)
        Me.lblTemplate.TabIndex = 91
        Me.lblTemplate.Text = "Template"
        '
        'txtDiameter
        '
        Me.txtDiameter.Location = New System.Drawing.Point(164, 469)
        Me.txtDiameter.Name = "txtDiameter"
        Me.txtDiameter.Size = New System.Drawing.Size(475, 23)
        Me.txtDiameter.TabIndex = 94
        '
        'lblDiameter
        '
        Me.lblDiameter.AutoSize = True
        Me.lblDiameter.Location = New System.Drawing.Point(21, 469)
        Me.lblDiameter.Name = "lblDiameter"
        Me.lblDiameter.Size = New System.Drawing.Size(55, 15)
        Me.lblDiameter.TabIndex = 93
        Me.lblDiameter.Text = "Diameter"
        '
        'PanelBody
        '
        Me.PanelBody.Controls.Add(Me.Label1)
        Me.PanelBody.Controls.Add(Me.BtnBrowseSolidEdgePartsTemplateDir)
        Me.PanelBody.Controls.Add(Me.TxtSolidEdgePartsTemplateDirectory)
        Me.PanelBody.Controls.Add(Me.lblExcelPath)
        Me.PanelBody.Controls.Add(Me.btnBrowseExcel)
        Me.PanelBody.Controls.Add(Me.txtExcelPath)
        Me.PanelBody.Controls.Add(Me.btnGetData)
        Me.PanelBody.Controls.Add(Me.btnClose)
        Me.PanelBody.Controls.Add(Me.txtFileName)
        Me.PanelBody.Controls.Add(Me.lblFileName)
        Me.PanelBody.Controls.Add(Me.cmbCategory)
        Me.PanelBody.Controls.Add(Me.lblCategory)
        Me.PanelBody.Controls.Add(Me.cmbType)
        Me.PanelBody.Controls.Add(Me.lblType)
        Me.PanelBody.Controls.Add(Me.cmbMaterialUsed)
        Me.PanelBody.Controls.Add(Me.txtMaterialSpec)
        Me.PanelBody.Controls.Add(Me.lblTemplateName)
        Me.PanelBody.Controls.Add(Me.lblMaterialSpec)
        Me.PanelBody.Controls.Add(Me.txtMaterialLibrary)
        Me.PanelBody.Controls.Add(Me.Label2)
        Me.PanelBody.Controls.Add(Me.cmbBECMaterial)
        Me.PanelBody.Controls.Add(Me.lblBendRadius)
        Me.PanelBody.Controls.Add(Me.cmbGageName)
        Me.PanelBody.Controls.Add(Me.btnCreatePart)
        Me.PanelBody.Controls.Add(Me.txtBendRadius)
        Me.PanelBody.Controls.Add(Me.txtBendType)
        Me.PanelBody.Controls.Add(Me.txtTemplate)
        Me.PanelBody.Controls.Add(Me.Label3)
        Me.PanelBody.Controls.Add(Me.lblTemplate)
        Me.PanelBody.Controls.Add(Me.lblDiameter)
        Me.PanelBody.Controls.Add(Me.lblMaterialLibrary)
        Me.PanelBody.Controls.Add(Me.txtLinearLength)
        Me.PanelBody.Controls.Add(Me.txtWidth)
        Me.PanelBody.Controls.Add(Me.txtDiameter)
        Me.PanelBody.Controls.Add(Me.lblBendType)
        Me.PanelBody.Controls.Add(Me.lblWidth)
        Me.PanelBody.Controls.Add(Me.lblLinearLength)
        Me.PanelBody.Controls.Add(Me.lblGageName)
        Me.PanelBody.Controls.Add(Me.cmbGageTable)
        Me.PanelBody.Controls.Add(Me.lblGageTable)
        Me.PanelBody.Controls.Add(Me.lblHeight)
        Me.PanelBody.Controls.Add(Me.txtHeight)
        Me.PanelBody.Controls.Add(Me.cmbThickness)
        Me.PanelBody.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelBody.Location = New System.Drawing.Point(0, 0)
        Me.PanelBody.Name = "PanelBody"
        Me.PanelBody.Size = New System.Drawing.Size(1063, 591)
        Me.PanelBody.TabIndex = 95
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(21, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(190, 15)
        Me.Label1.TabIndex = 102
        Me.Label1.Text = "SolidEdge Parts Template Directory"
        '
        'BtnBrowseSolidEdgePartsTemplateDir
        '
        Me.BtnBrowseSolidEdgePartsTemplateDir.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnBrowseSolidEdgePartsTemplateDir.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnBrowseSolidEdgePartsTemplateDir.Location = New System.Drawing.Point(952, 3)
        Me.BtnBrowseSolidEdgePartsTemplateDir.Name = "BtnBrowseSolidEdgePartsTemplateDir"
        Me.BtnBrowseSolidEdgePartsTemplateDir.Size = New System.Drawing.Size(100, 30)
        Me.BtnBrowseSolidEdgePartsTemplateDir.TabIndex = 104
        Me.BtnBrowseSolidEdgePartsTemplateDir.Text = "Browse"
        Me.BtnBrowseSolidEdgePartsTemplateDir.UseVisualStyleBackColor = True
        Me.BtnBrowseSolidEdgePartsTemplateDir.Visible = False
        '
        'TxtSolidEdgePartsTemplateDirectory
        '
        Me.TxtSolidEdgePartsTemplateDirectory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSolidEdgePartsTemplateDirectory.Enabled = False
        Me.TxtSolidEdgePartsTemplateDirectory.Location = New System.Drawing.Point(212, 7)
        Me.TxtSolidEdgePartsTemplateDirectory.Name = "TxtSolidEdgePartsTemplateDirectory"
        Me.TxtSolidEdgePartsTemplateDirectory.Size = New System.Drawing.Size(733, 23)
        Me.TxtSolidEdgePartsTemplateDirectory.TabIndex = 103
        '
        'lblExcelPath
        '
        Me.lblExcelPath.AutoSize = True
        Me.lblExcelPath.Location = New System.Drawing.Point(21, 43)
        Me.lblExcelPath.Name = "lblExcelPath"
        Me.lblExcelPath.Size = New System.Drawing.Size(131, 15)
        Me.lblExcelPath.TabIndex = 99
        Me.lblExcelPath.Text = "BEC Material Excel Path"
        '
        'btnBrowseExcel
        '
        Me.btnBrowseExcel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrowseExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnBrowseExcel.Location = New System.Drawing.Point(952, 35)
        Me.btnBrowseExcel.Name = "btnBrowseExcel"
        Me.btnBrowseExcel.Size = New System.Drawing.Size(100, 30)
        Me.btnBrowseExcel.TabIndex = 101
        Me.btnBrowseExcel.Text = "Browse"
        Me.btnBrowseExcel.UseVisualStyleBackColor = True
        Me.btnBrowseExcel.Visible = False
        '
        'txtExcelPath
        '
        Me.txtExcelPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtExcelPath.Enabled = False
        Me.txtExcelPath.Location = New System.Drawing.Point(212, 39)
        Me.txtExcelPath.Name = "txtExcelPath"
        Me.txtExcelPath.Size = New System.Drawing.Size(733, 23)
        Me.txtExcelPath.TabIndex = 100
        '
        'btnGetData
        '
        Me.btnGetData.Location = New System.Drawing.Point(21, 70)
        Me.btnGetData.Name = "btnGetData"
        Me.btnGetData.Size = New System.Drawing.Size(353, 30)
        Me.btnGetData.TabIndex = 98
        Me.btnGetData.Text = "Get BEC Material Data"
        Me.btnGetData.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(845, 549)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(100, 30)
        Me.btnClose.TabIndex = 97
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'txtFileName
        '
        Me.txtFileName.Location = New System.Drawing.Point(164, 111)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.Size = New System.Drawing.Size(475, 23)
        Me.txtFileName.TabIndex = 96
        '
        'lblFileName
        '
        Me.lblFileName.AutoSize = True
        Me.lblFileName.Location = New System.Drawing.Point(21, 115)
        Me.lblFileName.Name = "lblFileName"
        Me.lblFileName.Size = New System.Drawing.Size(60, 15)
        Me.lblFileName.TabIndex = 95
        Me.lblFileName.Text = "File Name"
        '
        'txtLinearLength
        '
        Me.txtLinearLength.Location = New System.Drawing.Point(164, 399)
        Me.txtLinearLength.Name = "txtLinearLength"
        Me.txtLinearLength.Size = New System.Drawing.Size(475, 23)
        Me.txtLinearLength.TabIndex = 105
        '
        'lblLinearLength
        '
        Me.lblLinearLength.AutoSize = True
        Me.lblLinearLength.Location = New System.Drawing.Point(21, 402)
        Me.lblLinearLength.Name = "lblLinearLength"
        Me.lblLinearLength.Size = New System.Drawing.Size(79, 15)
        Me.lblLinearLength.TabIndex = 106
        Me.lblLinearLength.Text = "Linear Length"
        '
        'CreateNewPartForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1063, 591)
        Me.Controls.Add(Me.PanelBody)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "CreateNewPartForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "New Part"
        Me.PanelBody.ResumeLayout(False)
        Me.PanelBody.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Label3 As Label
    Friend WithEvents cmbBECMaterial As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents cmbThickness As ComboBox
    Friend WithEvents lblTemplateName As Label
    Friend WithEvents cmbMaterialUsed As ComboBox
    Friend WithEvents lblType As Label
    Friend WithEvents cmbType As ComboBox
    Friend WithEvents lblCategory As Label
    Friend WithEvents cmbCategory As ComboBox
    Friend WithEvents lblHeight As Label
    Friend WithEvents lblWidth As Label
    Friend WithEvents txtHeight As TextBox
    Friend WithEvents txtWidth As TextBox
    Friend WithEvents btnCreatePart As Button
    Friend WithEvents lblGageTable As Label
    Friend WithEvents cmbGageTable As ComboBox
    Friend WithEvents lblGageName As Label
    Friend WithEvents cmbGageName As ComboBox
    Friend WithEvents txtBendRadius As TextBox
    Friend WithEvents txtBendType As TextBox
    Friend WithEvents lblBendRadius As Label
    Friend WithEvents lblBendType As Label
    Friend WithEvents txtMaterialLibrary As TextBox
    Friend WithEvents lblMaterialLibrary As Label
    Friend WithEvents txtMaterialSpec As TextBox
    Friend WithEvents lblMaterialSpec As Label
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents txtTemplate As TextBox
    Friend WithEvents lblTemplate As Label
    Friend WithEvents txtDiameter As TextBox
    Friend WithEvents lblDiameter As Label
    Friend WithEvents PanelBody As Panel
    Friend WithEvents txtFileName As TextBox
    Friend WithEvents lblFileName As Label
    Friend WithEvents btnClose As Button
    Friend WithEvents btnGetData As Button
    Friend WithEvents lblExcelPath As Label
    Friend WithEvents btnBrowseExcel As Button
    Friend WithEvents txtExcelPath As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents BtnBrowseSolidEdgePartsTemplateDir As Button
    Friend WithEvents TxtSolidEdgePartsTemplateDirectory As TextBox
    Friend WithEvents lblLinearLength As Label
    Friend WithEvents txtLinearLength As TextBox
End Class
