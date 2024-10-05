<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NewPartForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(NewPartForm))
        Me.lblTemplateFile = New System.Windows.Forms.Label()
        Me.txtExcelPath = New System.Windows.Forms.TextBox()
        Me.btnBrowseExcel = New System.Windows.Forms.Button()
        Me.btnTemplateLocation = New System.Windows.Forms.Button()
        Me.txtTemplateDirectoryPath = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblCategory = New System.Windows.Forms.Label()
        Me.cmbCategory = New System.Windows.Forms.ComboBox()
        Me.lblType = New System.Windows.Forms.Label()
        Me.cmbType = New System.Windows.Forms.ComboBox()
        Me.lblTemplateName = New System.Windows.Forms.Label()
        Me.cmbTemplate = New System.Windows.Forms.ComboBox()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnCreateDocument = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmbThickness = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmbBECMaterial = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmbOriginalTempalte = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'lblTemplateFile
        '
        Me.lblTemplateFile.AutoSize = True
        Me.lblTemplateFile.Location = New System.Drawing.Point(8, 22)
        Me.lblTemplateFile.Name = "lblTemplateFile"
        Me.lblTemplateFile.Size = New System.Drawing.Size(76, 15)
        Me.lblTemplateFile.TabIndex = 0
        Me.lblTemplateFile.Text = "Template File"
        '
        'txtExcelPath
        '
        Me.txtExcelPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtExcelPath.Location = New System.Drawing.Point(149, 18)
        Me.txtExcelPath.Name = "txtExcelPath"
        Me.txtExcelPath.Size = New System.Drawing.Size(796, 23)
        Me.txtExcelPath.TabIndex = 1
        '
        'btnBrowseExcel
        '
        Me.btnBrowseExcel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrowseExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnBrowseExcel.Location = New System.Drawing.Point(951, 14)
        Me.btnBrowseExcel.Name = "btnBrowseExcel"
        Me.btnBrowseExcel.Size = New System.Drawing.Size(100, 30)
        Me.btnBrowseExcel.TabIndex = 18
        Me.btnBrowseExcel.Text = "Browse"
        Me.btnBrowseExcel.UseVisualStyleBackColor = True
        '
        'btnTemplateLocation
        '
        Me.btnTemplateLocation.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnTemplateLocation.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnTemplateLocation.Location = New System.Drawing.Point(951, 46)
        Me.btnTemplateLocation.Name = "btnTemplateLocation"
        Me.btnTemplateLocation.Size = New System.Drawing.Size(100, 30)
        Me.btnTemplateLocation.TabIndex = 21
        Me.btnTemplateLocation.Text = "Browse"
        Me.btnTemplateLocation.UseVisualStyleBackColor = True
        '
        'txtTemplateDirectoryPath
        '
        Me.txtTemplateDirectoryPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTemplateDirectoryPath.Location = New System.Drawing.Point(149, 50)
        Me.txtTemplateDirectoryPath.Name = "txtTemplateDirectoryPath"
        Me.txtTemplateDirectoryPath.Size = New System.Drawing.Size(796, 23)
        Me.txtTemplateDirectoryPath.TabIndex = 20
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 54)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 15)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "Template Location"
        '
        'lblCategory
        '
        Me.lblCategory.AutoSize = True
        Me.lblCategory.Location = New System.Drawing.Point(8, 88)
        Me.lblCategory.Name = "lblCategory"
        Me.lblCategory.Size = New System.Drawing.Size(55, 15)
        Me.lblCategory.TabIndex = 53
        Me.lblCategory.Text = "Category"
        '
        'cmbCategory
        '
        Me.cmbCategory.FormattingEnabled = True
        Me.cmbCategory.Location = New System.Drawing.Point(149, 84)
        Me.cmbCategory.Name = "cmbCategory"
        Me.cmbCategory.Size = New System.Drawing.Size(181, 23)
        Me.cmbCategory.TabIndex = 52
        '
        'lblType
        '
        Me.lblType.AutoSize = True
        Me.lblType.Location = New System.Drawing.Point(8, 119)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(31, 15)
        Me.lblType.TabIndex = 55
        Me.lblType.Text = "Type"
        '
        'cmbType
        '
        Me.cmbType.FormattingEnabled = True
        Me.cmbType.Location = New System.Drawing.Point(149, 115)
        Me.cmbType.Name = "cmbType"
        Me.cmbType.Size = New System.Drawing.Size(181, 23)
        Me.cmbType.TabIndex = 54
        '
        'lblTemplateName
        '
        Me.lblTemplateName.AutoSize = True
        Me.lblTemplateName.Location = New System.Drawing.Point(8, 150)
        Me.lblTemplateName.Name = "lblTemplateName"
        Me.lblTemplateName.Size = New System.Drawing.Size(136, 15)
        Me.lblTemplateName.TabIndex = 57
        Me.lblTemplateName.Text = "BEC Code/Material Used"
        '
        'cmbTemplate
        '
        Me.cmbTemplate.FormattingEnabled = True
        Me.cmbTemplate.Location = New System.Drawing.Point(149, 146)
        Me.cmbTemplate.Name = "cmbTemplate"
        Me.cmbTemplate.Size = New System.Drawing.Size(181, 23)
        Me.cmbTemplate.TabIndex = 56
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnClose.Location = New System.Drawing.Point(951, 549)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(100, 30)
        Me.btnClose.TabIndex = 58
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnCreateDocument
        '
        Me.btnCreateDocument.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCreateDocument.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCreateDocument.Location = New System.Drawing.Point(800, 551)
        Me.btnCreateDocument.Name = "btnCreateDocument"
        Me.btnCreateDocument.Size = New System.Drawing.Size(145, 30)
        Me.btnCreateDocument.TabIndex = 59
        Me.btnCreateDocument.Text = "Create Document"
        Me.btnCreateDocument.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(342, 120)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(58, 15)
        Me.Label2.TabIndex = 61
        Me.Label2.Text = "Thickness"
        '
        'cmbThickness
        '
        Me.cmbThickness.FormattingEnabled = True
        Me.cmbThickness.Location = New System.Drawing.Point(429, 117)
        Me.cmbThickness.Name = "cmbThickness"
        Me.cmbThickness.Size = New System.Drawing.Size(181, 23)
        Me.cmbThickness.TabIndex = 60
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(342, 149)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(74, 15)
        Me.Label3.TabIndex = 63
        Me.Label3.Text = "BEC Material"
        '
        'cmbBECMaterial
        '
        Me.cmbBECMaterial.FormattingEnabled = True
        Me.cmbBECMaterial.Location = New System.Drawing.Point(429, 146)
        Me.cmbBECMaterial.Name = "cmbBECMaterial"
        Me.cmbBECMaterial.Size = New System.Drawing.Size(181, 23)
        Me.cmbBECMaterial.TabIndex = 62
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 230)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(135, 15)
        Me.Label4.TabIndex = 65
        Me.Label4.Text = "Original Template Name"
        Me.Label4.Visible = False
        '
        'cmbOriginalTempalte
        '
        Me.cmbOriginalTempalte.FormattingEnabled = True
        Me.cmbOriginalTempalte.Location = New System.Drawing.Point(149, 226)
        Me.cmbOriginalTempalte.Name = "cmbOriginalTempalte"
        Me.cmbOriginalTempalte.Size = New System.Drawing.Size(181, 23)
        Me.cmbOriginalTempalte.TabIndex = 64
        Me.cmbOriginalTempalte.Visible = False
        '
        'NewPartForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1063, 591)
        Me.Controls.Add(Me.cmbOriginalTempalte)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cmbBECMaterial)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cmbThickness)
        Me.Controls.Add(Me.btnCreateDocument)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.lblTemplateName)
        Me.Controls.Add(Me.cmbTemplate)
        Me.Controls.Add(Me.lblType)
        Me.Controls.Add(Me.cmbType)
        Me.Controls.Add(Me.lblCategory)
        Me.Controls.Add(Me.cmbCategory)
        Me.Controls.Add(Me.btnTemplateLocation)
        Me.Controls.Add(Me.txtTemplateDirectoryPath)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnBrowseExcel)
        Me.Controls.Add(Me.txtExcelPath)
        Me.Controls.Add(Me.lblTemplateFile)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "NewPartForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "New Part 1.0.25"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblTemplateFile As Label
    Friend WithEvents txtExcelPath As TextBox
    Friend WithEvents btnBrowseExcel As Button
    Friend WithEvents btnTemplateLocation As Button
    Friend WithEvents txtTemplateDirectoryPath As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents lblCategory As Label
    Friend WithEvents cmbCategory As ComboBox
    Friend WithEvents lblType As Label
    Friend WithEvents cmbType As ComboBox
    Friend WithEvents lblTemplateName As Label
    Friend WithEvents cmbTemplate As ComboBox
    Friend WithEvents btnClose As Button
    Friend WithEvents btnCreateDocument As Button
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents Label2 As Label
    Friend WithEvents cmbThickness As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents cmbBECMaterial As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents cmbOriginalTempalte As ComboBox
End Class
