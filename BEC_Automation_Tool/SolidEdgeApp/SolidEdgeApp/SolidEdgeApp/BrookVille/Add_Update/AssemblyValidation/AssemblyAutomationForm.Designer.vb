<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AssemblyAutomationForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AssemblyAutomationForm))
        Me.lblAssemblyName = New System.Windows.Forms.Label()
        Me.lblAssemblyPath = New System.Windows.Forms.Label()
        Me.dgvDocumentDetails = New System.Windows.Forms.DataGridView()
        Me.btnGetData = New System.Windows.Forms.Button()
        Me.btnOpenDocument = New System.Windows.Forms.Button()
        Me.btnApply = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.cmbMaterialUsed2_Mw = New System.Windows.Forms.ComboBox()
        Me.txtMaterialUsed_C = New System.Windows.Forms.TextBox()
        Me.cmbCategory = New System.Windows.Forms.ComboBox()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.txtSize2_Mw = New System.Windows.Forms.TextBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.txtGrade2_Mw = New System.Windows.Forms.TextBox()
        Me.txtGageName_Mw = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.txtThickness2_Mw = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.txtPartType_Mw = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtMaterialSpec2_Mw = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txtBECMaterial2_Mw = New System.Windows.Forms.TextBox()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.txtBendRadius_Mw = New System.Windows.Forms.TextBox()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.txtSize_C = New System.Windows.Forms.TextBox()
        Me.txtGrade_C = New System.Windows.Forms.TextBox()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.txtGageName_C = New System.Windows.Forms.TextBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.txtThickness_C = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.txtPartType_C = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.txtMaterialSpec_C = New System.Windows.Forms.TextBox()
        Me.txtCurrentMaterial_C = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.txtBendRadius_C = New System.Windows.Forms.TextBox()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.txtGageTable = New System.Windows.Forms.TextBox()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.cmbBendType_Mw = New System.Windows.Forms.ComboBox()
        Me.lblExcelPath = New System.Windows.Forms.Label()
        Me.btnBrowseExcel = New System.Windows.Forms.Button()
        Me.txtExcelPath = New System.Windows.Forms.TextBox()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.ComboBoxFields = New System.Windows.Forms.ComboBox()
        Me.txtSearch = New System.Windows.Forms.TextBox()
        Me.btnSearchFile = New System.Windows.Forms.Button()
        Me.MoButton1 = New MoCustomControls.MoButton()
        Me.MoCloseButton1 = New MoCustomControls.MoCloseButton()
        Me.Panel1 = New System.Windows.Forms.Panel()
        CType(Me.dgvDocumentDetails, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblAssemblyName
        '
        Me.lblAssemblyName.AutoSize = True
        Me.lblAssemblyName.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAssemblyName.Location = New System.Drawing.Point(211, 55)
        Me.lblAssemblyName.Name = "lblAssemblyName"
        Me.lblAssemblyName.Size = New System.Drawing.Size(102, 17)
        Me.lblAssemblyName.TabIndex = 0
        Me.lblAssemblyName.Text = "Assembly Name"
        '
        'lblAssemblyPath
        '
        Me.lblAssemblyPath.AutoSize = True
        Me.lblAssemblyPath.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAssemblyPath.Location = New System.Drawing.Point(12, 86)
        Me.lblAssemblyPath.Name = "lblAssemblyPath"
        Me.lblAssemblyPath.Size = New System.Drawing.Size(92, 17)
        Me.lblAssemblyPath.TabIndex = 1
        Me.lblAssemblyPath.Text = "Assembly Path"
        '
        'dgvDocumentDetails
        '
        Me.dgvDocumentDetails.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDocumentDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvDocumentDetails.Location = New System.Drawing.Point(0, 0)
        Me.dgvDocumentDetails.Name = "dgvDocumentDetails"
        Me.dgvDocumentDetails.Size = New System.Drawing.Size(494, 400)
        Me.dgvDocumentDetails.TabIndex = 2
        '
        'btnGetData
        '
        Me.btnGetData.Location = New System.Drawing.Point(12, 49)
        Me.btnGetData.Name = "btnGetData"
        Me.btnGetData.Size = New System.Drawing.Size(190, 30)
        Me.btnGetData.TabIndex = 3
        Me.btnGetData.Text = "Get Current Assembly Data"
        Me.btnGetData.UseVisualStyleBackColor = True
        '
        'btnOpenDocument
        '
        Me.btnOpenDocument.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnOpenDocument.Location = New System.Drawing.Point(13, 553)
        Me.btnOpenDocument.Name = "btnOpenDocument"
        Me.btnOpenDocument.Size = New System.Drawing.Size(143, 30)
        Me.btnOpenDocument.TabIndex = 4
        Me.btnOpenDocument.Text = "Open Document"
        Me.btnOpenDocument.UseVisualStyleBackColor = True
        Me.btnOpenDocument.Visible = False
        '
        'btnApply
        '
        Me.btnApply.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnApply.Location = New System.Drawing.Point(949, 553)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.Size = New System.Drawing.Size(100, 30)
        Me.btnApply.TabIndex = 5
        Me.btnApply.Text = "Apply"
        Me.btnApply.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(844, 553)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(100, 30)
        Me.btnClose.TabIndex = 6
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(13, 112)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.dgvDocumentDetails)
        Me.SplitContainer1.Panel1.RightToLeft = System.Windows.Forms.RightToLeft.No
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TableLayoutPanel2)
        Me.SplitContainer1.Panel2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SplitContainer1.Panel2MinSize = 470
        Me.SplitContainer1.Size = New System.Drawing.Size(1036, 400)
        Me.SplitContainer1.SplitterDistance = 494
        Me.SplitContainer1.TabIndex = 7
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.AllowDrop = True
        Me.TableLayoutPanel2.ColumnCount = 4
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 18.40149!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 31.41264!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 17.84387!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 32.34201!))
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 15
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(538, 400)
        Me.TableLayoutPanel2.TabIndex = 113
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(274, 40)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(79, 15)
        Me.Label27.TabIndex = 103
        Me.Label27.Text = "Material Used"
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Location = New System.Drawing.Point(-1, 38)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(79, 15)
        Me.Label32.TabIndex = 75
        Me.Label32.Text = "Material Used"
        '
        'Label33
        '
        Me.Label33.AutoSize = True
        Me.Label33.Location = New System.Drawing.Point(47, 0)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(84, 15)
        Me.Label33.TabIndex = 98
        Me.Label33.Text = "Current details"
        '
        'Label34
        '
        Me.Label34.AutoSize = True
        Me.Label34.Location = New System.Drawing.Point(323, 0)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(115, 15)
        Me.Label34.TabIndex = 99
        Me.Label34.Text = "Material-wise details"
        '
        'Label35
        '
        Me.Label35.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label35.AutoSize = True
        Me.Label35.Location = New System.Drawing.Point(785, 65)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(55, 15)
        Me.Label35.TabIndex = 53
        Me.Label35.Text = "Category"
        '
        'cmbMaterialUsed2_Mw
        '
        Me.cmbMaterialUsed2_Mw.FormattingEnabled = True
        Me.cmbMaterialUsed2_Mw.Location = New System.Drawing.Point(362, 37)
        Me.cmbMaterialUsed2_Mw.Name = "cmbMaterialUsed2_Mw"
        Me.cmbMaterialUsed2_Mw.Size = New System.Drawing.Size(168, 23)
        Me.cmbMaterialUsed2_Mw.TabIndex = 68
        '
        'txtMaterialUsed_C
        '
        Me.txtMaterialUsed_C.Location = New System.Drawing.Point(89, 37)
        Me.txtMaterialUsed_C.Name = "txtMaterialUsed_C"
        Me.txtMaterialUsed_C.Size = New System.Drawing.Size(153, 23)
        Me.txtMaterialUsed_C.TabIndex = 13
        '
        'cmbCategory
        '
        Me.cmbCategory.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.cmbCategory.FormattingEnabled = True
        Me.cmbCategory.Location = New System.Drawing.Point(873, 62)
        Me.cmbCategory.Name = "cmbCategory"
        Me.cmbCategory.Size = New System.Drawing.Size(168, 23)
        Me.cmbCategory.TabIndex = 52
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(274, 75)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(27, 15)
        Me.Label26.TabIndex = 105
        Me.Label26.Text = "Size"
        '
        'txtSize2_Mw
        '
        Me.txtSize2_Mw.Location = New System.Drawing.Point(362, 72)
        Me.txtSize2_Mw.Name = "txtSize2_Mw"
        Me.txtSize2_Mw.Size = New System.Drawing.Size(168, 23)
        Me.txtSize2_Mw.TabIndex = 69
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(274, 110)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(38, 15)
        Me.Label23.TabIndex = 106
        Me.Label23.Text = "Grade"
        '
        'txtGrade2_Mw
        '
        Me.txtGrade2_Mw.Location = New System.Drawing.Point(362, 107)
        Me.txtGrade2_Mw.Name = "txtGrade2_Mw"
        Me.txtGrade2_Mw.Size = New System.Drawing.Size(168, 23)
        Me.txtGrade2_Mw.TabIndex = 71
        '
        'txtGageName_Mw
        '
        Me.txtGageName_Mw.Location = New System.Drawing.Point(362, 142)
        Me.txtGageName_Mw.Name = "txtGageName_Mw"
        Me.txtGageName_Mw.Size = New System.Drawing.Size(168, 23)
        Me.txtGageName_Mw.TabIndex = 83
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(274, 145)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(69, 15)
        Me.Label20.TabIndex = 107
        Me.Label20.Text = "Gage Name"
        '
        'txtThickness2_Mw
        '
        Me.txtThickness2_Mw.Location = New System.Drawing.Point(362, 177)
        Me.txtThickness2_Mw.Name = "txtThickness2_Mw"
        Me.txtThickness2_Mw.Size = New System.Drawing.Size(168, 23)
        Me.txtThickness2_Mw.TabIndex = 70
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(274, 180)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(89, 15)
        Me.Label18.TabIndex = 108
        Me.Label18.Text = "Thickness(inch)"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(274, 215)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(55, 15)
        Me.Label16.TabIndex = 109
        Me.Label16.Text = "Part Type"
        '
        'txtPartType_Mw
        '
        Me.txtPartType_Mw.Location = New System.Drawing.Point(362, 212)
        Me.txtPartType_Mw.Name = "txtPartType_Mw"
        Me.txtPartType_Mw.Size = New System.Drawing.Size(168, 23)
        Me.txtPartType_Mw.TabIndex = 72
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(274, 250)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(78, 15)
        Me.Label15.TabIndex = 110
        Me.Label15.Text = "Material Spec"
        '
        'txtMaterialSpec2_Mw
        '
        Me.txtMaterialSpec2_Mw.Location = New System.Drawing.Point(361, 247)
        Me.txtMaterialSpec2_Mw.Name = "txtMaterialSpec2_Mw"
        Me.txtMaterialSpec2_Mw.Size = New System.Drawing.Size(168, 23)
        Me.txtMaterialSpec2_Mw.TabIndex = 73
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(274, 285)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(74, 15)
        Me.Label13.TabIndex = 111
        Me.Label13.Text = "BEC Material"
        '
        'txtBECMaterial2_Mw
        '
        Me.txtBECMaterial2_Mw.Location = New System.Drawing.Point(362, 282)
        Me.txtBECMaterial2_Mw.Name = "txtBECMaterial2_Mw"
        Me.txtBECMaterial2_Mw.Size = New System.Drawing.Size(168, 23)
        Me.txtBECMaterial2_Mw.TabIndex = 74
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Location = New System.Drawing.Point(274, 320)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(72, 15)
        Me.Label29.TabIndex = 104
        Me.Label29.Text = "Bend Radius"
        '
        'txtBendRadius_Mw
        '
        Me.txtBendRadius_Mw.Location = New System.Drawing.Point(361, 317)
        Me.txtBendRadius_Mw.Name = "txtBendRadius_Mw"
        Me.txtBendRadius_Mw.Size = New System.Drawing.Size(168, 23)
        Me.txtBendRadius_Mw.TabIndex = 84
        '
        'Label31
        '
        Me.Label31.AutoSize = True
        Me.Label31.Location = New System.Drawing.Point(-1, 75)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(27, 15)
        Me.Label31.TabIndex = 76
        Me.Label31.Text = "Size"
        '
        'txtSize_C
        '
        Me.txtSize_C.Location = New System.Drawing.Point(89, 73)
        Me.txtSize_C.Name = "txtSize_C"
        Me.txtSize_C.Size = New System.Drawing.Size(153, 23)
        Me.txtSize_C.TabIndex = 88
        '
        'txtGrade_C
        '
        Me.txtGrade_C.Location = New System.Drawing.Point(89, 109)
        Me.txtGrade_C.Name = "txtGrade_C"
        Me.txtGrade_C.Size = New System.Drawing.Size(153, 23)
        Me.txtGrade_C.TabIndex = 90
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(-1, 110)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(38, 15)
        Me.Label28.TabIndex = 78
        Me.Label28.Text = "Grade"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(-1, 146)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(69, 15)
        Me.Label21.TabIndex = 82
        Me.Label21.Text = "Gage Name"
        '
        'txtGageName_C
        '
        Me.txtGageName_C.Location = New System.Drawing.Point(89, 145)
        Me.txtGageName_C.Name = "txtGageName_C"
        Me.txtGageName_C.Size = New System.Drawing.Size(153, 23)
        Me.txtGageName_C.TabIndex = 93
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(-1, 182)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(89, 15)
        Me.Label22.TabIndex = 77
        Me.Label22.Text = "Thickness(inch)"
        '
        'txtThickness_C
        '
        Me.txtThickness_C.Location = New System.Drawing.Point(89, 181)
        Me.txtThickness_C.Name = "txtThickness_C"
        Me.txtThickness_C.Size = New System.Drawing.Size(153, 23)
        Me.txtThickness_C.TabIndex = 89
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(-1, 218)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(55, 15)
        Me.Label19.TabIndex = 79
        Me.Label19.Text = "Part Type"
        '
        'txtPartType_C
        '
        Me.txtPartType_C.Location = New System.Drawing.Point(89, 217)
        Me.txtPartType_C.Name = "txtPartType_C"
        Me.txtPartType_C.Size = New System.Drawing.Size(153, 23)
        Me.txtPartType_C.TabIndex = 1
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(-1, 254)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(78, 15)
        Me.Label17.TabIndex = 80
        Me.Label17.Text = "Material Spec"
        '
        'txtMaterialSpec_C
        '
        Me.txtMaterialSpec_C.Location = New System.Drawing.Point(89, 253)
        Me.txtMaterialSpec_C.Name = "txtMaterialSpec_C"
        Me.txtMaterialSpec_C.Size = New System.Drawing.Size(153, 23)
        Me.txtMaterialSpec_C.TabIndex = 91
        '
        'txtCurrentMaterial_C
        '
        Me.txtCurrentMaterial_C.Location = New System.Drawing.Point(89, 289)
        Me.txtCurrentMaterial_C.Name = "txtCurrentMaterial_C"
        Me.txtCurrentMaterial_C.Size = New System.Drawing.Size(153, 23)
        Me.txtCurrentMaterial_C.TabIndex = 97
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(-1, 290)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(74, 15)
        Me.Label14.TabIndex = 81
        Me.Label14.Text = "BEC Material"
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(-1, 326)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(72, 15)
        Me.Label24.TabIndex = 85
        Me.Label24.Text = "Bend Radius"
        '
        'txtBendRadius_C
        '
        Me.txtBendRadius_C.Location = New System.Drawing.Point(89, 325)
        Me.txtBendRadius_C.Name = "txtBendRadius_C"
        Me.txtBendRadius_C.Size = New System.Drawing.Size(153, 23)
        Me.txtBendRadius_C.TabIndex = 94
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(274, 390)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(64, 15)
        Me.Label25.TabIndex = 100
        Me.Label25.Text = "Gage Table"
        '
        'txtGageTable
        '
        Me.txtGageTable.Location = New System.Drawing.Point(362, 387)
        Me.txtGageTable.Name = "txtGageTable"
        Me.txtGageTable.Size = New System.Drawing.Size(168, 23)
        Me.txtGageTable.TabIndex = 86
        '
        'Label30
        '
        Me.Label30.AutoSize = True
        Me.Label30.Location = New System.Drawing.Point(274, 355)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(61, 15)
        Me.Label30.TabIndex = 100
        Me.Label30.Text = "Bend Type"
        '
        'cmbBendType_Mw
        '
        Me.cmbBendType_Mw.FormattingEnabled = True
        Me.cmbBendType_Mw.Location = New System.Drawing.Point(361, 352)
        Me.cmbBendType_Mw.Name = "cmbBendType_Mw"
        Me.cmbBendType_Mw.Size = New System.Drawing.Size(168, 23)
        Me.cmbBendType_Mw.TabIndex = 112
        '
        'lblExcelPath
        '
        Me.lblExcelPath.AutoSize = True
        Me.lblExcelPath.Location = New System.Drawing.Point(12, 17)
        Me.lblExcelPath.Name = "lblExcelPath"
        Me.lblExcelPath.Size = New System.Drawing.Size(131, 15)
        Me.lblExcelPath.TabIndex = 18
        Me.lblExcelPath.Text = "BEC Material Excel Path"
        '
        'btnBrowseExcel
        '
        Me.btnBrowseExcel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrowseExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnBrowseExcel.Location = New System.Drawing.Point(949, 11)
        Me.btnBrowseExcel.Name = "btnBrowseExcel"
        Me.btnBrowseExcel.Size = New System.Drawing.Size(100, 30)
        Me.btnBrowseExcel.TabIndex = 20
        Me.btnBrowseExcel.Text = "Browse"
        Me.btnBrowseExcel.UseVisualStyleBackColor = True
        Me.btnBrowseExcel.Visible = False
        '
        'txtExcelPath
        '
        Me.txtExcelPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtExcelPath.Enabled = False
        Me.txtExcelPath.Location = New System.Drawing.Point(150, 15)
        Me.txtExcelPath.Name = "txtExcelPath"
        Me.txtExcelPath.Size = New System.Drawing.Size(793, 23)
        Me.txtExcelPath.TabIndex = 19
        '
        'btnRefresh
        '
        Me.btnRefresh.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnRefresh.Location = New System.Drawing.Point(738, 553)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(100, 30)
        Me.btnRefresh.TabIndex = 21
        Me.btnRefresh.Text = "Refresh"
        Me.btnRefresh.UseVisualStyleBackColor = True
        '
        'ComboBoxFields
        '
        Me.ComboBoxFields.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ComboBoxFields.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxFields.FormattingEnabled = True
        Me.ComboBoxFields.Location = New System.Drawing.Point(13, 518)
        Me.ComboBoxFields.Name = "ComboBoxFields"
        Me.ComboBoxFields.Size = New System.Drawing.Size(140, 23)
        Me.ComboBoxFields.TabIndex = 34
        '
        'txtSearch
        '
        Me.txtSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtSearch.Location = New System.Drawing.Point(159, 518)
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Size = New System.Drawing.Size(191, 23)
        Me.txtSearch.TabIndex = 32
        '
        'btnSearchFile
        '
        Me.btnSearchFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSearchFile.Image = Global.SolidEdgeApp.My.Resources.Resources.search_16px
        Me.btnSearchFile.Location = New System.Drawing.Point(356, 516)
        Me.btnSearchFile.Name = "btnSearchFile"
        Me.btnSearchFile.Size = New System.Drawing.Size(43, 26)
        Me.btnSearchFile.TabIndex = 33
        Me.btnSearchFile.UseVisualStyleBackColor = True
        '
        'MoButton1
        '
        Me.MoButton1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.MoButton1.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(113, Byte), Integer), CType(CType(227, Byte), Integer))
        Me.MoButton1.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(113, Byte), Integer), CType(CType(227, Byte), Integer))
        Me.MoButton1.BorderColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(113, Byte), Integer), CType(CType(227, Byte), Integer))
        Me.MoButton1.BorderRadius = 3
        Me.MoButton1.BorderSize = 1
        Me.MoButton1.FlatAppearance.BorderSize = 0
        Me.MoButton1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.MoButton1.ForeColor = System.Drawing.Color.White
        Me.MoButton1.Location = New System.Drawing.Point(470, 558)
        Me.MoButton1.Name = "MoButton1"
        Me.MoButton1.Size = New System.Drawing.Size(143, 30)
        Me.MoButton1.TabIndex = 35
        Me.MoButton1.Text = "Create report"
        Me.MoButton1.TextColor = System.Drawing.Color.White
        Me.MoButton1.UseVisualStyleBackColor = False
        Me.MoButton1.Visible = False
        '
        'MoCloseButton1
        '
        Me.MoCloseButton1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.MoCloseButton1.BackColor = System.Drawing.Color.White
        Me.MoCloseButton1.BackgroundColor = System.Drawing.Color.White
        Me.MoCloseButton1.BorderColor = System.Drawing.Color.FromArgb(CType(CType(194, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(75, Byte), Integer))
        Me.MoCloseButton1.BorderRadius = 3
        Me.MoCloseButton1.BorderSize = 1
        Me.MoCloseButton1.FlatAppearance.BorderSize = 0
        Me.MoCloseButton1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.MoCloseButton1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(194, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(75, Byte), Integer))
        Me.MoCloseButton1.Location = New System.Drawing.Point(387, 558)
        Me.MoCloseButton1.Name = "MoCloseButton1"
        Me.MoCloseButton1.Size = New System.Drawing.Size(77, 30)
        Me.MoCloseButton1.TabIndex = 36
        Me.MoCloseButton1.Text = "Close"
        Me.MoCloseButton1.TextColor = System.Drawing.Color.FromArgb(CType(CType(194, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(75, Byte), Integer))
        Me.MoCloseButton1.UseVisualStyleBackColor = False
        Me.MoCloseButton1.Visible = False
        '
        'Panel1
        '
        Me.Panel1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Panel1.AutoSize = True
        Me.Panel1.Controls.Add(Me.txtGageTable)
        Me.Panel1.Controls.Add(Me.Label25)
        Me.Panel1.Controls.Add(Me.txtBendRadius_Mw)
        Me.Panel1.Controls.Add(Me.Label29)
        Me.Panel1.Controls.Add(Me.cmbBendType_Mw)
        Me.Panel1.Controls.Add(Me.Label30)
        Me.Panel1.Controls.Add(Me.txtBECMaterial2_Mw)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.txtMaterialSpec2_Mw)
        Me.Panel1.Controls.Add(Me.Label15)
        Me.Panel1.Controls.Add(Me.txtPartType_Mw)
        Me.Panel1.Controls.Add(Me.Label16)
        Me.Panel1.Controls.Add(Me.txtThickness2_Mw)
        Me.Panel1.Controls.Add(Me.txtGageName_Mw)
        Me.Panel1.Controls.Add(Me.Label18)
        Me.Panel1.Controls.Add(Me.txtGrade2_Mw)
        Me.Panel1.Controls.Add(Me.Label20)
        Me.Panel1.Controls.Add(Me.Label23)
        Me.Panel1.Controls.Add(Me.txtSize2_Mw)
        Me.Panel1.Controls.Add(Me.Label26)
        Me.Panel1.Controls.Add(Me.cmbMaterialUsed2_Mw)
        Me.Panel1.Controls.Add(Me.Label27)
        Me.Panel1.Controls.Add(Me.Label32)
        Me.Panel1.Controls.Add(Me.txtMaterialUsed_C)
        Me.Panel1.Controls.Add(Me.Label34)
        Me.Panel1.Controls.Add(Me.Label33)
        Me.Panel1.Controls.Add(Me.Label31)
        Me.Panel1.Controls.Add(Me.txtSize_C)
        Me.Panel1.Controls.Add(Me.txtGrade_C)
        Me.Panel1.Controls.Add(Me.Label28)
        Me.Panel1.Controls.Add(Me.Label21)
        Me.Panel1.Controls.Add(Me.txtGageName_C)
        Me.Panel1.Controls.Add(Me.Label22)
        Me.Panel1.Controls.Add(Me.txtThickness_C)
        Me.Panel1.Controls.Add(Me.Label19)
        Me.Panel1.Controls.Add(Me.txtPartType_C)
        Me.Panel1.Controls.Add(Me.Label17)
        Me.Panel1.Controls.Add(Me.txtMaterialSpec_C)
        Me.Panel1.Controls.Add(Me.txtBendRadius_C)
        Me.Panel1.Controls.Add(Me.Label24)
        Me.Panel1.Controls.Add(Me.txtCurrentMaterial_C)
        Me.Panel1.Controls.Add(Me.Label14)
        Me.Panel1.Location = New System.Drawing.Point(511, 109)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(539, 420)
        Me.Panel1.TabIndex = 3
        '
        'AssemblyAutomationForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1063, 591)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Label35)
        Me.Controls.Add(Me.MoCloseButton1)
        Me.Controls.Add(Me.cmbCategory)
        Me.Controls.Add(Me.MoButton1)
        Me.Controls.Add(Me.ComboBoxFields)
        Me.Controls.Add(Me.txtSearch)
        Me.Controls.Add(Me.btnSearchFile)
        Me.Controls.Add(Me.btnRefresh)
        Me.Controls.Add(Me.lblExcelPath)
        Me.Controls.Add(Me.btnBrowseExcel)
        Me.Controls.Add(Me.txtExcelPath)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnApply)
        Me.Controls.Add(Me.btnOpenDocument)
        Me.Controls.Add(Me.btnGetData)
        Me.Controls.Add(Me.lblAssemblyPath)
        Me.Controls.Add(Me.lblAssemblyName)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "AssemblyAutomationForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " Assembly Automation 1.0.27"
        CType(Me.dgvDocumentDetails, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblAssemblyName As Label
    Friend WithEvents lblAssemblyPath As Label
    Friend WithEvents dgvDocumentDetails As DataGridView
    Friend WithEvents btnGetData As Button
    Friend WithEvents btnOpenDocument As Button
    Friend WithEvents btnApply As Button
    Friend WithEvents btnClose As Button
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents lblExcelPath As Label
    Friend WithEvents btnBrowseExcel As Button
    Friend WithEvents txtExcelPath As TextBox
    Friend WithEvents cmbCategory As ComboBox
    Friend WithEvents txtThickness2_Mw As TextBox
    Friend WithEvents txtGageTable As TextBox
    Friend WithEvents txtMaterialSpec2_Mw As TextBox
    Friend WithEvents txtGrade2_Mw As TextBox
    Friend WithEvents txtSize2_Mw As TextBox
    Friend WithEvents txtPartType_Mw As TextBox
    Friend WithEvents txtBendRadius_Mw As TextBox
    Friend WithEvents cmbMaterialUsed2_Mw As ComboBox
    Friend WithEvents txtGageName_Mw As TextBox
    Friend WithEvents txtPartType_C As TextBox
    Friend WithEvents txtMaterialUsed_C As TextBox
    Friend WithEvents txtBendRadius_C As TextBox
    Friend WithEvents txtGageName_C As TextBox
    Friend WithEvents txtMaterialSpec_C As TextBox
    Friend WithEvents txtGrade_C As TextBox
    Friend WithEvents txtThickness_C As TextBox
    Friend WithEvents txtSize_C As TextBox
    Friend WithEvents btnRefresh As Button
    Friend WithEvents ComboBoxFields As ComboBox
    Friend WithEvents txtSearch As TextBox
    Friend WithEvents btnSearchFile As Button
    Friend WithEvents MoButton1 As MoCustomControls.MoButton
    Friend WithEvents MoCloseButton1 As MoCustomControls.MoCloseButton
    Friend WithEvents cmbBendType_Mw As ComboBox
    Friend WithEvents txtBECMaterial2_Mw As TextBox
    Friend WithEvents txtCurrentMaterial_C As TextBox
    Friend WithEvents TableLayoutPanel2 As TableLayoutPanel
    Friend WithEvents Label27 As Label
    Friend WithEvents Label32 As Label
    Friend WithEvents Label33 As Label
    Friend WithEvents Label34 As Label
    Friend WithEvents Label35 As Label
    Friend WithEvents Label26 As Label
    Friend WithEvents Label23 As Label
    Friend WithEvents Label20 As Label
    Friend WithEvents Label18 As Label
    Friend WithEvents Label16 As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents Label13 As Label
    Friend WithEvents Label29 As Label
    Friend WithEvents Label31 As Label
    Friend WithEvents Label28 As Label
    Friend WithEvents Label21 As Label
    Friend WithEvents Label22 As Label
    Friend WithEvents Label19 As Label
    Friend WithEvents Label17 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents Label24 As Label
    Friend WithEvents Label25 As Label
    Friend WithEvents Label30 As Label
    Friend WithEvents Panel1 As Panel
End Class
