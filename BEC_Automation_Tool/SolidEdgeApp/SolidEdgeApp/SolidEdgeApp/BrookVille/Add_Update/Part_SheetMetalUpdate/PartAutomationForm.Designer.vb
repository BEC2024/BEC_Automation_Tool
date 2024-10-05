<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PartAutomationForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PartAutomationForm))
        Me.lblPartType = New System.Windows.Forms.Label()
        Me.cmbPartType_Pw = New System.Windows.Forms.ComboBox()
        Me.gpPartProperties = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.cmbGageName = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmbBendTypeGageWise_Mw = New System.Windows.Forms.ComboBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.lblMaterialUsed = New System.Windows.Forms.Label()
        Me.txtBendRadius_Mw = New System.Windows.Forms.TextBox()
        Me.txtThickness2_Mw = New System.Windows.Forms.TextBox()
        Me.txtBECMaterial2_Mw = New System.Windows.Forms.TextBox()
        Me.txtMaterialSpec2_Mw = New System.Windows.Forms.TextBox()
        Me.txtPartType_Mw = New System.Windows.Forms.TextBox()
        Me.txtGrade2_Mw = New System.Windows.Forms.TextBox()
        Me.lblSize = New System.Windows.Forms.Label()
        Me.txtSize2_Mw = New System.Windows.Forms.TextBox()
        Me.lbl6 = New System.Windows.Forms.Label()
        Me.lblGrade = New System.Windows.Forms.Label()
        Me.cmbMaterialUsed2_Mw = New System.Windows.Forms.ComboBox()
        Me.lbl9 = New System.Windows.Forms.Label()
        Me.lbl4 = New System.Windows.Forms.Label()
        Me.lbl8 = New System.Windows.Forms.Label()
        Me.txtBendRadius_C = New System.Windows.Forms.TextBox()
        Me.lbl1 = New System.Windows.Forms.Label()
        Me.lbl5 = New System.Windows.Forms.Label()
        Me.lblBendRadius = New System.Windows.Forms.Label()
        Me.lblCurrentGageName = New System.Windows.Forms.Label()
        Me.lbl3 = New System.Windows.Forms.Label()
        Me.lbl7 = New System.Windows.Forms.Label()
        Me.txtGageName_C = New System.Windows.Forms.TextBox()
        Me.lbl2 = New System.Windows.Forms.Label()
        Me.lblThickness = New System.Windows.Forms.Label()
        Me.lblMaterialSpec = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtMaterialUsed_C = New System.Windows.Forms.TextBox()
        Me.txtSize_C = New System.Windows.Forms.TextBox()
        Me.txtGrade_C = New System.Windows.Forms.TextBox()
        Me.txtCurrentMaterial_C = New System.Windows.Forms.TextBox()
        Me.txtThickness_C = New System.Windows.Forms.TextBox()
        Me.txtMaterialSpec_C = New System.Windows.Forms.TextBox()
        Me.txtPartType_C = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.rbPartTypewise = New System.Windows.Forms.RadioButton()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.cmbSize_Pw = New System.Windows.Forms.ComboBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.cmbGrade_Pw = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmbBendTypeGageWise_Pw = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cmbGageName_Pw = New System.Windows.Forms.ComboBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.cmbThickness_Pw = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cmbBendRadius_Pw = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.cmbMaterialUsed_Pw = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.cmbMaterialSpec_Pw = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtBECMaterial_Pw = New System.Windows.Forms.TextBox()
        Me.rbMaterialWise = New System.Windows.Forms.RadioButton()
        Me.lblGageTable = New System.Windows.Forms.Label()
        Me.txtGageTable = New System.Windows.Forms.TextBox()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.gpDefaultMaterialExcel = New System.Windows.Forms.GroupBox()
        Me.btnGetData = New System.Windows.Forms.Button()
        Me.lblCategory = New System.Windows.Forms.Label()
        Me.cmbMaterialLib = New System.Windows.Forms.ComboBox()
        Me.cmbCategory = New System.Windows.Forms.ComboBox()
        Me.lblMaterialLib = New System.Windows.Forms.Label()
        Me.lblExcelPath = New System.Windows.Forms.Label()
        Me.btnBrowseExcel = New System.Windows.Forms.Button()
        Me.txtExcelPath = New System.Windows.Forms.TextBox()
        Me.btnApply = New System.Windows.Forms.Button()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.btnShowGuideLines = New System.Windows.Forms.Button()
        Me.txtImageName = New System.Windows.Forms.TextBox()
        Me.gpPartProperties.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.FlowLayoutPanel1.SuspendLayout()
        Me.gpDefaultMaterialExcel.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblPartType
        '
        Me.lblPartType.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lblPartType.AutoSize = True
        Me.lblPartType.Location = New System.Drawing.Point(3, 217)
        Me.lblPartType.Name = "lblPartType"
        Me.lblPartType.Size = New System.Drawing.Size(55, 15)
        Me.lblPartType.TabIndex = 0
        Me.lblPartType.Text = "Part Type"
        '
        'cmbPartType_Pw
        '
        Me.cmbPartType_Pw.FormattingEnabled = True
        Me.cmbPartType_Pw.Location = New System.Drawing.Point(3, 43)
        Me.cmbPartType_Pw.Name = "cmbPartType_Pw"
        Me.cmbPartType_Pw.Size = New System.Drawing.Size(181, 23)
        Me.cmbPartType_Pw.TabIndex = 2
        '
        'gpPartProperties
        '
        Me.gpPartProperties.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gpPartProperties.Controls.Add(Me.TableLayoutPanel1)
        Me.gpPartProperties.Controls.Add(Me.FlowLayoutPanel1)
        Me.gpPartProperties.Location = New System.Drawing.Point(12, 151)
        Me.gpPartProperties.Name = "gpPartProperties"
        Me.gpPartProperties.Size = New System.Drawing.Size(1039, 395)
        Me.gpPartProperties.TabIndex = 3
        Me.gpPartProperties.TabStop = False
        Me.gpPartProperties.Text = "Part Properties"
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 5
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.84157!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 32.28935!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 1.742231!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.84157!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 32.28529!))
        Me.TableLayoutPanel1.Controls.Add(Me.cmbGageName, 4, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.cmbBendTypeGageWise_Mw, 4, 11)
        Me.TableLayoutPanel1.Controls.Add(Me.Label14, 3, 11)
        Me.TableLayoutPanel1.Controls.Add(Me.lblMaterialUsed, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.txtBendRadius_Mw, 4, 10)
        Me.TableLayoutPanel1.Controls.Add(Me.txtThickness2_Mw, 4, 6)
        Me.TableLayoutPanel1.Controls.Add(Me.txtBECMaterial2_Mw, 4, 9)
        Me.TableLayoutPanel1.Controls.Add(Me.txtMaterialSpec2_Mw, 4, 8)
        Me.TableLayoutPanel1.Controls.Add(Me.txtPartType_Mw, 4, 7)
        Me.TableLayoutPanel1.Controls.Add(Me.txtGrade2_Mw, 4, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.lblSize, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.txtSize2_Mw, 4, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.lbl6, 3, 10)
        Me.TableLayoutPanel1.Controls.Add(Me.lblGrade, 0, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.cmbMaterialUsed2_Mw, 4, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.lbl9, 3, 9)
        Me.TableLayoutPanel1.Controls.Add(Me.lbl4, 3, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.lbl8, 3, 8)
        Me.TableLayoutPanel1.Controls.Add(Me.txtBendRadius_C, 1, 10)
        Me.TableLayoutPanel1.Controls.Add(Me.lbl1, 3, 7)
        Me.TableLayoutPanel1.Controls.Add(Me.lbl5, 3, 6)
        Me.TableLayoutPanel1.Controls.Add(Me.lblBendRadius, 0, 10)
        Me.TableLayoutPanel1.Controls.Add(Me.lblCurrentGageName, 0, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.lbl3, 3, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.lbl7, 3, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.txtGageName_C, 1, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.lbl2, 3, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.lblThickness, 0, 6)
        Me.TableLayoutPanel1.Controls.Add(Me.lblPartType, 0, 7)
        Me.TableLayoutPanel1.Controls.Add(Me.lblMaterialSpec, 0, 8)
        Me.TableLayoutPanel1.Controls.Add(Me.Label2, 0, 9)
        Me.TableLayoutPanel1.Controls.Add(Me.txtMaterialUsed_C, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.txtSize_C, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.txtGrade_C, 1, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.txtCurrentMaterial_C, 1, 9)
        Me.TableLayoutPanel1.Controls.Add(Me.txtThickness_C, 1, 6)
        Me.TableLayoutPanel1.Controls.Add(Me.txtMaterialSpec_C, 1, 8)
        Me.TableLayoutPanel1.Controls.Add(Me.txtPartType_C, 1, 7)
        Me.TableLayoutPanel1.Controls.Add(Me.Label3, 4, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(6, 22)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 12
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11.47541!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 4.918033!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 8.333335!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 8.333335!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 8.333335!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 8.333335!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 8.333335!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 8.333335!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 8.333335!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 8.333335!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 8.333335!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 8.333335!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(870, 366)
        Me.TableLayoutPanel1.TabIndex = 83
        '
        'cmbGageName
        '
        Me.cmbGageName.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbGageName.FormattingEnabled = True
        Me.cmbGageName.Location = New System.Drawing.Point(590, 153)
        Me.cmbGageName.Name = "cmbGageName"
        Me.cmbGageName.Size = New System.Drawing.Size(277, 23)
        Me.cmbGageName.TabIndex = 100
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(231, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(109, 15)
        Me.Label1.TabIndex = 31
        Me.Label1.Text = "Current Part Details"
        '
        'cmbBendTypeGageWise_Mw
        '
        Me.cmbBendTypeGageWise_Mw.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbBendTypeGageWise_Mw.FormattingEnabled = True
        Me.cmbBendTypeGageWise_Mw.Location = New System.Drawing.Point(590, 336)
        Me.cmbBendTypeGageWise_Mw.Name = "cmbBendTypeGageWise_Mw"
        Me.cmbBendTypeGageWise_Mw.Size = New System.Drawing.Size(277, 23)
        Me.cmbBendTypeGageWise_Mw.TabIndex = 80
        '
        'Label14
        '
        Me.Label14.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(444, 340)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(61, 15)
        Me.Label14.TabIndex = 81
        Me.Label14.Text = "Bend Type"
        '
        'lblMaterialUsed
        '
        Me.lblMaterialUsed.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lblMaterialUsed.AutoSize = True
        Me.lblMaterialUsed.Location = New System.Drawing.Point(3, 67)
        Me.lblMaterialUsed.Name = "lblMaterialUsed"
        Me.lblMaterialUsed.Size = New System.Drawing.Size(79, 15)
        Me.lblMaterialUsed.TabIndex = 12
        Me.lblMaterialUsed.Text = "Material Used"
        '
        'txtBendRadius_Mw
        '
        Me.txtBendRadius_Mw.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtBendRadius_Mw.Location = New System.Drawing.Point(590, 303)
        Me.txtBendRadius_Mw.Name = "txtBendRadius_Mw"
        Me.txtBendRadius_Mw.Size = New System.Drawing.Size(277, 23)
        Me.txtBendRadius_Mw.TabIndex = 64
        '
        'txtThickness2_Mw
        '
        Me.txtThickness2_Mw.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtThickness2_Mw.Location = New System.Drawing.Point(590, 183)
        Me.txtThickness2_Mw.Name = "txtThickness2_Mw"
        Me.txtThickness2_Mw.Size = New System.Drawing.Size(277, 23)
        Me.txtThickness2_Mw.TabIndex = 42
        '
        'txtBECMaterial2_Mw
        '
        Me.txtBECMaterial2_Mw.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtBECMaterial2_Mw.Location = New System.Drawing.Point(590, 273)
        Me.txtBECMaterial2_Mw.Name = "txtBECMaterial2_Mw"
        Me.txtBECMaterial2_Mw.Size = New System.Drawing.Size(277, 23)
        Me.txtBECMaterial2_Mw.TabIndex = 46
        '
        'txtMaterialSpec2_Mw
        '
        Me.txtMaterialSpec2_Mw.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMaterialSpec2_Mw.Location = New System.Drawing.Point(590, 243)
        Me.txtMaterialSpec2_Mw.Name = "txtMaterialSpec2_Mw"
        Me.txtMaterialSpec2_Mw.Size = New System.Drawing.Size(277, 23)
        Me.txtMaterialSpec2_Mw.TabIndex = 45
        '
        'txtPartType_Mw
        '
        Me.txtPartType_Mw.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPartType_Mw.Location = New System.Drawing.Point(590, 213)
        Me.txtPartType_Mw.Name = "txtPartType_Mw"
        Me.txtPartType_Mw.Size = New System.Drawing.Size(277, 23)
        Me.txtPartType_Mw.TabIndex = 44
        '
        'txtGrade2_Mw
        '
        Me.txtGrade2_Mw.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtGrade2_Mw.Location = New System.Drawing.Point(590, 123)
        Me.txtGrade2_Mw.Name = "txtGrade2_Mw"
        Me.txtGrade2_Mw.Size = New System.Drawing.Size(277, 23)
        Me.txtGrade2_Mw.TabIndex = 43
        '
        'lblSize
        '
        Me.lblSize.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lblSize.AutoSize = True
        Me.lblSize.Location = New System.Drawing.Point(3, 97)
        Me.lblSize.Name = "lblSize"
        Me.lblSize.Size = New System.Drawing.Size(27, 15)
        Me.lblSize.TabIndex = 3
        Me.lblSize.Text = "Size"
        '
        'txtSize2_Mw
        '
        Me.txtSize2_Mw.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSize2_Mw.Location = New System.Drawing.Point(590, 93)
        Me.txtSize2_Mw.Name = "txtSize2_Mw"
        Me.txtSize2_Mw.Size = New System.Drawing.Size(277, 23)
        Me.txtSize2_Mw.TabIndex = 41
        '
        'lbl6
        '
        Me.lbl6.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lbl6.AutoSize = True
        Me.lbl6.Location = New System.Drawing.Point(444, 307)
        Me.lbl6.Name = "lbl6"
        Me.lbl6.Size = New System.Drawing.Size(72, 15)
        Me.lbl6.TabIndex = 65
        Me.lbl6.Text = "Bend Radius"
        '
        'lblGrade
        '
        Me.lblGrade.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lblGrade.AutoSize = True
        Me.lblGrade.Location = New System.Drawing.Point(3, 127)
        Me.lblGrade.Name = "lblGrade"
        Me.lblGrade.Size = New System.Drawing.Size(38, 15)
        Me.lblGrade.TabIndex = 9
        Me.lblGrade.Text = "Grade"
        '
        'cmbMaterialUsed2_Mw
        '
        Me.cmbMaterialUsed2_Mw.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbMaterialUsed2_Mw.FormattingEnabled = True
        Me.cmbMaterialUsed2_Mw.Location = New System.Drawing.Point(590, 63)
        Me.cmbMaterialUsed2_Mw.Name = "cmbMaterialUsed2_Mw"
        Me.cmbMaterialUsed2_Mw.Size = New System.Drawing.Size(277, 23)
        Me.cmbMaterialUsed2_Mw.TabIndex = 33
        '
        'lbl9
        '
        Me.lbl9.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lbl9.AutoSize = True
        Me.lbl9.Location = New System.Drawing.Point(444, 277)
        Me.lbl9.Name = "lbl9"
        Me.lbl9.Size = New System.Drawing.Size(74, 15)
        Me.lbl9.TabIndex = 55
        Me.lbl9.Text = "BEC Material"
        '
        'lbl4
        '
        Me.lbl4.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lbl4.AutoSize = True
        Me.lbl4.Location = New System.Drawing.Point(444, 157)
        Me.lbl4.Name = "lbl4"
        Me.lbl4.Size = New System.Drawing.Size(69, 15)
        Me.lbl4.TabIndex = 58
        Me.lbl4.Text = "Gage Name"
        '
        'lbl8
        '
        Me.lbl8.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lbl8.AutoSize = True
        Me.lbl8.Location = New System.Drawing.Point(444, 247)
        Me.lbl8.Name = "lbl8"
        Me.lbl8.Size = New System.Drawing.Size(78, 15)
        Me.lbl8.TabIndex = 54
        Me.lbl8.Text = "Material Spec"
        '
        'txtBendRadius_C
        '
        Me.txtBendRadius_C.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtBendRadius_C.Location = New System.Drawing.Point(149, 303)
        Me.txtBendRadius_C.Name = "txtBendRadius_C"
        Me.txtBendRadius_C.Size = New System.Drawing.Size(274, 23)
        Me.txtBendRadius_C.TabIndex = 62
        '
        'lbl1
        '
        Me.lbl1.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lbl1.AutoSize = True
        Me.lbl1.Location = New System.Drawing.Point(444, 217)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(55, 15)
        Me.lbl1.TabIndex = 49
        Me.lbl1.Text = "Part Type"
        '
        'lbl5
        '
        Me.lbl5.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lbl5.AutoSize = True
        Me.lbl5.Location = New System.Drawing.Point(444, 187)
        Me.lbl5.Name = "lbl5"
        Me.lbl5.Size = New System.Drawing.Size(138, 15)
        Me.lbl5.TabIndex = 51
        Me.lbl5.Text = "Material Thickness (inch)"
        '
        'lblBendRadius
        '
        Me.lblBendRadius.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lblBendRadius.AutoSize = True
        Me.lblBendRadius.Location = New System.Drawing.Point(3, 307)
        Me.lblBendRadius.Name = "lblBendRadius"
        Me.lblBendRadius.Size = New System.Drawing.Size(72, 15)
        Me.lblBendRadius.TabIndex = 61
        Me.lblBendRadius.Text = "Bend Radius"
        '
        'lblCurrentGageName
        '
        Me.lblCurrentGageName.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lblCurrentGageName.AutoSize = True
        Me.lblCurrentGageName.Location = New System.Drawing.Point(3, 157)
        Me.lblCurrentGageName.Name = "lblCurrentGageName"
        Me.lblCurrentGageName.Size = New System.Drawing.Size(69, 15)
        Me.lblCurrentGageName.TabIndex = 56
        Me.lblCurrentGageName.Text = "Gage Name"
        '
        'lbl3
        '
        Me.lbl3.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lbl3.AutoSize = True
        Me.lbl3.Location = New System.Drawing.Point(444, 127)
        Me.lbl3.Name = "lbl3"
        Me.lbl3.Size = New System.Drawing.Size(38, 15)
        Me.lbl3.TabIndex = 52
        Me.lbl3.Text = "Grade"
        '
        'lbl7
        '
        Me.lbl7.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lbl7.AutoSize = True
        Me.lbl7.Location = New System.Drawing.Point(444, 67)
        Me.lbl7.Name = "lbl7"
        Me.lbl7.Size = New System.Drawing.Size(79, 15)
        Me.lbl7.TabIndex = 53
        Me.lbl7.Text = "Material Used"
        '
        'txtGageName_C
        '
        Me.txtGageName_C.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtGageName_C.Location = New System.Drawing.Point(149, 153)
        Me.txtGageName_C.Name = "txtGageName_C"
        Me.txtGageName_C.Size = New System.Drawing.Size(274, 23)
        Me.txtGageName_C.TabIndex = 57
        '
        'lbl2
        '
        Me.lbl2.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lbl2.AutoSize = True
        Me.lbl2.Location = New System.Drawing.Point(444, 97)
        Me.lbl2.Name = "lbl2"
        Me.lbl2.Size = New System.Drawing.Size(27, 15)
        Me.lbl2.TabIndex = 50
        Me.lbl2.Text = "Size"
        '
        'lblThickness
        '
        Me.lblThickness.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lblThickness.AutoSize = True
        Me.lblThickness.Location = New System.Drawing.Point(3, 187)
        Me.lblThickness.Name = "lblThickness"
        Me.lblThickness.Size = New System.Drawing.Size(138, 15)
        Me.lblThickness.TabIndex = 6
        Me.lblThickness.Text = "Material Thickness (inch)"
        '
        'lblMaterialSpec
        '
        Me.lblMaterialSpec.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lblMaterialSpec.AutoSize = True
        Me.lblMaterialSpec.Location = New System.Drawing.Point(3, 247)
        Me.lblMaterialSpec.Name = "lblMaterialSpec"
        Me.lblMaterialSpec.Size = New System.Drawing.Size(78, 15)
        Me.lblMaterialSpec.TabIndex = 15
        Me.lblMaterialSpec.Text = "Material Spec"
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 277)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(74, 15)
        Me.Label2.TabIndex = 28
        Me.Label2.Text = "BEC Material"
        '
        'txtMaterialUsed_C
        '
        Me.txtMaterialUsed_C.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMaterialUsed_C.Location = New System.Drawing.Point(149, 63)
        Me.txtMaterialUsed_C.Name = "txtMaterialUsed_C"
        Me.txtMaterialUsed_C.Size = New System.Drawing.Size(274, 23)
        Me.txtMaterialUsed_C.TabIndex = 13
        '
        'txtSize_C
        '
        Me.txtSize_C.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSize_C.Enabled = False
        Me.txtSize_C.Location = New System.Drawing.Point(149, 93)
        Me.txtSize_C.Name = "txtSize_C"
        Me.txtSize_C.Size = New System.Drawing.Size(274, 23)
        Me.txtSize_C.TabIndex = 4
        '
        'txtGrade_C
        '
        Me.txtGrade_C.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtGrade_C.Enabled = False
        Me.txtGrade_C.Location = New System.Drawing.Point(149, 123)
        Me.txtGrade_C.Name = "txtGrade_C"
        Me.txtGrade_C.Size = New System.Drawing.Size(274, 23)
        Me.txtGrade_C.TabIndex = 10
        '
        'txtCurrentMaterial_C
        '
        Me.txtCurrentMaterial_C.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCurrentMaterial_C.Location = New System.Drawing.Point(149, 273)
        Me.txtCurrentMaterial_C.Name = "txtCurrentMaterial_C"
        Me.txtCurrentMaterial_C.Size = New System.Drawing.Size(274, 23)
        Me.txtCurrentMaterial_C.TabIndex = 30
        '
        'txtThickness_C
        '
        Me.txtThickness_C.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtThickness_C.Location = New System.Drawing.Point(149, 183)
        Me.txtThickness_C.Name = "txtThickness_C"
        Me.txtThickness_C.Size = New System.Drawing.Size(274, 23)
        Me.txtThickness_C.TabIndex = 7
        '
        'txtMaterialSpec_C
        '
        Me.txtMaterialSpec_C.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMaterialSpec_C.Location = New System.Drawing.Point(149, 243)
        Me.txtMaterialSpec_C.Name = "txtMaterialSpec_C"
        Me.txtMaterialSpec_C.Size = New System.Drawing.Size(274, 23)
        Me.txtMaterialSpec_C.TabIndex = 16
        '
        'txtPartType_C
        '
        Me.txtPartType_C.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPartType_C.Enabled = False
        Me.txtPartType_C.Location = New System.Drawing.Point(149, 213)
        Me.txtPartType_C.Name = "txtPartType_C"
        Me.txtPartType_C.Size = New System.Drawing.Size(274, 23)
        Me.txtPartType_C.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(682, 27)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(93, 15)
        Me.Label3.TabIndex = 32
        Me.Label3.Text = "New Part Details"
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Controls.Add(Me.rbPartTypewise)
        Me.FlowLayoutPanel1.Controls.Add(Me.Label13)
        Me.FlowLayoutPanel1.Controls.Add(Me.cmbPartType_Pw)
        Me.FlowLayoutPanel1.Controls.Add(Me.Label12)
        Me.FlowLayoutPanel1.Controls.Add(Me.cmbSize_Pw)
        Me.FlowLayoutPanel1.Controls.Add(Me.Label10)
        Me.FlowLayoutPanel1.Controls.Add(Me.cmbGrade_Pw)
        Me.FlowLayoutPanel1.Controls.Add(Me.Label4)
        Me.FlowLayoutPanel1.Controls.Add(Me.cmbBendTypeGageWise_Pw)
        Me.FlowLayoutPanel1.Controls.Add(Me.Label6)
        Me.FlowLayoutPanel1.Controls.Add(Me.cmbGageName_Pw)
        Me.FlowLayoutPanel1.Controls.Add(Me.Label11)
        Me.FlowLayoutPanel1.Controls.Add(Me.cmbThickness_Pw)
        Me.FlowLayoutPanel1.Controls.Add(Me.Label5)
        Me.FlowLayoutPanel1.Controls.Add(Me.cmbBendRadius_Pw)
        Me.FlowLayoutPanel1.Controls.Add(Me.Label9)
        Me.FlowLayoutPanel1.Controls.Add(Me.cmbMaterialUsed_Pw)
        Me.FlowLayoutPanel1.Controls.Add(Me.Label8)
        Me.FlowLayoutPanel1.Controls.Add(Me.cmbMaterialSpec_Pw)
        Me.FlowLayoutPanel1.Controls.Add(Me.Label7)
        Me.FlowLayoutPanel1.Controls.Add(Me.txtBECMaterial_Pw)
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(879, 22)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(151, 453)
        Me.FlowLayoutPanel1.TabIndex = 82
        Me.FlowLayoutPanel1.Visible = False
        '
        'rbPartTypewise
        '
        Me.rbPartTypewise.AutoSize = True
        Me.rbPartTypewise.Location = New System.Drawing.Point(3, 3)
        Me.rbPartTypewise.Name = "rbPartTypewise"
        Me.rbPartTypewise.Size = New System.Drawing.Size(98, 19)
        Me.rbPartTypewise.TabIndex = 47
        Me.rbPartTypewise.Text = "Part type wise"
        Me.rbPartTypewise.UseVisualStyleBackColor = True
        Me.rbPartTypewise.Visible = False
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(3, 25)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(55, 15)
        Me.Label13.TabIndex = 71
        Me.Label13.Text = "Part Type"
        Me.Label13.Visible = False
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(3, 69)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(27, 15)
        Me.Label12.TabIndex = 72
        Me.Label12.Text = "Size"
        Me.Label12.Visible = False
        '
        'cmbSize_Pw
        '
        Me.cmbSize_Pw.FormattingEnabled = True
        Me.cmbSize_Pw.Location = New System.Drawing.Point(3, 87)
        Me.cmbSize_Pw.Name = "cmbSize_Pw"
        Me.cmbSize_Pw.Size = New System.Drawing.Size(181, 23)
        Me.cmbSize_Pw.TabIndex = 5
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(3, 113)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(38, 15)
        Me.Label10.TabIndex = 74
        Me.Label10.Text = "Grade"
        Me.Label10.Visible = False
        '
        'cmbGrade_Pw
        '
        Me.cmbGrade_Pw.FormattingEnabled = True
        Me.cmbGrade_Pw.Location = New System.Drawing.Point(3, 131)
        Me.cmbGrade_Pw.Name = "cmbGrade_Pw"
        Me.cmbGrade_Pw.Size = New System.Drawing.Size(181, 23)
        Me.cmbGrade_Pw.TabIndex = 11
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 157)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(61, 15)
        Me.Label4.TabIndex = 70
        Me.Label4.Text = "Bend Type"
        Me.Label4.Visible = False
        '
        'cmbBendTypeGageWise_Pw
        '
        Me.cmbBendTypeGageWise_Pw.FormattingEnabled = True
        Me.cmbBendTypeGageWise_Pw.Location = New System.Drawing.Point(3, 175)
        Me.cmbBendTypeGageWise_Pw.Name = "cmbBendTypeGageWise_Pw"
        Me.cmbBendTypeGageWise_Pw.Size = New System.Drawing.Size(181, 23)
        Me.cmbBendTypeGageWise_Pw.TabIndex = 69
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 201)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(69, 15)
        Me.Label6.TabIndex = 78
        Me.Label6.Text = "Gage Name"
        Me.Label6.Visible = False
        '
        'cmbGageName_Pw
        '
        Me.cmbGageName_Pw.FormattingEnabled = True
        Me.cmbGageName_Pw.Location = New System.Drawing.Point(3, 219)
        Me.cmbGageName_Pw.Name = "cmbGageName_Pw"
        Me.cmbGageName_Pw.Size = New System.Drawing.Size(181, 23)
        Me.cmbGageName_Pw.TabIndex = 60
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(3, 245)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(138, 15)
        Me.Label11.TabIndex = 73
        Me.Label11.Text = "Material Thickness (inch)"
        Me.Label11.Visible = False
        '
        'cmbThickness_Pw
        '
        Me.cmbThickness_Pw.FormattingEnabled = True
        Me.cmbThickness_Pw.Location = New System.Drawing.Point(3, 263)
        Me.cmbThickness_Pw.Name = "cmbThickness_Pw"
        Me.cmbThickness_Pw.Size = New System.Drawing.Size(181, 23)
        Me.cmbThickness_Pw.TabIndex = 8
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 289)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(72, 15)
        Me.Label5.TabIndex = 79
        Me.Label5.Text = "Bend Radius"
        Me.Label5.Visible = False
        '
        'cmbBendRadius_Pw
        '
        Me.cmbBendRadius_Pw.FormattingEnabled = True
        Me.cmbBendRadius_Pw.Location = New System.Drawing.Point(3, 307)
        Me.cmbBendRadius_Pw.Name = "cmbBendRadius_Pw"
        Me.cmbBendRadius_Pw.Size = New System.Drawing.Size(181, 23)
        Me.cmbBendRadius_Pw.TabIndex = 63
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(3, 333)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(79, 15)
        Me.Label9.TabIndex = 75
        Me.Label9.Text = "Material Used"
        Me.Label9.Visible = False
        '
        'cmbMaterialUsed_Pw
        '
        Me.cmbMaterialUsed_Pw.FormattingEnabled = True
        Me.cmbMaterialUsed_Pw.Location = New System.Drawing.Point(3, 351)
        Me.cmbMaterialUsed_Pw.Name = "cmbMaterialUsed_Pw"
        Me.cmbMaterialUsed_Pw.Size = New System.Drawing.Size(181, 23)
        Me.cmbMaterialUsed_Pw.TabIndex = 14
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(3, 377)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(78, 15)
        Me.Label8.TabIndex = 76
        Me.Label8.Text = "Material Spec"
        Me.Label8.Visible = False
        '
        'cmbMaterialSpec_Pw
        '
        Me.cmbMaterialSpec_Pw.FormattingEnabled = True
        Me.cmbMaterialSpec_Pw.Location = New System.Drawing.Point(3, 395)
        Me.cmbMaterialSpec_Pw.Name = "cmbMaterialSpec_Pw"
        Me.cmbMaterialSpec_Pw.Size = New System.Drawing.Size(181, 23)
        Me.cmbMaterialSpec_Pw.TabIndex = 17
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(3, 421)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(74, 15)
        Me.Label7.TabIndex = 77
        Me.Label7.Text = "BEC Material"
        Me.Label7.Visible = False
        '
        'txtBECMaterial_Pw
        '
        Me.txtBECMaterial_Pw.Location = New System.Drawing.Point(3, 439)
        Me.txtBECMaterial_Pw.Name = "txtBECMaterial_Pw"
        Me.txtBECMaterial_Pw.Size = New System.Drawing.Size(181, 23)
        Me.txtBECMaterial_Pw.TabIndex = 29
        '
        'rbMaterialWise
        '
        Me.rbMaterialWise.AutoSize = True
        Me.rbMaterialWise.Checked = True
        Me.rbMaterialWise.Location = New System.Drawing.Point(773, 80)
        Me.rbMaterialWise.Name = "rbMaterialWise"
        Me.rbMaterialWise.Size = New System.Drawing.Size(94, 19)
        Me.rbMaterialWise.TabIndex = 48
        Me.rbMaterialWise.TabStop = True
        Me.rbMaterialWise.Text = "Material wise"
        Me.rbMaterialWise.UseVisualStyleBackColor = True
        Me.rbMaterialWise.Visible = False
        '
        'lblGageTable
        '
        Me.lblGageTable.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblGageTable.AutoSize = True
        Me.lblGageTable.Location = New System.Drawing.Point(581, 562)
        Me.lblGageTable.Name = "lblGageTable"
        Me.lblGageTable.Size = New System.Drawing.Size(64, 15)
        Me.lblGageTable.TabIndex = 68
        Me.lblGageTable.Text = "Gage Table"
        Me.lblGageTable.Visible = False
        '
        'txtGageTable
        '
        Me.txtGageTable.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtGageTable.Location = New System.Drawing.Point(651, 558)
        Me.txtGageTable.Name = "txtGageTable"
        Me.txtGageTable.Size = New System.Drawing.Size(181, 23)
        Me.txtGageTable.TabIndex = 67
        Me.txtGageTable.Visible = False
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(837, 554)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(100, 30)
        Me.btnClose.TabIndex = 5
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'gpDefaultMaterialExcel
        '
        Me.gpDefaultMaterialExcel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gpDefaultMaterialExcel.Controls.Add(Me.btnGetData)
        Me.gpDefaultMaterialExcel.Controls.Add(Me.lblCategory)
        Me.gpDefaultMaterialExcel.Controls.Add(Me.rbMaterialWise)
        Me.gpDefaultMaterialExcel.Controls.Add(Me.cmbMaterialLib)
        Me.gpDefaultMaterialExcel.Controls.Add(Me.cmbCategory)
        Me.gpDefaultMaterialExcel.Controls.Add(Me.lblMaterialLib)
        Me.gpDefaultMaterialExcel.Controls.Add(Me.lblExcelPath)
        Me.gpDefaultMaterialExcel.Controls.Add(Me.btnBrowseExcel)
        Me.gpDefaultMaterialExcel.Controls.Add(Me.txtExcelPath)
        Me.gpDefaultMaterialExcel.Location = New System.Drawing.Point(12, 12)
        Me.gpDefaultMaterialExcel.Name = "gpDefaultMaterialExcel"
        Me.gpDefaultMaterialExcel.Size = New System.Drawing.Size(1039, 133)
        Me.gpDefaultMaterialExcel.TabIndex = 6
        Me.gpDefaultMaterialExcel.TabStop = False
        Me.gpDefaultMaterialExcel.Text = "Standard Part Details"
        '
        'btnGetData
        '
        Me.btnGetData.Location = New System.Drawing.Point(17, 72)
        Me.btnGetData.Name = "btnGetData"
        Me.btnGetData.Size = New System.Drawing.Size(195, 30)
        Me.btnGetData.TabIndex = 99
        Me.btnGetData.Text = "Get BEC Material Data"
        Me.btnGetData.UseVisualStyleBackColor = True
        '
        'lblCategory
        '
        Me.lblCategory.AutoSize = True
        Me.lblCategory.Location = New System.Drawing.Point(225, 80)
        Me.lblCategory.Name = "lblCategory"
        Me.lblCategory.Size = New System.Drawing.Size(55, 15)
        Me.lblCategory.TabIndex = 51
        Me.lblCategory.Text = "Category"
        '
        'cmbMaterialLib
        '
        Me.cmbMaterialLib.FormattingEnabled = True
        Me.cmbMaterialLib.Location = New System.Drawing.Point(554, 76)
        Me.cmbMaterialLib.Name = "cmbMaterialLib"
        Me.cmbMaterialLib.Size = New System.Drawing.Size(181, 23)
        Me.cmbMaterialLib.TabIndex = 69
        Me.cmbMaterialLib.Visible = False
        '
        'cmbCategory
        '
        Me.cmbCategory.FormattingEnabled = True
        Me.cmbCategory.Location = New System.Drawing.Point(293, 76)
        Me.cmbCategory.Name = "cmbCategory"
        Me.cmbCategory.Size = New System.Drawing.Size(181, 23)
        Me.cmbCategory.TabIndex = 50
        '
        'lblMaterialLib
        '
        Me.lblMaterialLib.AutoSize = True
        Me.lblMaterialLib.Location = New System.Drawing.Point(487, 80)
        Me.lblMaterialLib.Name = "lblMaterialLib"
        Me.lblMaterialLib.Size = New System.Drawing.Size(69, 15)
        Me.lblMaterialLib.TabIndex = 68
        Me.lblMaterialLib.Text = "Material Lib"
        Me.lblMaterialLib.Visible = False
        '
        'lblExcelPath
        '
        Me.lblExcelPath.AutoSize = True
        Me.lblExcelPath.Location = New System.Drawing.Point(17, 44)
        Me.lblExcelPath.Name = "lblExcelPath"
        Me.lblExcelPath.Size = New System.Drawing.Size(131, 15)
        Me.lblExcelPath.TabIndex = 15
        Me.lblExcelPath.Text = "BEC Material Excel Path"
        '
        'btnBrowseExcel
        '
        Me.btnBrowseExcel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrowseExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnBrowseExcel.Location = New System.Drawing.Point(930, 36)
        Me.btnBrowseExcel.Name = "btnBrowseExcel"
        Me.btnBrowseExcel.Size = New System.Drawing.Size(100, 30)
        Me.btnBrowseExcel.TabIndex = 17
        Me.btnBrowseExcel.Text = "Browse"
        Me.btnBrowseExcel.UseVisualStyleBackColor = True
        Me.btnBrowseExcel.Visible = False
        '
        'txtExcelPath
        '
        Me.txtExcelPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtExcelPath.Enabled = False
        Me.txtExcelPath.Location = New System.Drawing.Point(158, 40)
        Me.txtExcelPath.Name = "txtExcelPath"
        Me.txtExcelPath.Size = New System.Drawing.Size(767, 23)
        Me.txtExcelPath.TabIndex = 16
        '
        'btnApply
        '
        Me.btnApply.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnApply.Location = New System.Drawing.Point(943, 554)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.Size = New System.Drawing.Size(100, 30)
        Me.btnApply.TabIndex = 18
        Me.btnApply.Text = "Apply"
        Me.btnApply.UseVisualStyleBackColor = True
        '
        'btnRefresh
        '
        Me.btnRefresh.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnRefresh.Location = New System.Drawing.Point(12, 554)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(100, 30)
        Me.btnRefresh.TabIndex = 19
        Me.btnRefresh.Text = "Refresh"
        Me.btnRefresh.UseVisualStyleBackColor = True
        '
        'btnShowGuideLines
        '
        Me.btnShowGuideLines.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnShowGuideLines.Location = New System.Drawing.Point(305, 554)
        Me.btnShowGuideLines.Name = "btnShowGuideLines"
        Me.btnShowGuideLines.Size = New System.Drawing.Size(134, 30)
        Me.btnShowGuideLines.TabIndex = 20
        Me.btnShowGuideLines.Text = "Show Guidelines"
        Me.btnShowGuideLines.UseVisualStyleBackColor = True
        Me.btnShowGuideLines.Visible = False
        '
        'txtImageName
        '
        Me.txtImageName.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtImageName.Location = New System.Drawing.Point(118, 556)
        Me.txtImageName.Name = "txtImageName"
        Me.txtImageName.Size = New System.Drawing.Size(181, 23)
        Me.txtImageName.TabIndex = 68
        Me.txtImageName.Visible = False
        '
        'PartAutomationForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(1063, 591)
        Me.Controls.Add(Me.txtImageName)
        Me.Controls.Add(Me.btnShowGuideLines)
        Me.Controls.Add(Me.btnRefresh)
        Me.Controls.Add(Me.btnApply)
        Me.Controls.Add(Me.gpDefaultMaterialExcel)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.gpPartProperties)
        Me.Controls.Add(Me.txtGageTable)
        Me.Controls.Add(Me.lblGageTable)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "PartAutomationForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = " Part Automation 1.0.27"
        Me.gpPartProperties.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.FlowLayoutPanel1.PerformLayout()
        Me.gpDefaultMaterialExcel.ResumeLayout(False)
        Me.gpDefaultMaterialExcel.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblPartType As Label
    Friend WithEvents cmbPartType_Pw As ComboBox
    Friend WithEvents gpPartProperties As GroupBox
    Friend WithEvents cmbSize_Pw As ComboBox
    Friend WithEvents lblSize As Label
    Friend WithEvents txtSize_C As TextBox
    Friend WithEvents cmbThickness_Pw As ComboBox
    Friend WithEvents lblThickness As Label
    Friend WithEvents txtThickness_C As TextBox
    Friend WithEvents cmbMaterialSpec_Pw As ComboBox
    Friend WithEvents lblMaterialSpec As Label
    Friend WithEvents txtMaterialSpec_C As TextBox
    Friend WithEvents cmbMaterialUsed_Pw As ComboBox
    Friend WithEvents lblMaterialUsed As Label
    Friend WithEvents cmbGrade_Pw As ComboBox
    Friend WithEvents lblGrade As Label
    Friend WithEvents txtGrade_C As TextBox
    Friend WithEvents btnClose As Button
    Friend WithEvents gpDefaultMaterialExcel As GroupBox
    Friend WithEvents lblExcelPath As Label
    Friend WithEvents btnBrowseExcel As Button
    Friend WithEvents txtExcelPath As TextBox
    Friend WithEvents txtBECMaterial_Pw As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents txtCurrentMaterial_C As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents cmbMaterialUsed2_Mw As ComboBox
    Friend WithEvents txtBECMaterial2_Mw As TextBox
    Friend WithEvents txtMaterialSpec2_Mw As TextBox
    Friend WithEvents txtPartType_Mw As TextBox
    Friend WithEvents txtGrade2_Mw As TextBox
    Friend WithEvents txtThickness2_Mw As TextBox
    Friend WithEvents txtSize2_Mw As TextBox
    Friend WithEvents rbMaterialWise As RadioButton
    Friend WithEvents rbPartTypewise As RadioButton
    Friend WithEvents lbl9 As Label
    Friend WithEvents lbl8 As Label
    Friend WithEvents lbl7 As Label
    Friend WithEvents lbl3 As Label
    Friend WithEvents lbl5 As Label
    Friend WithEvents lbl2 As Label
    Friend WithEvents lbl1 As Label
    Friend WithEvents btnApply As Button
    Friend WithEvents btnRefresh As Button
    Friend WithEvents txtGageName_C As TextBox
    Friend WithEvents lblCurrentGageName As Label
    Friend WithEvents cmbGageName_Pw As ComboBox
    Friend WithEvents lbl4 As Label
    Friend WithEvents lbl6 As Label
    Friend WithEvents lblBendRadius As Label
    Friend WithEvents txtBendRadius_C As TextBox
    Friend WithEvents cmbBendRadius_Pw As ComboBox
    Friend WithEvents txtBendRadius_Mw As TextBox
    Friend WithEvents lblMaterialLib As Label
    Friend WithEvents cmbMaterialLib As ComboBox
    Friend WithEvents txtGageTable As TextBox
    Friend WithEvents lblCategory As Label
    Friend WithEvents cmbCategory As ComboBox
    Friend WithEvents txtMaterialUsed_C As TextBox
    Friend WithEvents txtPartType_C As TextBox
    Friend WithEvents btnShowGuideLines As Button
    Friend WithEvents txtImageName As TextBox
    Friend WithEvents lblGageTable As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents cmbBendTypeGageWise_Pw As ComboBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents Label12 As Label
    Friend WithEvents Label13 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents cmbBendTypeGageWise_Mw As ComboBox
    Friend WithEvents btnGetData As Button
    Friend WithEvents FlowLayoutPanel1 As FlowLayoutPanel
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents cmbGageName As ComboBox
End Class
