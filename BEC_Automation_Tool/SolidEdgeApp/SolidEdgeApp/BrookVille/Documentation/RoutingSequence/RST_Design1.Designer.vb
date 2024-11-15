<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class RST_Design1
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.BtnReset = New System.Windows.Forms.Button()
        Me.BtnCalculateAndSave = New System.Windows.Forms.Button()
        Me.DGV_Calculator = New System.Windows.Forms.DataGridView()
        Me.BtnPreview = New System.Windows.Forms.Button()
        Me.cmbCategory = New System.Windows.Forms.ComboBox()
        Me.BtnApplyValues = New System.Windows.Forms.Button()
        Me.BtnDelete = New System.Windows.Forms.Button()
        Me.BtnDown = New System.Windows.Forms.Button()
        Me.BtnUp = New System.Windows.Forms.Button()
        Me.cmbProcess = New System.Windows.Forms.ComboBox()
        Me.BtnGetDataAndReset = New System.Windows.Forms.Button()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.btnSelect = New System.Windows.Forms.Button()
        Me.txtFoldername = New System.Windows.Forms.TextBox()
        Me.txtFilename = New System.Windows.Forms.TextBox()
        Me.lblDir = New System.Windows.Forms.Label()
        Me.lblFilePath = New System.Windows.Forms.Label()
        Me.dgvsub2 = New System.Windows.Forms.DataGridView()
        Me.lblSubGridTitle = New System.Windows.Forms.Label()
        Me.lblUserName = New System.Windows.Forms.Label()
        Me.lblMainGridTitle = New System.Windows.Forms.Label()
        Me.btnApply = New System.Windows.Forms.Button()
        Me.btnAproveSequence = New System.Windows.Forms.Button()
        Me.dgvSub = New System.Windows.Forms.DataGridView()
        Me.dgvMain = New System.Windows.Forms.DataGridView()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.DGV_Calculator, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvsub2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvSub, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'BtnReset
        '
        Me.BtnReset.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BtnReset.Enabled = False
        Me.BtnReset.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnReset.Location = New System.Drawing.Point(638, 544)
        Me.BtnReset.Name = "BtnReset"
        Me.BtnReset.Size = New System.Drawing.Size(117, 27)
        Me.BtnReset.TabIndex = 23
        Me.BtnReset.Text = "Reset"
        Me.BtnReset.UseVisualStyleBackColor = True
        '
        'BtnCalculateAndSave
        '
        Me.BtnCalculateAndSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BtnCalculateAndSave.Enabled = False
        Me.BtnCalculateAndSave.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnCalculateAndSave.Location = New System.Drawing.Point(519, 544)
        Me.BtnCalculateAndSave.Name = "BtnCalculateAndSave"
        Me.BtnCalculateAndSave.Size = New System.Drawing.Size(117, 27)
        Me.BtnCalculateAndSave.TabIndex = 22
        Me.BtnCalculateAndSave.Text = "Calculate"
        Me.BtnCalculateAndSave.UseVisualStyleBackColor = True
        '
        'DGV_Calculator
        '
        Me.DGV_Calculator.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.DGV_Calculator.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV_Calculator.Location = New System.Drawing.Point(501, 313)
        Me.DGV_Calculator.Name = "DGV_Calculator"
        Me.DGV_Calculator.RowHeadersVisible = False
        Me.DGV_Calculator.Size = New System.Drawing.Size(254, 225)
        Me.DGV_Calculator.TabIndex = 20
        '
        'BtnPreview
        '
        Me.BtnPreview.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnPreview.Enabled = False
        Me.BtnPreview.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnPreview.Image = Global.SolidEdgeApp.My.Resources.Resources.SolidEdgePreview
        Me.BtnPreview.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnPreview.Location = New System.Drawing.Point(863, 232)
        Me.BtnPreview.Name = "BtnPreview"
        Me.BtnPreview.Size = New System.Drawing.Size(156, 32)
        Me.BtnPreview.TabIndex = 19
        Me.BtnPreview.Text = "Preview"
        Me.BtnPreview.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.BtnPreview.UseVisualStyleBackColor = True
        '
        'cmbCategory
        '
        Me.cmbCategory.BackColor = System.Drawing.SystemColors.Window
        Me.cmbCategory.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbCategory.FormattingEnabled = True
        Me.cmbCategory.Location = New System.Drawing.Point(106, 79)
        Me.cmbCategory.Name = "cmbCategory"
        Me.cmbCategory.Size = New System.Drawing.Size(123, 23)
        Me.cmbCategory.TabIndex = 16
        Me.cmbCategory.Text = "CATEGORY#"
        '
        'BtnApplyValues
        '
        Me.BtnApplyValues.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnApplyValues.Enabled = False
        Me.BtnApplyValues.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnApplyValues.Location = New System.Drawing.Point(784, 544)
        Me.BtnApplyValues.Name = "BtnApplyValues"
        Me.BtnApplyValues.Size = New System.Drawing.Size(117, 27)
        Me.BtnApplyValues.TabIndex = 15
        Me.BtnApplyValues.Text = "Apply Values"
        Me.BtnApplyValues.UseVisualStyleBackColor = True
        '
        'BtnDelete
        '
        Me.BtnDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnDelete.Enabled = False
        Me.BtnDelete.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnDelete.Location = New System.Drawing.Point(863, 201)
        Me.BtnDelete.Name = "BtnDelete"
        Me.BtnDelete.Size = New System.Drawing.Size(156, 28)
        Me.BtnDelete.TabIndex = 14
        Me.BtnDelete.Text = "Delete"
        Me.BtnDelete.UseVisualStyleBackColor = True
        '
        'BtnDown
        '
        Me.BtnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnDown.Enabled = False
        Me.BtnDown.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnDown.Location = New System.Drawing.Point(946, 168)
        Me.BtnDown.Name = "BtnDown"
        Me.BtnDown.Size = New System.Drawing.Size(73, 27)
        Me.BtnDown.TabIndex = 13
        Me.BtnDown.Text = "Down"
        Me.BtnDown.UseVisualStyleBackColor = True
        '
        'BtnUp
        '
        Me.BtnUp.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnUp.Enabled = False
        Me.BtnUp.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnUp.Location = New System.Drawing.Point(863, 168)
        Me.BtnUp.Name = "BtnUp"
        Me.BtnUp.Size = New System.Drawing.Size(73, 27)
        Me.BtnUp.TabIndex = 12
        Me.BtnUp.Text = "UP"
        Me.BtnUp.UseVisualStyleBackColor = True
        '
        'cmbProcess
        '
        Me.cmbProcess.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbProcess.BackColor = System.Drawing.SystemColors.Window
        Me.cmbProcess.Enabled = False
        Me.cmbProcess.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbProcess.FormattingEnabled = True
        Me.cmbProcess.Location = New System.Drawing.Point(863, 139)
        Me.cmbProcess.Name = "cmbProcess"
        Me.cmbProcess.Size = New System.Drawing.Size(156, 23)
        Me.cmbProcess.TabIndex = 11
        Me.cmbProcess.Text = "PROCESS#"
        '
        'BtnGetDataAndReset
        '
        Me.BtnGetDataAndReset.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnGetDataAndReset.Location = New System.Drawing.Point(235, 77)
        Me.BtnGetDataAndReset.Name = "BtnGetDataAndReset"
        Me.BtnGetDataAndReset.Size = New System.Drawing.Size(123, 27)
        Me.BtnGetDataAndReset.TabIndex = 12
        Me.BtnGetDataAndReset.Text = "Get Data"
        Me.BtnGetDataAndReset.UseVisualStyleBackColor = True
        '
        'btnBrowse
        '
        Me.btnBrowse.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrowse.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBrowse.Location = New System.Drawing.Point(932, 41)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(87, 27)
        Me.btnBrowse.TabIndex = 11
        Me.btnBrowse.Text = "Browse"
        Me.btnBrowse.UseVisualStyleBackColor = True
        Me.btnBrowse.Visible = False
        '
        'btnSelect
        '
        Me.btnSelect.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSelect.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSelect.Location = New System.Drawing.Point(932, 10)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(87, 27)
        Me.btnSelect.TabIndex = 10
        Me.btnSelect.Text = "Select"
        Me.btnSelect.UseVisualStyleBackColor = True
        '
        'txtFoldername
        '
        Me.txtFoldername.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFoldername.Enabled = False
        Me.txtFoldername.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFoldername.Location = New System.Drawing.Point(106, 43)
        Me.txtFoldername.Name = "txtFoldername"
        Me.txtFoldername.Size = New System.Drawing.Size(820, 23)
        Me.txtFoldername.TabIndex = 9
        '
        'txtFilename
        '
        Me.txtFilename.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFilename.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFilename.Location = New System.Drawing.Point(106, 12)
        Me.txtFilename.Name = "txtFilename"
        Me.txtFilename.Size = New System.Drawing.Size(820, 23)
        Me.txtFilename.TabIndex = 8
        '
        'lblDir
        '
        Me.lblDir.AutoSize = True
        Me.lblDir.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDir.Location = New System.Drawing.Point(6, 46)
        Me.lblDir.Name = "lblDir"
        Me.lblDir.Size = New System.Drawing.Size(96, 15)
        Me.lblDir.TabIndex = 7
        Me.lblDir.Text = "Output Directory"
        '
        'lblFilePath
        '
        Me.lblFilePath.AutoSize = True
        Me.lblFilePath.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFilePath.Location = New System.Drawing.Point(6, 15)
        Me.lblFilePath.Name = "lblFilePath"
        Me.lblFilePath.Size = New System.Drawing.Size(73, 15)
        Me.lblFilePath.TabIndex = 6
        Me.lblFilePath.Text = "Input Report"
        '
        'dgvsub2
        '
        Me.dgvsub2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvsub2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvsub2.Location = New System.Drawing.Point(761, 313)
        Me.dgvsub2.Name = "dgvsub2"
        Me.dgvsub2.RowHeadersVisible = False
        Me.dgvsub2.Size = New System.Drawing.Size(258, 225)
        Me.dgvsub2.TabIndex = 8
        '
        'lblSubGridTitle
        '
        Me.lblSubGridTitle.AutoSize = True
        Me.lblSubGridTitle.BackColor = System.Drawing.Color.Transparent
        Me.lblSubGridTitle.Location = New System.Drawing.Point(501, 116)
        Me.lblSubGridTitle.Name = "lblSubGridTitle"
        Me.lblSubGridTitle.Size = New System.Drawing.Size(41, 15)
        Me.lblSubGridTitle.TabIndex = 7
        Me.lblSubGridTitle.Text = "Label1"
        '
        'lblUserName
        '
        Me.lblUserName.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblUserName.AutoSize = True
        Me.lblUserName.BackColor = System.Drawing.Color.Transparent
        Me.lblUserName.Location = New System.Drawing.Point(860, 116)
        Me.lblUserName.Name = "lblUserName"
        Me.lblUserName.Size = New System.Drawing.Size(41, 15)
        Me.lblUserName.TabIndex = 6
        Me.lblUserName.Text = "Label1"
        '
        'lblMainGridTitle
        '
        Me.lblMainGridTitle.AutoSize = True
        Me.lblMainGridTitle.BackColor = System.Drawing.Color.Transparent
        Me.lblMainGridTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMainGridTitle.Location = New System.Drawing.Point(6, 117)
        Me.lblMainGridTitle.Name = "lblMainGridTitle"
        Me.lblMainGridTitle.Size = New System.Drawing.Size(39, 13)
        Me.lblMainGridTitle.TabIndex = 5
        Me.lblMainGridTitle.Text = "Label1"
        '
        'btnApply
        '
        Me.btnApply.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnApply.Enabled = False
        Me.btnApply.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnApply.Location = New System.Drawing.Point(863, 266)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.Size = New System.Drawing.Size(156, 32)
        Me.btnApply.TabIndex = 4
        Me.btnApply.Text = "Apply Sequence"
        Me.btnApply.UseVisualStyleBackColor = True
        '
        'btnAproveSequence
        '
        Me.btnAproveSequence.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAproveSequence.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAproveSequence.Location = New System.Drawing.Point(902, 544)
        Me.btnAproveSequence.Name = "btnAproveSequence"
        Me.btnAproveSequence.Size = New System.Drawing.Size(117, 27)
        Me.btnAproveSequence.TabIndex = 3
        Me.btnAproveSequence.Text = "AproveSequence"
        Me.btnAproveSequence.UseVisualStyleBackColor = True
        '
        'dgvSub
        '
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.AliceBlue
        Me.dgvSub.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvSub.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvSub.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(146, Byte), Integer), CType(CType(164, Byte), Integer), CType(CType(223, Byte), Integer))
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvSub.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.dgvSub.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(210, Byte), Integer), CType(CType(234, Byte), Integer))
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvSub.DefaultCellStyle = DataGridViewCellStyle3
        Me.dgvSub.Enabled = False
        Me.dgvSub.GridColor = System.Drawing.Color.FromArgb(CType(CType(146, Byte), Integer), CType(CType(164, Byte), Integer), CType(CType(223, Byte), Integer))
        Me.dgvSub.Location = New System.Drawing.Point(501, 139)
        Me.dgvSub.Name = "dgvSub"
        Me.dgvSub.ReadOnly = True
        Me.dgvSub.RowHeadersVisible = False
        Me.dgvSub.Size = New System.Drawing.Size(353, 159)
        Me.dgvSub.TabIndex = 1
        '
        'dgvMain
        '
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.White
        Me.dgvMain.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle4
        Me.dgvMain.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dgvMain.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.Color.FromArgb(CType(CType(146, Byte), Integer), CType(CType(164, Byte), Integer), CType(CType(223, Byte), Integer))
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvMain.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle5
        Me.dgvMain.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.Color.White
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvMain.DefaultCellStyle = DataGridViewCellStyle6
        Me.dgvMain.GridColor = System.Drawing.Color.FromArgb(CType(CType(146, Byte), Integer), CType(CType(164, Byte), Integer), CType(CType(223, Byte), Integer))
        Me.dgvMain.Location = New System.Drawing.Point(6, 139)
        Me.dgvMain.Name = "dgvMain"
        Me.dgvMain.ReadOnly = True
        Me.dgvMain.RowHeadersVisible = False
        Me.dgvMain.Size = New System.Drawing.Size(474, 429)
        Me.dgvMain.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.btnBrowse)
        Me.Panel1.Controls.Add(Me.txtFoldername)
        Me.Panel1.Controls.Add(Me.BtnGetDataAndReset)
        Me.Panel1.Controls.Add(Me.lblDir)
        Me.Panel1.Controls.Add(Me.btnSelect)
        Me.Panel1.Controls.Add(Me.BtnReset)
        Me.Panel1.Controls.Add(Me.txtFilename)
        Me.Panel1.Controls.Add(Me.BtnCalculateAndSave)
        Me.Panel1.Controls.Add(Me.DGV_Calculator)
        Me.Panel1.Controls.Add(Me.lblFilePath)
        Me.Panel1.Controls.Add(Me.BtnPreview)
        Me.Panel1.Controls.Add(Me.cmbCategory)
        Me.Panel1.Controls.Add(Me.BtnApplyValues)
        Me.Panel1.Controls.Add(Me.BtnDelete)
        Me.Panel1.Controls.Add(Me.BtnDown)
        Me.Panel1.Controls.Add(Me.BtnUp)
        Me.Panel1.Controls.Add(Me.cmbProcess)
        Me.Panel1.Controls.Add(Me.dgvsub2)
        Me.Panel1.Controls.Add(Me.lblSubGridTitle)
        Me.Panel1.Controls.Add(Me.lblUserName)
        Me.Panel1.Controls.Add(Me.lblMainGridTitle)
        Me.Panel1.Controls.Add(Me.btnApply)
        Me.Panel1.Controls.Add(Me.btnAproveSequence)
        Me.Panel1.Controls.Add(Me.dgvSub)
        Me.Panel1.Controls.Add(Me.dgvMain)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1062, 591)
        Me.Panel1.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(6, 80)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(55, 15)
        Me.Label1.TabIndex = 24
        Me.Label1.Text = "Category"
        '
        'RST_Design1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1062, 591)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MinimumSize = New System.Drawing.Size(1078, 630)
        Me.Name = "RST_Design1"
        Me.Text = "RST_Design1"
        CType(Me.DGV_Calculator, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvsub2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvSub, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvsub2 As DataGridView
    Friend WithEvents lblSubGridTitle As Label
    Friend WithEvents lblUserName As Label
    Friend WithEvents lblMainGridTitle As Label
    Friend WithEvents btnApply As Button
    Friend WithEvents btnAproveSequence As Button
    Friend WithEvents dgvSub As DataGridView
    Friend WithEvents dgvMain As DataGridView
    Friend WithEvents btnBrowse As Button
    Friend WithEvents btnSelect As Button
    Friend WithEvents txtFoldername As TextBox
    Friend WithEvents txtFilename As TextBox
    Friend WithEvents lblDir As Label
    Friend WithEvents lblFilePath As Label
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents cmbProcess As ComboBox
    Friend WithEvents BtnUp As Button
    Friend WithEvents BtnDown As Button
    Friend WithEvents BtnDelete As Button
    Friend WithEvents BtnApplyValues As Button
    Friend WithEvents cmbCategory As ComboBox
    Friend WithEvents BtnPreview As Button
    Friend WithEvents DGV_Calculator As DataGridView
    Friend WithEvents BtnCalculateAndSave As Button
    Friend WithEvents BtnReset As Button
    Friend WithEvents BtnGetDataAndReset As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Panel1 As Panel
End Class
