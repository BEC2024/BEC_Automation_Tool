<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MTC_MTR_ReviewForm
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
        Me.dgvDocumentDetails = New System.Windows.Forms.DataGridView()
        Me.btnExportExcel = New System.Windows.Forms.Button()
        Me.ComboBoxFields = New System.Windows.Forms.ComboBox()
        Me.txtSearch = New System.Windows.Forms.TextBox()
        Me.btnSearchFile = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.btnGetCurrentAssembly = New System.Windows.Forms.Button()
        Me.btnAssemblyCheck = New System.Windows.Forms.Button()
        Me.btnPartCheck = New System.Windows.Forms.Button()
        Me.btnSheetMetalCheck = New System.Windows.Forms.Button()
        Me.btn_Alldata = New System.Windows.Forms.Button()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.FlowLayoutPanel2 = New System.Windows.Forms.FlowLayoutPanel()
        Me.ButtonClose = New System.Windows.Forms.Button()
        Me.ButtonExportExcelMTC = New System.Windows.Forms.Button()
        Me.BtnExportExcelMTR = New System.Windows.Forms.Button()
        Me.lblBaselinePath = New System.Windows.Forms.Label()
        Me.FlowLayoutPanel3 = New System.Windows.Forms.FlowLayoutPanel()
        Me.txtBaseLineDirPath = New System.Windows.Forms.TextBox()
        Me.btnBrowseBaselinePath = New System.Windows.Forms.Button()
        CType(Me.dgvDocumentDetails, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.FlowLayoutPanel1.SuspendLayout()
        Me.FlowLayoutPanel2.SuspendLayout()
        Me.FlowLayoutPanel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvDocumentDetails
        '
        Me.dgvDocumentDetails.AllowUserToAddRows = False
        Me.dgvDocumentDetails.AllowUserToDeleteRows = False
        Me.dgvDocumentDetails.AllowUserToResizeColumns = False
        Me.dgvDocumentDetails.AllowUserToResizeRows = False
        Me.dgvDocumentDetails.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.TableLayoutPanel1.SetColumnSpan(Me.dgvDocumentDetails, 8)
        Me.dgvDocumentDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvDocumentDetails.Location = New System.Drawing.Point(3, 77)
        Me.dgvDocumentDetails.Name = "dgvDocumentDetails"
        Me.dgvDocumentDetails.Size = New System.Drawing.Size(1175, 505)
        Me.dgvDocumentDetails.TabIndex = 8
        '
        'btnExportExcel
        '
        Me.btnExportExcel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportExcel.Location = New System.Drawing.Point(553, 588)
        Me.btnExportExcel.Name = "btnExportExcel"
        Me.btnExportExcel.Size = New System.Drawing.Size(73, 29)
        Me.btnExportExcel.TabIndex = 9
        Me.btnExportExcel.Text = "Export Excel"
        Me.btnExportExcel.UseVisualStyleBackColor = True
        Me.btnExportExcel.Visible = False
        '
        'ComboBoxFields
        '
        Me.ComboBoxFields.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxFields.FormattingEnabled = True
        Me.ComboBoxFields.Location = New System.Drawing.Point(192, 3)
        Me.ComboBoxFields.Name = "ComboBoxFields"
        Me.ComboBoxFields.Size = New System.Drawing.Size(183, 23)
        Me.ComboBoxFields.TabIndex = 37
        '
        'txtSearch
        '
        Me.txtSearch.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.txtSearch.Location = New System.Drawing.Point(3, 3)
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Size = New System.Drawing.Size(183, 23)
        Me.txtSearch.TabIndex = 35
        '
        'btnSearchFile
        '
        Me.btnSearchFile.Image = Global.SolidEdgeApp.My.Resources.Resources.search_16px
        Me.btnSearchFile.Location = New System.Drawing.Point(381, 3)
        Me.btnSearchFile.Name = "btnSearchFile"
        Me.btnSearchFile.Size = New System.Drawing.Size(50, 23)
        Me.btnSearchFile.TabIndex = 36
        Me.btnSearchFile.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 8
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 106.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 234.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 92.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 49.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 175.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 269.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 552.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.FlowLayoutPanel1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.dgvDocumentDetails, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.FlowLayoutPanel2, 7, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.lblBaselinePath, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.FlowLayoutPanel3, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.btnExportExcel, 6, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.txtBaseLineDirPath, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.btnBrowseBaselinePath, 2, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 4
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 37.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 37.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(1181, 620)
        Me.TableLayoutPanel1.TabIndex = 41
        '
        'FlowLayoutPanel1
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.FlowLayoutPanel1, 8)
        Me.FlowLayoutPanel1.Controls.Add(Me.btnGetCurrentAssembly)
        Me.FlowLayoutPanel1.Controls.Add(Me.btnAssemblyCheck)
        Me.FlowLayoutPanel1.Controls.Add(Me.btnPartCheck)
        Me.FlowLayoutPanel1.Controls.Add(Me.btnSheetMetalCheck)
        Me.FlowLayoutPanel1.Controls.Add(Me.btn_Alldata)
        Me.FlowLayoutPanel1.Controls.Add(Me.CheckBox1)
        Me.FlowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.FlowLayoutPanel1.Margin = New System.Windows.Forms.Padding(0)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(1181, 37)
        Me.FlowLayoutPanel1.TabIndex = 45
        '
        'btnGetCurrentAssembly
        '
        Me.btnGetCurrentAssembly.Location = New System.Drawing.Point(3, 3)
        Me.btnGetCurrentAssembly.Name = "btnGetCurrentAssembly"
        Me.btnGetCurrentAssembly.Size = New System.Drawing.Size(171, 29)
        Me.btnGetCurrentAssembly.TabIndex = 7
        Me.btnGetCurrentAssembly.Text = "Get Current Assembly Data"
        Me.btnGetCurrentAssembly.UseVisualStyleBackColor = True
        '
        'btnAssemblyCheck
        '
        Me.btnAssemblyCheck.Location = New System.Drawing.Point(180, 3)
        Me.btnAssemblyCheck.Name = "btnAssemblyCheck"
        Me.btnAssemblyCheck.Size = New System.Drawing.Size(171, 29)
        Me.btnAssemblyCheck.TabIndex = 38
        Me.btnAssemblyCheck.Text = "Assembly Check"
        Me.btnAssemblyCheck.UseVisualStyleBackColor = True
        '
        'btnPartCheck
        '
        Me.btnPartCheck.Location = New System.Drawing.Point(357, 3)
        Me.btnPartCheck.Name = "btnPartCheck"
        Me.btnPartCheck.Size = New System.Drawing.Size(171, 29)
        Me.btnPartCheck.TabIndex = 39
        Me.btnPartCheck.Text = "Part Check"
        Me.btnPartCheck.UseVisualStyleBackColor = True
        '
        'btnSheetMetalCheck
        '
        Me.btnSheetMetalCheck.Location = New System.Drawing.Point(534, 3)
        Me.btnSheetMetalCheck.Name = "btnSheetMetalCheck"
        Me.btnSheetMetalCheck.Size = New System.Drawing.Size(171, 29)
        Me.btnSheetMetalCheck.TabIndex = 40
        Me.btnSheetMetalCheck.Text = "Sheet Metal Check"
        Me.btnSheetMetalCheck.UseVisualStyleBackColor = True
        '
        'btn_Alldata
        '
        Me.btn_Alldata.Location = New System.Drawing.Point(711, 3)
        Me.btn_Alldata.Name = "btn_Alldata"
        Me.btn_Alldata.Size = New System.Drawing.Size(171, 29)
        Me.btn_Alldata.TabIndex = 41
        Me.btn_Alldata.Text = "All Data"
        Me.btn_Alldata.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(888, 8)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(71, 19)
        Me.CheckBox1.TabIndex = 43
        Me.CheckBox1.Text = "SelectAll"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'FlowLayoutPanel2
        '
        Me.FlowLayoutPanel2.Controls.Add(Me.ButtonClose)
        Me.FlowLayoutPanel2.Controls.Add(Me.ButtonExportExcelMTC)
        Me.FlowLayoutPanel2.Controls.Add(Me.BtnExportExcelMTR)
        Me.FlowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FlowLayoutPanel2.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft
        Me.FlowLayoutPanel2.Location = New System.Drawing.Point(629, 585)
        Me.FlowLayoutPanel2.Margin = New System.Windows.Forms.Padding(0)
        Me.FlowLayoutPanel2.Name = "FlowLayoutPanel2"
        Me.FlowLayoutPanel2.Size = New System.Drawing.Size(552, 35)
        Me.FlowLayoutPanel2.TabIndex = 46
        '
        'ButtonClose
        '
        Me.ButtonClose.Location = New System.Drawing.Point(400, 3)
        Me.ButtonClose.Name = "ButtonClose"
        Me.ButtonClose.Size = New System.Drawing.Size(149, 29)
        Me.ButtonClose.TabIndex = 10
        Me.ButtonClose.Text = "Close"
        Me.ButtonClose.UseVisualStyleBackColor = True
        '
        'ButtonExportExcelMTC
        '
        Me.ButtonExportExcelMTC.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonExportExcelMTC.Location = New System.Drawing.Point(223, 3)
        Me.ButtonExportExcelMTC.Name = "ButtonExportExcelMTC"
        Me.ButtonExportExcelMTC.Size = New System.Drawing.Size(171, 29)
        Me.ButtonExportExcelMTC.TabIndex = 42
        Me.ButtonExportExcelMTC.Text = "Export Excel MTC"
        Me.ButtonExportExcelMTC.UseVisualStyleBackColor = True
        '
        'BtnExportExcelMTR
        '
        Me.BtnExportExcelMTR.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnExportExcelMTR.Location = New System.Drawing.Point(46, 3)
        Me.BtnExportExcelMTR.Name = "BtnExportExcelMTR"
        Me.BtnExportExcelMTR.Size = New System.Drawing.Size(171, 29)
        Me.BtnExportExcelMTR.TabIndex = 43
        Me.BtnExportExcelMTR.Text = "Export Excel MTR"
        Me.BtnExportExcelMTR.UseVisualStyleBackColor = True
        '
        'lblBaselinePath
        '
        Me.lblBaselinePath.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lblBaselinePath.AutoSize = True
        Me.lblBaselinePath.Location = New System.Drawing.Point(3, 48)
        Me.lblBaselinePath.Name = "lblBaselinePath"
        Me.lblBaselinePath.Size = New System.Drawing.Size(95, 15)
        Me.lblBaselinePath.TabIndex = 47
        Me.lblBaselinePath.Text = "Baseline Dir Path"
        '
        'FlowLayoutPanel3
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.FlowLayoutPanel3, 5)
        Me.FlowLayoutPanel3.Controls.Add(Me.txtSearch)
        Me.FlowLayoutPanel3.Controls.Add(Me.ComboBoxFields)
        Me.FlowLayoutPanel3.Controls.Add(Me.btnSearchFile)
        Me.FlowLayoutPanel3.Location = New System.Drawing.Point(3, 588)
        Me.FlowLayoutPanel3.Name = "FlowLayoutPanel3"
        Me.FlowLayoutPanel3.Size = New System.Drawing.Size(544, 29)
        Me.FlowLayoutPanel3.TabIndex = 48
        '
        'txtBaseLineDirPath
        '
        Me.txtBaseLineDirPath.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtBaseLineDirPath.Location = New System.Drawing.Point(109, 44)
        Me.txtBaseLineDirPath.Name = "txtBaseLineDirPath"
        Me.txtBaseLineDirPath.Size = New System.Drawing.Size(228, 23)
        Me.txtBaseLineDirPath.TabIndex = 49
        '
        'btnBrowseBaselinePath
        '
        Me.btnBrowseBaselinePath.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrowseBaselinePath.Location = New System.Drawing.Point(343, 42)
        Me.btnBrowseBaselinePath.Name = "btnBrowseBaselinePath"
        Me.btnBrowseBaselinePath.Size = New System.Drawing.Size(86, 26)
        Me.btnBrowseBaselinePath.TabIndex = 50
        Me.btnBrowseBaselinePath.Text = "Browse"
        Me.btnBrowseBaselinePath.UseVisualStyleBackColor = True
        '
        'MTC_MTR_ReviewForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1181, 620)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.KeyPreview = True
        Me.Name = "MTC_MTR_ReviewForm"
        Me.Text = "MTC & MTR Review"
        CType(Me.dgvDocumentDetails, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.FlowLayoutPanel1.PerformLayout()
        Me.FlowLayoutPanel2.ResumeLayout(False)
        Me.FlowLayoutPanel3.ResumeLayout(False)
        Me.FlowLayoutPanel3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents dgvDocumentDetails As DataGridView
    Friend WithEvents btnExportExcel As Button
    Friend WithEvents ComboBoxFields As ComboBox
    Friend WithEvents txtSearch As TextBox
    Friend WithEvents btnSearchFile As Button
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents ButtonClose As Button
    Friend WithEvents ButtonExportExcelMTC As Button
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents FlowLayoutPanel1 As FlowLayoutPanel
    Friend WithEvents btnGetCurrentAssembly As Button
    Friend WithEvents btnAssemblyCheck As Button
    Friend WithEvents btnPartCheck As Button
    Friend WithEvents btnSheetMetalCheck As Button
    Friend WithEvents btn_Alldata As Button
    Friend WithEvents FlowLayoutPanel2 As FlowLayoutPanel
    Friend WithEvents BtnExportExcelMTR As Button
    Friend WithEvents lblBaselinePath As Label
    Friend WithEvents FlowLayoutPanel3 As FlowLayoutPanel
    Friend WithEvents txtBaseLineDirPath As TextBox
    Friend WithEvents btnBrowseBaselinePath As Button
End Class
