<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class VirtualAssemblyStructureForm2
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
        Me.chkAddUserAssembly = New System.Windows.Forms.CheckBox()
        Me.lbl_ReferenceModel = New System.Windows.Forms.Label()
        Me.btnopenfile = New System.Windows.Forms.Button()
        Me.txtfilepath = New System.Windows.Forms.TextBox()
        Me.lblDirectoryLocation = New System.Windows.Forms.Label()
        Me.btnDirectoryPath = New System.Windows.Forms.Button()
        Me.txtDirectoryPath = New System.Windows.Forms.TextBox()
        Me.btnCreateVirtaulAssembly = New System.Windows.Forms.Button()
        Me.lblExcelPath = New System.Windows.Forms.Label()
        Me.btnBrowseExcel = New System.Windows.Forms.Button()
        Me.txtExcelPath = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.SuspendLayout()
        '
        'chkAddUserAssembly
        '
        Me.chkAddUserAssembly.AutoSize = True
        Me.chkAddUserAssembly.Location = New System.Drawing.Point(113, 91)
        Me.chkAddUserAssembly.Name = "chkAddUserAssembly"
        Me.chkAddUserAssembly.Size = New System.Drawing.Size(128, 19)
        Me.chkAddUserAssembly.TabIndex = 42
        Me.chkAddUserAssembly.Text = "Add User Assembly"
        Me.chkAddUserAssembly.UseVisualStyleBackColor = True
        Me.chkAddUserAssembly.Visible = False
        '
        'lbl_ReferenceModel
        '
        Me.lbl_ReferenceModel.AutoSize = True
        Me.lbl_ReferenceModel.Enabled = False
        Me.lbl_ReferenceModel.Location = New System.Drawing.Point(11, 120)
        Me.lbl_ReferenceModel.Name = "lbl_ReferenceModel"
        Me.lbl_ReferenceModel.Size = New System.Drawing.Size(96, 15)
        Me.lbl_ReferenceModel.TabIndex = 39
        Me.lbl_ReferenceModel.Text = "Reference Model"
        Me.lbl_ReferenceModel.Visible = False
        '
        'btnopenfile
        '
        Me.btnopenfile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnopenfile.Enabled = False
        Me.btnopenfile.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnopenfile.Location = New System.Drawing.Point(744, 115)
        Me.btnopenfile.Name = "btnopenfile"
        Me.btnopenfile.Size = New System.Drawing.Size(153, 25)
        Me.btnopenfile.TabIndex = 41
        Me.btnopenfile.Text = "Browse"
        Me.btnopenfile.UseVisualStyleBackColor = True
        Me.btnopenfile.Visible = False
        '
        'txtfilepath
        '
        Me.txtfilepath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtfilepath.Enabled = False
        Me.txtfilepath.Location = New System.Drawing.Point(113, 116)
        Me.txtfilepath.Name = "txtfilepath"
        Me.txtfilepath.Size = New System.Drawing.Size(625, 23)
        Me.txtfilepath.TabIndex = 40
        Me.txtfilepath.Visible = False
        '
        'lblDirectoryLocation
        '
        Me.lblDirectoryLocation.AutoSize = True
        Me.lblDirectoryLocation.Location = New System.Drawing.Point(11, 56)
        Me.lblDirectoryLocation.Name = "lblDirectoryLocation"
        Me.lblDirectoryLocation.Size = New System.Drawing.Size(96, 15)
        Me.lblDirectoryLocation.TabIndex = 36
        Me.lblDirectoryLocation.Text = "Output Directory"
        '
        'btnDirectoryPath
        '
        Me.btnDirectoryPath.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDirectoryPath.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDirectoryPath.Location = New System.Drawing.Point(744, 51)
        Me.btnDirectoryPath.Name = "btnDirectoryPath"
        Me.btnDirectoryPath.Size = New System.Drawing.Size(153, 25)
        Me.btnDirectoryPath.TabIndex = 38
        Me.btnDirectoryPath.Text = "Browse"
        Me.btnDirectoryPath.UseVisualStyleBackColor = True
        '
        'txtDirectoryPath
        '
        Me.txtDirectoryPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDirectoryPath.Location = New System.Drawing.Point(113, 52)
        Me.txtDirectoryPath.Name = "txtDirectoryPath"
        Me.txtDirectoryPath.Size = New System.Drawing.Size(625, 23)
        Me.txtDirectoryPath.TabIndex = 37
        '
        'btnCreateVirtaulAssembly
        '
        Me.btnCreateVirtaulAssembly.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCreateVirtaulAssembly.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCreateVirtaulAssembly.Location = New System.Drawing.Point(883, 554)
        Me.btnCreateVirtaulAssembly.Name = "btnCreateVirtaulAssembly"
        Me.btnCreateVirtaulAssembly.Size = New System.Drawing.Size(168, 30)
        Me.btnCreateVirtaulAssembly.TabIndex = 35
        Me.btnCreateVirtaulAssembly.Text = "Generate Assembly"
        Me.btnCreateVirtaulAssembly.UseVisualStyleBackColor = True
        '
        'lblExcelPath
        '
        Me.lblExcelPath.AutoSize = True
        Me.lblExcelPath.Location = New System.Drawing.Point(11, 26)
        Me.lblExcelPath.Name = "lblExcelPath"
        Me.lblExcelPath.Size = New System.Drawing.Size(99, 15)
        Me.lblExcelPath.TabIndex = 32
        Me.lblExcelPath.Text = "Hedge Excel Path"
        '
        'btnBrowseExcel
        '
        Me.btnBrowseExcel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrowseExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnBrowseExcel.Location = New System.Drawing.Point(744, 21)
        Me.btnBrowseExcel.Name = "btnBrowseExcel"
        Me.btnBrowseExcel.Size = New System.Drawing.Size(153, 25)
        Me.btnBrowseExcel.TabIndex = 34
        Me.btnBrowseExcel.Text = "Browse"
        Me.btnBrowseExcel.UseVisualStyleBackColor = True
        '
        'txtExcelPath
        '
        Me.txtExcelPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtExcelPath.Location = New System.Drawing.Point(113, 22)
        Me.txtExcelPath.Name = "txtExcelPath"
        Me.txtExcelPath.Size = New System.Drawing.Size(625, 23)
        Me.txtExcelPath.TabIndex = 33
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(802, 554)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 30)
        Me.Button1.TabIndex = 43
        Me.Button1.Text = "Close"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'VirtualAssemblyStructureForm2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1063, 591)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.chkAddUserAssembly)
        Me.Controls.Add(Me.lbl_ReferenceModel)
        Me.Controls.Add(Me.btnopenfile)
        Me.Controls.Add(Me.txtfilepath)
        Me.Controls.Add(Me.lblDirectoryLocation)
        Me.Controls.Add(Me.btnDirectoryPath)
        Me.Controls.Add(Me.txtDirectoryPath)
        Me.Controls.Add(Me.btnCreateVirtaulAssembly)
        Me.Controls.Add(Me.lblExcelPath)
        Me.Controls.Add(Me.btnBrowseExcel)
        Me.Controls.Add(Me.txtExcelPath)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.Name = "VirtualAssemblyStructureForm2"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "VirtualAssemblyStructureForm2"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents chkAddUserAssembly As CheckBox
    Friend WithEvents lbl_ReferenceModel As Label
    Friend WithEvents btnopenfile As Button
    Friend WithEvents txtfilepath As TextBox
    Friend WithEvents lblDirectoryLocation As Label
    Friend WithEvents btnDirectoryPath As Button
    Friend WithEvents txtDirectoryPath As TextBox
    Friend WithEvents btnCreateVirtaulAssembly As Button
    Friend WithEvents lblExcelPath As Label
    Friend WithEvents btnBrowseExcel As Button
    Friend WithEvents txtExcelPath As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
End Class
