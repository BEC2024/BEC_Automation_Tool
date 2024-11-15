<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class VirtualAssemblyStructureForm
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
        Me.lblExcelPath = New System.Windows.Forms.Label()
        Me.btnBrowseExcel = New System.Windows.Forms.Button()
        Me.txtExcelPath = New System.Windows.Forms.TextBox()
        Me.btnCreateVirtaulAssembly = New System.Windows.Forms.Button()
        Me.lblDirectoryLocation = New System.Windows.Forms.Label()
        Me.btnDirectoryPath = New System.Windows.Forms.Button()
        Me.txtDirectoryPath = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtfilepath = New System.Windows.Forms.TextBox()
        Me.btnopenfile = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblExcelPath
        '
        Me.lblExcelPath.AutoSize = True
        Me.lblExcelPath.Location = New System.Drawing.Point(13, 30)
        Me.lblExcelPath.Name = "lblExcelPath"
        Me.lblExcelPath.Size = New System.Drawing.Size(99, 15)
        Me.lblExcelPath.TabIndex = 21
        Me.lblExcelPath.Text = "Hedge Excel Path"
        '
        'btnBrowseExcel
        '
        Me.btnBrowseExcel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrowseExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnBrowseExcel.Location = New System.Drawing.Point(718, 22)
        Me.btnBrowseExcel.Name = "btnBrowseExcel"
        Me.btnBrowseExcel.Size = New System.Drawing.Size(156, 30)
        Me.btnBrowseExcel.TabIndex = 23
        Me.btnBrowseExcel.Text = "Browse"
        Me.btnBrowseExcel.UseVisualStyleBackColor = True
        '
        'txtExcelPath
        '
        Me.txtExcelPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtExcelPath.Location = New System.Drawing.Point(113, 26)
        Me.txtExcelPath.Name = "txtExcelPath"
        Me.txtExcelPath.Size = New System.Drawing.Size(597, 23)
        Me.txtExcelPath.TabIndex = 22
        '
        'btnCreateVirtaulAssembly
        '
        Me.btnCreateVirtaulAssembly.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCreateVirtaulAssembly.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCreateVirtaulAssembly.Location = New System.Drawing.Point(895, 556)
        Me.btnCreateVirtaulAssembly.Name = "btnCreateVirtaulAssembly"
        Me.btnCreateVirtaulAssembly.Size = New System.Drawing.Size(156, 30)
        Me.btnCreateVirtaulAssembly.TabIndex = 24
        Me.btnCreateVirtaulAssembly.Text = "Generate Assembly"
        Me.btnCreateVirtaulAssembly.UseVisualStyleBackColor = True
        '
        'lblDirectoryLocation
        '
        Me.lblDirectoryLocation.AutoSize = True
        Me.lblDirectoryLocation.Location = New System.Drawing.Point(13, 67)
        Me.lblDirectoryLocation.Name = "lblDirectoryLocation"
        Me.lblDirectoryLocation.Size = New System.Drawing.Size(92, 15)
        Me.lblDirectoryLocation.TabIndex = 25
        Me.lblDirectoryLocation.Text = "Export Directory"
        '
        'btnDirectoryPath
        '
        Me.btnDirectoryPath.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDirectoryPath.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDirectoryPath.Location = New System.Drawing.Point(718, 59)
        Me.btnDirectoryPath.Name = "btnDirectoryPath"
        Me.btnDirectoryPath.Size = New System.Drawing.Size(156, 30)
        Me.btnDirectoryPath.TabIndex = 27
        Me.btnDirectoryPath.Text = "Browse "
        Me.btnDirectoryPath.UseVisualStyleBackColor = True
        '
        'txtDirectoryPath
        '
        Me.txtDirectoryPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDirectoryPath.Location = New System.Drawing.Point(113, 63)
        Me.txtDirectoryPath.Name = "txtDirectoryPath"
        Me.txtDirectoryPath.Size = New System.Drawing.Size(597, 23)
        Me.txtDirectoryPath.TabIndex = 26
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 105)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 15)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "Reference Model"
        '
        'txtfilepath
        '
        Me.txtfilepath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtfilepath.Location = New System.Drawing.Point(113, 101)
        Me.txtfilepath.Name = "txtfilepath"
        Me.txtfilepath.Size = New System.Drawing.Size(597, 23)
        Me.txtfilepath.TabIndex = 29
        '
        'btnopenfile
        '
        Me.btnopenfile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnopenfile.Location = New System.Drawing.Point(718, 97)
        Me.btnopenfile.Name = "btnopenfile"
        Me.btnopenfile.Size = New System.Drawing.Size(156, 30)
        Me.btnopenfile.TabIndex = 30
        Me.btnopenfile.Text = "Browse"
        Me.btnopenfile.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(815, 556)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 30)
        Me.Button1.TabIndex = 31
        Me.Button1.Text = "Close"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'VirtualAssemblyStructureForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1063, 591)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btnopenfile)
        Me.Controls.Add(Me.txtfilepath)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblDirectoryLocation)
        Me.Controls.Add(Me.btnDirectoryPath)
        Me.Controls.Add(Me.txtDirectoryPath)
        Me.Controls.Add(Me.btnCreateVirtaulAssembly)
        Me.Controls.Add(Me.lblExcelPath)
        Me.Controls.Add(Me.btnBrowseExcel)
        Me.Controls.Add(Me.txtExcelPath)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "VirtualAssemblyStructureForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " Virtual Assembly Structure"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblExcelPath As Label
    Friend WithEvents btnBrowseExcel As Button
    Friend WithEvents txtExcelPath As TextBox
    Friend WithEvents btnCreateVirtaulAssembly As Button
    Friend WithEvents lblDirectoryLocation As Label
    Friend WithEvents btnDirectoryPath As Button
    Friend WithEvents txtDirectoryPath As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents txtfilepath As TextBox
    Friend WithEvents btnopenfile As Button
    Friend WithEvents Button1 As Button
End Class
