<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Author_Updation
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
        Me.txtEmployeeExcelPath = New System.Windows.Forms.TextBox()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblExcelPath
        '
        Me.lblExcelPath.AutoSize = True
        Me.lblExcelPath.Location = New System.Drawing.Point(8, 25)
        Me.lblExcelPath.Name = "lblExcelPath"
        Me.lblExcelPath.Size = New System.Drawing.Size(107, 13)
        Me.lblExcelPath.TabIndex = 21
        Me.lblExcelPath.Text = "Employee Excel Path"
        '
        'btnBrowseExcel
        '
        Me.btnBrowseExcel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrowseExcel.Enabled = False
        Me.btnBrowseExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnBrowseExcel.Location = New System.Drawing.Point(865, 19)
        Me.btnBrowseExcel.Name = "btnBrowseExcel"
        Me.btnBrowseExcel.Size = New System.Drawing.Size(100, 30)
        Me.btnBrowseExcel.TabIndex = 23
        Me.btnBrowseExcel.Text = "Browse"
        Me.btnBrowseExcel.UseVisualStyleBackColor = True
        Me.btnBrowseExcel.Visible = False
        '
        'txtEmployeeExcelPath
        '
        Me.txtEmployeeExcelPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtEmployeeExcelPath.Enabled = False
        Me.txtEmployeeExcelPath.Location = New System.Drawing.Point(146, 23)
        Me.txtEmployeeExcelPath.Name = "txtEmployeeExcelPath"
        Me.txtEmployeeExcelPath.Size = New System.Drawing.Size(713, 20)
        Me.txtEmployeeExcelPath.TabIndex = 22
        '
        'btnUpdate
        '
        Me.btnUpdate.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.btnUpdate.Location = New System.Drawing.Point(403, 88)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(104, 33)
        Me.btnUpdate.TabIndex = 24
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'Author_Updation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(983, 147)
        Me.Controls.Add(Me.btnUpdate)
        Me.Controls.Add(Me.lblExcelPath)
        Me.Controls.Add(Me.btnBrowseExcel)
        Me.Controls.Add(Me.txtEmployeeExcelPath)
        Me.MaximumSize = New System.Drawing.Size(999, 186)
        Me.MinimumSize = New System.Drawing.Size(999, 186)
        Me.Name = "Author_Updation"
        Me.Text = "Author Updation"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblExcelPath As Label
    Friend WithEvents btnBrowseExcel As Button
    Friend WithEvents txtEmployeeExcelPath As TextBox
    Friend WithEvents btnUpdate As Button
End Class
