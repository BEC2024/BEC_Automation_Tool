<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CopyPartForm
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
        Me.btnCopyPartForm = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.btnUpdateAll = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnCopyPartForm
        '
        Me.btnCopyPartForm.Location = New System.Drawing.Point(12, 34)
        Me.btnCopyPartForm.Name = "btnCopyPartForm"
        Me.btnCopyPartForm.Size = New System.Drawing.Size(176, 44)
        Me.btnCopyPartForm.TabIndex = 0
        Me.btnCopyPartForm.Text = "Copy Part Same Location"
        Me.btnCopyPartForm.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(194, 34)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(196, 44)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Transfer Selected Document"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'btnUpdateAll
        '
        Me.btnUpdateAll.Location = New System.Drawing.Point(396, 34)
        Me.btnUpdateAll.Name = "btnUpdateAll"
        Me.btnUpdateAll.Size = New System.Drawing.Size(196, 44)
        Me.btnUpdateAll.TabIndex = 2
        Me.btnUpdateAll.Text = "Update All Open Documents"
        Me.btnUpdateAll.UseVisualStyleBackColor = True
        '
        'CopyPartForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1063, 591)
        Me.Controls.Add(Me.btnUpdateAll)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btnCopyPartForm)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "CopyPartForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Copy Part"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnCopyPartForm As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents btnUpdateAll As Button
End Class
