﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InterferenceAssemblyForm
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
        Me.btnCheckInterference = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnCheckInterference
        '
        Me.btnCheckInterference.Location = New System.Drawing.Point(13, 13)
        Me.btnCheckInterference.Name = "btnCheckInterference"
        Me.btnCheckInterference.Size = New System.Drawing.Size(106, 23)
        Me.btnCheckInterference.TabIndex = 0
        Me.btnCheckInterference.Text = "Check Interference"
        Me.btnCheckInterference.UseVisualStyleBackColor = True
        '
        'InterferenceAssemblyForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(294, 47)
        Me.Controls.Add(Me.btnCheckInterference)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "InterferenceAssemblyForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Assembly Interference"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnCheckInterference As Button
End Class
