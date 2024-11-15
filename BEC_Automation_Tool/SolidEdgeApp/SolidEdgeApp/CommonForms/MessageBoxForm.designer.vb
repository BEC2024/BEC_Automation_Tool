Imports System.Windows.Forms

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MessageBoxForm
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
        Me.components = New System.ComponentModel.Container()
        Me.PanelMessage = New System.Windows.Forms.Panel()
        Me.RichTextBoxMessage = New System.Windows.Forms.RichTextBox()
        Me.PanelFooter = New System.Windows.Forms.Panel()
        Me.ButtonSave = New System.Windows.Forms.Button()
        Me.ButtonClose = New System.Windows.Forms.Button()
        Me.PanelTitleSeparator = New System.Windows.Forms.Panel()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.PanelMessage.SuspendLayout()
        Me.PanelFooter.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelMessage
        '
        Me.PanelMessage.Controls.Add(Me.RichTextBoxMessage)
        Me.PanelMessage.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelMessage.Location = New System.Drawing.Point(0, 0)
        Me.PanelMessage.Name = "PanelMessage"
        Me.PanelMessage.Size = New System.Drawing.Size(373, 167)
        Me.PanelMessage.TabIndex = 2
        '
        'RichTextBoxMessage
        '
        Me.RichTextBoxMessage.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RichTextBoxMessage.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RichTextBoxMessage.Location = New System.Drawing.Point(9, 9)
        Me.RichTextBoxMessage.Name = "RichTextBoxMessage"
        Me.RichTextBoxMessage.Size = New System.Drawing.Size(356, 115)
        Me.RichTextBoxMessage.TabIndex = 1
        Me.RichTextBoxMessage.Text = ""
        '
        'PanelFooter
        '
        Me.PanelFooter.Controls.Add(Me.ButtonSave)
        Me.PanelFooter.Controls.Add(Me.ButtonClose)
        Me.PanelFooter.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PanelFooter.Location = New System.Drawing.Point(0, 131)
        Me.PanelFooter.Name = "PanelFooter"
        Me.PanelFooter.Size = New System.Drawing.Size(373, 36)
        Me.PanelFooter.TabIndex = 1
        '
        'ButtonSave
        '
        Me.ButtonSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonSave.Image = Global.SolidEdgeApp.My.Resources.Resources.csvIcon
        Me.ButtonSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ButtonSave.Location = New System.Drawing.Point(195, 5)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(97, 26)
        Me.ButtonSave.TabIndex = 1
        Me.ButtonSave.Text = "  Save CSV"
        Me.ButtonSave.UseVisualStyleBackColor = True
        Me.ButtonSave.Visible = False
        '
        'ButtonClose
        '
        Me.ButtonClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonClose.Image = Global.SolidEdgeApp.My.Resources.Resources.close
        Me.ButtonClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ButtonClose.Location = New System.Drawing.Point(298, 5)
        Me.ButtonClose.Name = "ButtonClose"
        Me.ButtonClose.Size = New System.Drawing.Size(70, 26)
        Me.ButtonClose.TabIndex = 0
        Me.ButtonClose.Text = "  Close"
        Me.ButtonClose.UseVisualStyleBackColor = True
        '
        'PanelTitleSeparator
        '
        Me.PanelTitleSeparator.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelTitleSeparator.Location = New System.Drawing.Point(0, 0)
        Me.PanelTitleSeparator.Name = "PanelTitleSeparator"
        Me.PanelTitleSeparator.Size = New System.Drawing.Size(373, 2)
        Me.PanelTitleSeparator.TabIndex = 4
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'MessageBoxForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(373, 167)
        Me.Controls.Add(Me.PanelTitleSeparator)
        Me.Controls.Add(Me.PanelFooter)
        Me.Controls.Add(Me.PanelMessage)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.KeyPreview = True
        Me.MinimumSize = New System.Drawing.Size(389, 197)
        Me.Name = "MessageBoxForm"
        Me.Opacity = 0.95R
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Message"
        Me.TopMost = True
        Me.PanelMessage.ResumeLayout(False)
        Me.PanelFooter.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents PanelMessage As Panel
    Friend WithEvents PanelFooter As Panel
    Friend WithEvents RichTextBoxMessage As RichTextBox
    Friend WithEvents ButtonClose As Button
    Friend WithEvents PanelTitleSeparator As Panel
    Friend WithEvents ButtonSave As Button
    Friend WithEvents Timer1 As Timer
End Class
