<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class GuidelineForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(GuidelineForm))
        Me.btnFlange = New System.Windows.Forms.Button()
        Me.btnHem = New System.Windows.Forms.Button()
        Me.btnLouver = New System.Windows.Forms.Button()
        Me.btnBend = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.btnPrev = New System.Windows.Forms.Button()
        Me.lblProgress = New System.Windows.Forms.Label()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnFlange
        '
        Me.btnFlange.Location = New System.Drawing.Point(14, 44)
        Me.btnFlange.Name = "btnFlange"
        Me.btnFlange.Size = New System.Drawing.Size(160, 37)
        Me.btnFlange.TabIndex = 0
        Me.btnFlange.Text = "Flange"
        Me.btnFlange.UseVisualStyleBackColor = True
        '
        'btnHem
        '
        Me.btnHem.Location = New System.Drawing.Point(14, 88)
        Me.btnHem.Name = "btnHem"
        Me.btnHem.Size = New System.Drawing.Size(160, 37)
        Me.btnHem.TabIndex = 1
        Me.btnHem.Text = "Hem"
        Me.btnHem.UseVisualStyleBackColor = True
        '
        'btnLouver
        '
        Me.btnLouver.Location = New System.Drawing.Point(14, 132)
        Me.btnLouver.Name = "btnLouver"
        Me.btnLouver.Size = New System.Drawing.Size(160, 37)
        Me.btnLouver.TabIndex = 2
        Me.btnLouver.Text = "Louver"
        Me.btnLouver.UseVisualStyleBackColor = True
        '
        'btnBend
        '
        Me.btnBend.Location = New System.Drawing.Point(14, 175)
        Me.btnBend.Name = "btnBend"
        Me.btnBend.Size = New System.Drawing.Size(160, 37)
        Me.btnBend.TabIndex = 3
        Me.btnBend.Text = "Bend"
        Me.btnBend.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.PictureBox1.Location = New System.Drawing.Point(181, 44)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(525, 351)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 4
        Me.PictureBox1.TabStop = False
        '
        'btnNext
        '
        Me.btnNext.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnNext.Location = New System.Drawing.Point(606, 403)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(100, 30)
        Me.btnNext.TabIndex = 5
        Me.btnNext.Text = "Next"
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'btnPrev
        '
        Me.btnPrev.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPrev.Location = New System.Drawing.Point(500, 403)
        Me.btnPrev.Name = "btnPrev"
        Me.btnPrev.Size = New System.Drawing.Size(100, 30)
        Me.btnPrev.TabIndex = 6
        Me.btnPrev.Text = "Prev"
        Me.btnPrev.UseVisualStyleBackColor = True
        '
        'lblProgress
        '
        Me.lblProgress.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblProgress.AutoSize = True
        Me.lblProgress.Location = New System.Drawing.Point(178, 403)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(24, 15)
        Me.lblProgress.TabIndex = 7
        Me.lblProgress.Text = "0/0"
        '
        'GuidelineForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(721, 438)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.btnPrev)
        Me.Controls.Add(Me.btnNext)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.btnBend)
        Me.Controls.Add(Me.btnLouver)
        Me.Controls.Add(Me.btnHem)
        Me.Controls.Add(Me.btnFlange)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "GuidelineForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "  Guideline 1.0.27"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnFlange As Button
    Friend WithEvents btnHem As Button
    Friend WithEvents btnLouver As Button
    Friend WithEvents btnBend As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents btnNext As Button
    Friend WithEvents btnPrev As Button
    Friend WithEvents lblProgress As Label
End Class
