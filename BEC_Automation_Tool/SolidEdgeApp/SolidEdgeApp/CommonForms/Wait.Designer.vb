<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Wait
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Wait))
        Me.LabelWait = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.LabelMessage = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ButtonStop = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.lblProgressCount = New System.Windows.Forms.Label()
        Me.lblProgressInformation = New System.Windows.Forms.Label()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'LabelWait
        '
        Me.LabelWait.AutoSize = True
        Me.LabelWait.Location = New System.Drawing.Point(14, 111)
        Me.LabelWait.Name = "LabelWait"
        Me.LabelWait.Size = New System.Drawing.Size(177, 15)
        Me.LabelWait.TabIndex = 14
        Me.LabelWait.Text = "This might take several minutes."
        Me.LabelWait.Visible = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Image = Global.SolidEdgeApp.My.Resources.Resources.progress
        Me.PictureBox1.Location = New System.Drawing.Point(88, 53)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(203, 52)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox1.TabIndex = 16
        Me.PictureBox1.TabStop = False
        '
        'LabelMessage
        '
        Me.LabelMessage.AutoSize = True
        Me.LabelMessage.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelMessage.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.LabelMessage.Location = New System.Drawing.Point(14, 14)
        Me.LabelMessage.Name = "LabelMessage"
        Me.LabelMessage.Size = New System.Drawing.Size(87, 21)
        Me.LabelMessage.TabIndex = 17
        Me.LabelMessage.Text = "Please wait"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.Controls.Add(Me.ButtonStop)
        Me.Panel1.Controls.Add(Me.LabelWait)
        Me.Panel1.Controls.Add(Me.TableLayoutPanel1)
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Controls.Add(Me.LabelMessage)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(379, 181)
        Me.Panel1.TabIndex = 18
        '
        'ButtonStop
        '
        Me.ButtonStop.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        'Me.ButtonStop.Image = Global.SolidEdgeApp.My.Resources.Resources.close
        Me.ButtonStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ButtonStop.Location = New System.Drawing.Point(151, 108)
        Me.ButtonStop.Name = "ButtonStop"
        Me.ButtonStop.Size = New System.Drawing.Size(77, 25)
        Me.ButtonStop.TabIndex = 19
        Me.ButtonStop.Text = " STOP"
        Me.ButtonStop.UseVisualStyleBackColor = True
        Me.ButtonStop.Visible = False
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 76.51715!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 23.48285!))
        Me.TableLayoutPanel1.Controls.Add(Me.lblProgressCount, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.lblProgressInformation, 0, 0)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 159)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(3, 43, 3, 3)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(379, 22)
        Me.TableLayoutPanel1.TabIndex = 18
        '
        'lblProgressCount
        '
        Me.lblProgressCount.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblProgressCount.AutoSize = True
        Me.lblProgressCount.Location = New System.Drawing.Point(348, 3)
        Me.lblProgressCount.Margin = New System.Windows.Forms.Padding(3, 3, 7, 3)
        Me.lblProgressCount.Name = "lblProgressCount"
        Me.lblProgressCount.Size = New System.Drawing.Size(24, 15)
        Me.lblProgressCount.TabIndex = 19
        Me.lblProgressCount.Text = "0/0"
        Me.lblProgressCount.Visible = False
        '
        'lblProgressInformation
        '
        Me.lblProgressInformation.AutoSize = True
        Me.lblProgressInformation.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.lblProgressInformation.Location = New System.Drawing.Point(14, 3)
        Me.lblProgressInformation.Margin = New System.Windows.Forms.Padding(14, 3, 3, 3)
        Me.lblProgressInformation.Name = "lblProgressInformation"
        Me.lblProgressInformation.Size = New System.Drawing.Size(118, 15)
        Me.lblProgressInformation.TabIndex = 15
        Me.lblProgressInformation.Text = "Progress information"
        '
        'Wait
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(379, 181)
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MinimumSize = New System.Drawing.Size(395, 197)
        Me.Name = "Wait"
        Me.Opacity = 0.9R
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Please Wait"
        Me.TopMost = True
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents LabelWait As Windows.Forms.Label
    Friend WithEvents PictureBox1 As Windows.Forms.PictureBox
    Friend WithEvents LabelMessage As Windows.Forms.Label
    Friend WithEvents Panel1 As Windows.Forms.Panel
    Friend WithEvents TableLayoutPanel1 As Windows.Forms.TableLayoutPanel
    Friend WithEvents lblProgressInformation As Windows.Forms.Label
    Friend WithEvents lblProgressCount As Windows.Forms.Label
    Friend WithEvents ButtonStop As Windows.Forms.Button
End Class
