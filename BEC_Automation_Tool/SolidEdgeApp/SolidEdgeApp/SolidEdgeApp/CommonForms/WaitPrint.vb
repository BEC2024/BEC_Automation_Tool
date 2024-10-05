Public Class WaitPrint

    Public stopped As Boolean = False
    'Private count As Integer = 1
    Private documentNumbers As Integer = 1
    Private allowCoolMove As Boolean = False
    Private myCoolPoint As New Drawing.Point

    Public Sub New()
        InitializeComponent()
        'ProgressBarPrinting.Maximum = docNumbers
        'documentNumbers = docNumbers
        ProgressBarPrinting.Value = 0
        'LabelWait.Text = "Initializing Inventor ..."
    End Sub

    Private Sub Panel1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseDown
        allowCoolMove = True
        myCoolPoint = New Drawing.Point(e.X, e.Y)
        Me.Cursor = Cursors.SizeAll
    End Sub

    Private Sub Panel1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseMove
        If Not allowCoolMove Then
            Return
        End If

        'OvalShape1.Location = New Point(OvalShape1.Location.X + e.X - myCoolPoint.X, OvalShape1.Location.Y + e.Y - myCoolPoint.Y)
        Location = New Drawing.Point(Me.Location.X + e.X - myCoolPoint.X, Me.Location.Y + e.Y - myCoolPoint.Y)

    End Sub
    Private Sub Panel1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseUp
        allowCoolMove = False
        Me.Cursor = Cursors.Default
    End Sub

    Public Delegate Sub DocNoDelegate(ByVal msg As Integer)
    Public Sub setDocNumber(ByVal docNumber As Integer)
        If Me.InvokeRequired Then
            Me.Invoke(New DocNoDelegate(AddressOf setDocNumber), docNumber)
        Else
            Try
                documentNumbers = docNumber
                ProgressBarPrinting.Maximum = docNumber
                PictureBoxWait.Visible = False
                ProgressBarPrinting.Visible = True
                'Console.WriteLine("docNumber  " + docNumber.ToString)
            Catch ex As Exception

            End Try
        End If
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        stopped = True
        Me.Close()
    End Sub

    Public Delegate Sub IncrDelegate(ByVal msg As String)

    Public Sub IncrementPB(ByVal msg As String)

        If Me.InvokeRequired Then
            Me.Invoke(New IncrDelegate(AddressOf IncrementPB), msg)
        Else
            Try
                LabelWait.Text = msg
                ProgressBarPrinting.Value += 1
                ProgressBarPrinting.Refresh()
                'count += 1
                'Console.WriteLine(ProgressBarPrinting.Value.ToString + "    " + msg)
                'If count > documentNumbers Then
                '    Exit Sub
                'End If
                'LabelWait.Text = "Procesing step " + count.ToString + " of " + documentNumbers.ToString

                'ProgressBarPrinting.PerformStep()
                'ProgressBarPrinting.PerformClick()

            Catch ex As Exception
                Exit Sub
            End Try
        End If
    End Sub

 
    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class