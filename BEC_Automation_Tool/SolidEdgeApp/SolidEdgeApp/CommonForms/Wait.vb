Imports System.Windows.Forms
Public Class Wait
    Public stopped As Boolean = False

    Public Sub dispose2()
        If Me.InvokeRequired Then
            Me.Invoke(New MethodInvoker(AddressOf dispose2))
        Else
            Me.Dispose()
        End If
    End Sub


    Public Delegate Sub SetWaitMessageDelegate(message As String)

    Public Sub SetWaitMessage(message As String)
        If Me.InvokeRequired Then
            Me.Invoke(New SetWaitMessageDelegate(AddressOf SetWaitMessage), message)
        Else
            Try
                LabelMessage.Text = message
            Catch ex As Exception
            End Try
        End If
    End Sub


    Public Delegate Sub SetProgressInformationVisibilityDelegate(flag As Boolean)

    Public Sub SetProgressInformationVisibility(flag As Boolean)
        If Me.InvokeRequired Then
            Me.Invoke(New SetProgressInformationVisibilityDelegate(AddressOf SetProgressInformationVisibility), flag)
        Else
            Try
                lblProgressInformation.Visible = flag
            Catch ex As Exception
            End Try
        End If
    End Sub

    Public Delegate Sub SetProgressCountVisibilityDelegate(flag As Boolean)

    Public Sub SetProgressCountVisibility(flag As Boolean)
        If Me.InvokeRequired Then
            Me.Invoke(New SetProgressCountVisibilityDelegate(AddressOf SetProgressCountVisibility), flag)
        Else
            Try
                lblProgressCount.Visible = flag
            Catch ex As Exception
            End Try
        End If
    End Sub


    Public Delegate Sub SetProgressInformationMessageDelegate(message As String)

    Public Sub SetProgressInformationMessage(message As String)
        If Me.InvokeRequired Then
            Me.Invoke(New SetProgressInformationMessageDelegate(AddressOf SetProgressInformationMessage), message)
        Else
            Try
                lblProgressInformation.Text = message
            Catch ex As Exception
            End Try
        End If
    End Sub




    Public Delegate Sub SetProgressCountMessageDelegate(message As String)

    Public Sub SetProgressCountMessage(message As String)
        If Me.InvokeRequired Then
            Me.Invoke(New SetProgressCountMessageDelegate(AddressOf SetProgressCountMessage), message)
        Else
            Try
                lblProgressCount.Text = message
            Catch ex As Exception
            End Try
        End If
    End Sub



    Public Delegate Sub SetStopVisibilityDelegate(visible As Boolean)

    Public Sub SetStopVisibility(visible As Boolean)
        If Me.InvokeRequired Then
            Me.Invoke(New SetStopVisibilityDelegate(AddressOf SetStopVisibility), visible)
        Else
            Try
                ButtonStop.Visible = visible
            Catch ex As Exception
            End Try
        End If
    End Sub


    Public Delegate Sub CloseWaitDelegate()
    Public Sub CloseWait()
        If Me.InvokeRequired Then
            Me.Invoke(New SetProgressCountMessageDelegate(AddressOf CloseWait))
        Else
            Try
                Me.Close()
            Catch ex As Exception
            End Try
        End If
    End Sub


    Private Sub Wait_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblProgressInformation.Text = String.Empty
        lblProgressCount.Text = String.Empty
    End Sub

    Private Sub TableLayoutPanel1_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel1.Paint

    End Sub

    Private Sub Wait_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            'If MessageBox.Show("Are you sure to stop processing?", "Stop processing", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
            'stopped = True
            'Me.Close()
            'End If
            StopProcess()
        End If
    End Sub

    Private Sub Wait_SizeChanged(sender As Object, e As EventArgs) Handles MyBase.SizeChanged
        Try
            If Me.Height > 175 Then
                Me.Height = 175
            End If
        Catch ex As Exception
            ' 395, 197
            Me.Width = 395
            Me.Height = 197
        End Try

    End Sub

    Private Sub StopProcess()
        If MessageBox.Show("Are you sure to stop processing?", "Stop processing", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
            stopped = True
        End If
    End Sub

    Private Sub ButtonStop_Click(sender As Object, e As EventArgs) Handles ButtonStop.Click
        StopProcess()
    End Sub
End Class