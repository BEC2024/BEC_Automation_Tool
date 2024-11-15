Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Timers

Public Class MessageBoxForm

    Private currentMessageType As MessageType
    Private message As String = String.Empty
    Private title As String = String.Empty

    Private errorColor As Color = Color.FromArgb(156, 0, 6)
    Private informationColor As Color = Color.FromArgb(91, 155, 213)
    Private WarningColor As Color = Color.FromArgb(255, 202, 40)
    Private isSaveRequired As Boolean = False
    Dim second As Integer

    Sub New(ByVal title As String, ByVal message As String, ByVal currentMessageType As MessageType, ByVal isSaveRequired As Boolean)

        InitializeComponent()

        Me.title = title
        Me.message = message

        Me.currentMessageType = currentMessageType

        Me.Text = Me.title
        Me.RichTextBoxMessage.Text = Me.message

        Me.isSaveRequired = isSaveRequired

    End Sub

    Public Enum MessageType
        InformationMessage
        ErrorMessage
        WarningMessage
    End Enum

    <DllImport("user32.dll", EntryPoint:="ReleaseCapture")>
    Shared Function ReleaseCapture() As IntPtr
    End Function

    <DllImportAttribute("user32.dll", EntryPoint:="SendMessage")>
    Public Shared Function SendMessage(hWnd As IntPtr, Msg As Integer, wParam As Integer, lParam As Integer) As Integer
    End Function

    Private Sub ButtonClose_Click(sender As Object, e As EventArgs) Handles ButtonClose.Click
        Me.Close()
    End Sub

    Private Sub MessageBoxForm_KeyUp(sender As Object, e As KeyEventArgs) Handles MyBase.KeyUp
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub MessageBoxForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        SetControlProperties()
        SetFormTheme(currentMessageType)
        ButtonSave.Visible = False
        ButtonClose.Visible = False
        ' Application.DoEvents()
        'Dim tmr As New System.Timers.Timer()
        'tmr.Interval = 5000
        'tmr.Enabled = True
        'tmr.Start()
        'AddHandler tmr.Elapsed, AddressOf OnTimedEvent

        'Timer1.Interval = 1000
        'Timer1.Start() 'Timer starts functioning


        Dim x As Integer
        Dim y As Integer
        x = Screen.PrimaryScreen.WorkingArea.Width
        y = Screen.PrimaryScreen.WorkingArea.Height - Me.Height - 10

        Me.ShowInTaskbar = False

        Do Until x = Screen.PrimaryScreen.WorkingArea.Width - Me.Width - 1
            x = x - 1
            Me.Location = New Point(x, y)
            'text visible here

            Application.DoEvents()
        Loop

        Threading.Thread.Sleep(1500)
        Me.Dispose()

    End Sub


    'Private Delegate Sub CloseFormCallback()

    'Private Sub CloseForm()
    '    If InvokeRequired Then
    '        Dim d As New CloseFormCallback(AddressOf CloseForm)
    '        Invoke(d, Nothing)
    '    Else
    '        Close()
    '    End If
    'End Sub

    'Private Sub OnTimedEvent(ByVal sender As Object, ByVal e As ElapsedEventArgs)
    '    CloseForm()
    'End Sub

    Private Sub PanelTitleBar_MouseDown(sender As Object, e As MouseEventArgs)
        Try
            ReleaseCapture()
            SendMessage(Me.Handle, &H112, &HF012, 0)
        Catch ex As Exception
        End Try

    End Sub

    Private Sub SetControlProperties()
        'Form
        Me.BackColor = Color.WhiteSmoke

        'Message
        RichTextBoxMessage.BackColor = Color.WhiteSmoke
        PanelMessage.BackColor = Color.Transparent
        Me.RichTextBoxMessage.ReadOnly = True

        ButtonSave.Visible = isSaveRequired
    End Sub

    Private Sub SetError()
        Me.Icon = My.Resources.errorIcon_48x48

        'Footer
        PanelFooter.BackColor = errorColor

        'Title
        PanelTitleSeparator.BackColor = errorColor

    End Sub

    Private Sub SetFormTheme(ByVal currentMessageType As MessageType)

        If currentMessageType = MessageType.InformationMessage Then
            SetInformation()
        ElseIf currentMessageType = MessageType.ErrorMessage Then
            SetError()
        ElseIf currentMessageType = MessageType.WarningMessage Then
            SetWarning()
        End If

    End Sub

    Private Sub SetInformation()
        Me.Icon = My.Resources.informationIcon_48x48


        'Footer
        PanelFooter.BackColor = informationColor

        'Title
        PanelTitleSeparator.BackColor = informationColor

    End Sub
    Private Sub SetWarning()
        Me.Icon = My.Resources.importantIcon_48x48

        'Footer
        PanelFooter.BackColor = WarningColor

        'Title
        PanelTitleSeparator.BackColor = WarningColor

    End Sub

    Private Sub ButtonSave_Click(sender As Object, e As EventArgs) Handles ButtonSave.Click
        Try
            'Dim dialog As New SaveFileDialog()
            'dialog.Filter = "CSV Files|*.csv"
            'Dim result As DialogResult = dialog.ShowDialog()
            'If result <> DialogResult.OK Then
            '    Return
            'End If

            'Dim filename As String = dialog.FileName

            'IO.File.WriteAllText(filename, RichTextBoxMessage.Text)

            ' CSVUtils.SaveCSV(RichTextBoxMessage.Text)
        Catch ex As Exception
            MessageBox.Show($"Error in save csv{vbNewLine}{vbNewLine} {ex.Message} {vbNewLine} {ex.StackTrace}")
        End Try
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        second = second + 1
        If second >= 3 Then
            Timer1.Stop() 'Timer stops functioning
            Me.Close()
        End If
    End Sub
End Class