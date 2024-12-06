Imports WK.Libraries.BetterFolderBrowserNS

Public Class KPI_ReportForm
    Dim a As New KPI_ReportBL
    Dim mtcObj As New MTC_Report
    Dim mtrObj As New MTR_Report
    Dim comobj As New CommonReport
    Dim files As New ArrayList()


#Region "Browse fn"
    Public Sub Browsefile()
        Try
            'If Not IO.File.Exists(txtFilename.Text) Then
            '    MessageBox.Show("Please select valid File")
            'End If
            Using dialog As New OpenFileDialog
                dialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
                If dialog.ShowDialog() = DialogResult.OK Then
                    txtFilename.Text = dialog.FileName
                End If
            End Using
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    Private Sub BtnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        ' browsefile()
        BrowseFolder()
    End Sub
    Private Sub BtnFolderBrowse_Click(sender As Object, e As EventArgs) Handles btnFolderBrowse.Click

    End Sub

    Public Sub BrowseFolder()
        Dim mtcfiles As New ArrayList()
        Dim mtrfiles As New ArrayList()



        If files.Count > 0 Then
            files.Clear()
        End If




#Region "Better folder"
        Try
            Dim reportDir As String = ""

            Try
                Dim BetterFolderBrowser As New BetterFolderBrowser 'With {
                '    .Title = "Select folders",
                '    .RootFolder = "C:\\",
                '    .Multiselect = False
                '}
                If BetterFolderBrowser.ShowDialog() = DialogResult.OK Then
                    reportDir = BetterFolderBrowser.SelectedFolder

                    Dim di As New IO.DirectoryInfo(reportDir)
                    For Each d As String In IO.Directory.GetDirectories(reportDir)

                        Dim di1 As New IO.DirectoryInfo(d)
                        'For Each f As String In IO.Directory.GetFiles(d)

                        '    files.Add(f)

                        'Next

                        Dim aryFi1 As IO.FileInfo() = di1.GetFiles("*.xls*")
                        'Dim fi As IO.FileInfo
                        For Each fi In aryFi1
                            If Not fi.ToString.Contains("$") Or Not fi.ToString.Contains("KPI_Report") Then

                                files.Add(di1.ToString + "\" + fi.ToString)
                                If fi.Name.Contains("MTC") Then
                                    mtcfiles.Add(di1.ToString + "\" + fi.ToString)
                                ElseIf fi.Name.Contains("MTR") Then
                                    mtrfiles.Add(di1.ToString + "\" + fi.ToString)
                                End If

                            End If
                        Next
                    Next

                    Dim aryFi As IO.FileInfo() = di.GetFiles("*.xls*")
                    'Dim fi As IO.FileInfo
                    For Each fi In aryFi
                        If Not fi.ToString.Contains("$") Or fi.ToString.Contains("KPI_Report") Then

                            files.Add(di.ToString + "\" + fi.ToString)
                            If fi.Name.Contains("MTC") Then
                                mtcfiles.Add(di.ToString + "\" + fi.ToString)
                            ElseIf fi.Name.Contains("MTR") Then
                                mtrfiles.Add(di.ToString + "\" + fi.ToString)
                            End If
                        End If
                    Next
                End If
            Catch ex As Exception
                If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
                    Dim path As String = FolderBrowserDialog1.SelectedPath
                    reportDir = path
                End If
            End Try

            If Not IO.Directory.Exists(reportDir) Then
                Exit Sub
            End If

            txtFolder.Text = reportDir

        Catch ex As Exception

        End Try
#End Region

#Region "Comment"
        'If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
        '    txtFolder.Text = FolderBrowserDialog1.SelectedPath

        '    '    Dim di As New IO.DirectoryInfo(txtFolder.Text)
        '    '    Dim aryFi As IO.FileInfo() = di.GetFiles("*.xls*")
        '    '    'Dim fi As IO.FileInfo
        '    '    For Each fi In aryFi
        '    '        If Not fi.ToString.Contains("$") Or Not fi.ToString.Contains("KPI_Report") Then

        '    '            files.Add(di.ToString + "\" + fi.ToString)
        '    '            If fi.Name.Contains("MTC") Then
        '    '                mtcfiles.Add(di.ToString + "\" + fi.ToString)
        '    '            ElseIf fi.Name.Contains("MTR") Then
        '    '                mtrfiles.Add(di.ToString + "\" + fi.ToString)
        '    '            End If

        '    '        End If
        '    '    Next

        '    'Else
        '    '    MsgBox("Please select directory")
        '    'End If


        '    Dim di As New IO.DirectoryInfo(txtFolder.Text)
        '    For Each d As String In IO.Directory.GetDirectories(txtFolder.Text)

        '        Dim di1 As New IO.DirectoryInfo(d)
        '        'For Each f As String In IO.Directory.GetFiles(d)

        '        '    files.Add(f)

        '        'Next

        '        Dim aryFi1 As IO.FileInfo() = di1.GetFiles("*.xls*")
        '        'Dim fi As IO.FileInfo
        '        For Each fi In aryFi1
        '            If Not fi.ToString.Contains("$") Or Not fi.ToString.Contains("KPI_Report") Then

        '                files.Add(di1.ToString + "\" + fi.ToString)
        '                If fi.Name.Contains("MTC") Then
        '                    mtcfiles.Add(di1.ToString + "\" + fi.ToString)
        '                ElseIf fi.Name.Contains("MTR") Then
        '                    mtrfiles.Add(di1.ToString + "\" + fi.ToString)
        '                End If

        '            End If
        '        Next
        '    Next

        '    Dim aryFi As IO.FileInfo() = di.GetFiles("*.xls*")
        '    'Dim fi As IO.FileInfo
        '    For Each fi In aryFi
        '        If Not fi.ToString.Contains("$") Or fi.ToString.Contains("KPI_Report") Then

        '            files.Add(di.ToString + "\" + fi.ToString)
        '            If fi.Name.Contains("MTC") Then
        '                mtcfiles.Add(di.ToString + "\" + fi.ToString)
        '            ElseIf fi.Name.Contains("MTR") Then
        '                mtrfiles.Add(di.ToString + "\" + fi.ToString)
        '            End If
        '        End If
        '    Next

        'Else
        '    MsgBox("Please select directory")
        'End If
#End Region


        mtcObj.files = mtcfiles
        mtrObj.files = mtrfiles

        comobj.files = files
        comobj.MTCfiles = mtcfiles
        comobj.MTRfiles = mtrfiles
        comobj.dir = txtFolder.Text

    End Sub

#End Region

#Region "Generate Report"

    Public Sub GenerateReport()
        'If radMTC.Checked = False And radMTR.Checked = False Then
        '    MsgBox("please Choose Either MTC or MTR")
        'ElseIf radMTC.Checked = True Then
        '    waitStartSave()
        '    MTCReport()
        '    WaitEndSave()
        'ElseIf radMTR.Checked = True Then
        '    waitStartSave()
        '    MTR_Report()
        '    WaitEndSave()
        'End If

        'MTCReport()
        'MsgBox("MTC Report Generated")
        'MTR_Report()
        'MsgBox("MTR Report Generated")

        CommonReport()

    End Sub

    Public Sub CommonReport()
        a.commonReport(comobj)
    End Sub

    Public Sub MTR_Report()
        Try

            WaitStartSave()
            mtrObj.dir = txtFolder.Text
            'mtrObj.files = files
            For mtrObj.i = 0 To mtrObj.files.Count - 1
                mtrObj.excelFilePath = mtrObj.files(mtrObj.i)
                a.MTR_Report(mtrObj)
            Next
            WaitEndSave()

        Catch ex As Exception
            MessageBox.Show($"Error While Fetching MTR Report{vbNewLine}{ex.Message}{ex.StackTrace}", "Message")
            CustomLogUtil.Log("While Fetching MTR Report", ex.Message, ex.StackTrace)
        End Try
    End Sub
    Public Sub MTCReport()
        Try
            WaitStartSave()
            mtcObj.dir = txtFolder.Text
            ' mtcObj.files = files
            'mtcObj.excelFilePath = txtFilename.Text
            For mtcObj.i = 0 To mtcObj.files.Count - 1
                mtcObj.excelFilePath = mtcObj.files(mtcObj.i)
                a.MTCReport(mtcObj)
            Next
            WaitEndSave()
            ' a.CountTotalError(mtcObj)
        Catch ex As Exception
            MessageBox.Show($"Error While Fetching MTC Report{vbNewLine}{ex.Message}{ex.StackTrace}", "Message")
            CustomLogUtil.Log("While Fetching MTC Report", ex.Message, ex.StackTrace)
        End Try

    End Sub


    Private Sub BtnGenerateReport_Click(sender As Object, e As EventArgs) Handles btnGenerateReport.Click
        Try
            GetFilesFromFolder()
            WaitStartSave()
            GenerateReport()
            WaitEndSave()
            MessageBox.Show("KPI_Report successfully created", "Message")
            CustomLogUtil.Heading("KPI_Report successfully created")


            If IO.Directory.Exists(comobj.dir) Then
                Process.Start(comobj.dir)
            End If
            Reset()
            ReleaseObject(comobj)
            ' ActiveForm.Dispose()

        Catch ex As Exception
            MessageBox.Show($"Error While Generating KPI Report{vbNewLine}{ex.Message}{ex.StackTrace}")
            CustomLogUtil.Log("Error While Generating KPI Report", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
#End Region

#Region "common fn"

    Dim waitFormObj As Wait
    Public Sub WaitStartSave()
        '==Processing==
        Dim waitThread As System.Threading.Thread
        waitThread = New System.Threading.Thread(AddressOf LaunchWaitSave)
        waitThread.Start()
        Threading.Thread.Sleep(1000)
        waitFormObj.SetWaitMessage("In progress..")

        waitFormObj.SetProgressInformationVisibility(True)
        waitFormObj.SetProgressInformationMessage("")

        waitFormObj.SetProgressCountVisibility(True)
        waitFormObj.SetProgressCountMessage("0/0")
        '
    End Sub

    Public Sub LaunchWaitSave()
        waitFormObj = New Wait()
        waitFormObj.ShowDialog()
    End Sub

    Public Sub WaitEndSave()
        If waitFormObj IsNot Nothing Then
            waitFormObj.dispose2()
            waitFormObj = Nothing
        End If
    End Sub
    Public Sub Reset()
        txtFolder.Clear()
        chkMTR.Checked = False
    End Sub
    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub


    Private Sub KPI_ReportForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CustomLogUtil.Heading("KPI Form Open.....")
        Me.Text = $"{Me.Text} ({GlobalEntity.Version}) "
        txtFolder.Text = Config.configObj.mtcMtrReportsExportDirLocation

    End Sub

    Public Sub GetFilesFromFolder()

        Dim reportDir As String = txtFolder.Text
        Dim mtcfiles As New ArrayList()

        '28th Oct 2024
        'Dim mtrfiles As New ArrayList()



        If files.Count > 0 Then
            files.Clear()
        End If

        Dim di As New IO.DirectoryInfo(reportDir)
        For Each d As String In IO.Directory.GetDirectories(reportDir)

            Dim di1 As New IO.DirectoryInfo(d)
            'For Each f As String In IO.Directory.GetFiles(d)

            '    files.Add(f)

            'Next

            Dim aryFi1 As IO.FileInfo() = di1.GetFiles("*.xls*")
            'Dim fi As IO.FileInfo
            For Each fi In aryFi1
                If Not fi.ToString.Contains("$") Or Not fi.ToString.Contains("KPI_Report") Then

                    files.Add(di1.ToString + "\" + fi.ToString)
                    If fi.Name.Contains("MTC") Then
                        mtcfiles.Add(di1.ToString + "\" + fi.ToString)

                        '28th Oct 2024
                        'ElseIf fi.Name.Contains("MTR") Then
                        '    mtrfiles.Add(di1.ToString + "\" + fi.ToString)
                    End If

                End If
            Next
        Next

        Dim aryFi As IO.FileInfo() = di.GetFiles("*.xls*")

        'Dim fi As IO.FileInfo

        For Each fi In aryFi
            If Not fi.ToString.Contains("$") Or fi.ToString.Contains("KPI_Report") Then

                files.Add(di.ToString + "\" + fi.ToString)
                If fi.Name.Contains("MTC") Then
                    mtcfiles.Add(di.ToString + "\" + fi.ToString)

                    '28th Oct 2024
                    'ElseIf fi.Name.Contains("MTR") Then
                    '    mtrfiles.Add(di.ToString + "\" + fi.ToString)
                End If
            End If
        Next

        mtcObj.files = mtcfiles

        '28th Oct 2024
        'mtrObj.files = mtrfiles

        comobj.files = files
        comobj.MTCfiles = mtcfiles

        '28th Oct 2024
        'comobj.MTRfiles = mtrfiles
        comobj.dir = txtFolder.Text

    End Sub


#End Region

End Class