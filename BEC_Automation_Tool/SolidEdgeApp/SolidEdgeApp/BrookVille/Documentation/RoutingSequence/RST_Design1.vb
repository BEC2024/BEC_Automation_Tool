Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Drawing.Imaging
Imports WK.Libraries.BetterFolderBrowserNS

Public Class RST_Design1
    Dim bs As New BindingSource()
    Dim Maindt As DataTable
    Dim Subdt As New DataTable
    Dim subdt2 As New DataTable
    Dim resSave As New RST_BL
    Dim rstObj As New RountingSequenceClass
    Dim order As New ArrayList()
    Dim ProdTime As New ArrayList()
    Dim MoveTime As New ArrayList()
    'Dim filepath As String = "C:\Users\pratikg\Desktop\SolidEdge.jpg"
#Region "Wait"
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
#End Region

    Private Sub RST_Design1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CustomLogUtil.Heading("Routing Sequence Form Open.....")
        Lbldeclaration()
        CmbCategoryItems()
        Me.Text = $"{Me.Text} ({GlobalEntity.Version}) "
        SetConfigOutputDirectory()
        'SetConfigExcelPath()
    End Sub
    Public Sub SetConfigOutputDirectory()
        txtFoldername.Text = Config.configObj.RoutingSequenceOutputDirectory
    End Sub
    Public Sub SetConfigExcelPath()
        txtFilename.Text = Config.configObj.RoutingSequenceExcelPath
        rstObj.excelFilepath = txtFilename.Text
    End Sub

    Private Sub BtnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
        SelectFile(rstObj)
    End Sub
    Public Sub ShowDgvMain(rstObj As RountingSequenceClass)
        'If rstObj.Maindt.Rows.Count = 0 And rstObj.Maindt.Columns.Count = 0 Then
        '    ShowExistingMtcReport(rstObj)
        'End If

        dgvMain.DataSource = rstObj.Maindt
        dgvMain.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
        dgvMain.Columns(0).Width = 100
        For i = 0 To dgvMain.Rows.Count - 1
            If dgvMain.Rows(i).Cells(0).Value = "ProdTime" Or dgvMain.Rows(i).Cells(0).Value = "MoveTime" Then
                dgvMain.Rows(i).DefaultCellStyle.BackColor = Color.AliceBlue
            Else
                dgvMain.Rows(i).DefaultCellStyle.BackColor = Color.LightSteelBlue
            End If
        Next
    End Sub

    Public Sub ShowDgvSub(rstObj As RountingSequenceClass)
        'Dim selectedGR As DataGridViewRow
        'selectedGR = dgvMain.Rows(e.RowIndex)

        'dgvSub.ColumnCount = 2
        'dgvSub.Columns(0).Name = "ORDER#"
        'dgvSub.Columns(1).Name = "WC#"

        'Dim cmb As New DataGridViewComboBoxColumn()
        'cmb.HeaderText = "PROCESS#"
        'cmb.Name = "PROCESS#"
        'cmb.MaxDropDownItems = 4
        'cmb.Items.Add("Nesting")
        'cmb.Items.Add("Burn")
        'cmb.Items.Add("Grind")
        'cmb.Items.Add("Flange Bend")
        'dgvSub.Columns.Add(cmb)

        'Dim row As String() = New String() {"10", "1000"}
        'dgvSub.Rows.Add(row)
        'row = New String() {"20", "2000"}
        'dgvSub.Rows.Add(row)
        'row = New String() {"30", "3000"}
        'dgvSub.Rows.Add(row)
        'row = New String() {"40", "4000"}
        'dgvSub.Rows.Add(row)
        'row = New String() {"50", "4000"}
        'dgvSub.Rows.Add(row)
        'row = New String() {"60", "4000"}
        'dgvSub.Rows.Add(row)
        'row = New String() {"70", "4000"}
        'dgvSub.Rows.Add(row)
        'row = New String() {"80", "4000"}
        'dgvSub.Rows.Add(row)






        rstObj.dt2.Columns.Clear()
        'rstObj.dt2.Columns.Add(New DataColumn("ORDER#"))
        rstObj.dt2.Columns.Add(New DataColumn("PROCESS#"))
        rstObj.dt2.Columns.Add(New DataColumn("WC#"))
        rstObj.dt2.Columns.Add(New DataColumn("PTIME#"))
        rstObj.dt2.Columns.Add(New DataColumn("MTIME#"))




        bs.DataSource = rstObj.dt2

        dgvSub.DataSource = bs
        dgvSub.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
        dgvSub.Columns(0).Width = 100

    End Sub

    Public Sub Set_dt2data()
        rstObj.dt2.Clear()
        order.Clear()
        ProdTime.Clear()
        MoveTime.Clear()
        Try

            For k = 0 To rstObj.Maindt.Rows.Count - 1 Step 3

                If rstObj.PartName = rstObj.Maindt(k)(0) Then
                    ' MsgBox(rstObj.Maindt(i)(5).Texts.ToString)

                    For j = 1 To rstObj.Maindt.Columns.Count - 1
                        If Not dgvMain.Rows(k).Cells(j).Value.ToString = "" Then
                            ' MsgBox(dgvMain.Rows(i).Cells(j).Value.ToString)
                            order.Add(rstObj.Maindt(k)(j))
                        End If
                        If Not dgvMain.Rows(k + 1).Cells(j).Value.ToString = "" Then
                            ProdTime.Add(rstObj.Maindt(k + 1)(j))
                        End If
                        If Not dgvMain.Rows(k + 2).Cells(j).Value.ToString = "" Then
                            MoveTime.Add(rstObj.Maindt(k + 2)(j))
                        End If
                    Next

                    Dim count = order.Count
                    'For k = 1 To order.Count
                    '    If order(k) = "" Then
                    '        count = k
                    '        Exit For
                    '    End If
                    'Next


                    Dim row1 As DataRow = rstObj.dt2.NewRow()
                    rstObj.dt2.Rows.Add(row1)
                    'row1(0) = "10"

                    If count >= 1 Then
                        For i = 0 To rstObj.dtProcess.Rows.Count - 1
                            If order(0) = rstObj.dtProcess(i)(1) Then
                                'cmbProcess.Items.Remove(rstObj.dtProcess(i)(1))
                                cmbProcess.Items.Remove(rstObj.dtProcess(i)(0))
                                row1(0) = rstObj.dtProcess(i)(0)
                                ' MsgBox(rstObj.dtProcess(i)(0))
                            End If
                        Next
                        row1(1) = order(0)
                        row1(2) = ProdTime(0)
                        row1(3) = MoveTime(0)
                    End If



                    Dim row2 As DataRow = rstObj.dt2.NewRow()
                    rstObj.dt2.Rows.Add(row2)
                    'row2(0) = "20"
                    If count >= 2 Then
                        For i = 0 To rstObj.dtProcess.Rows.Count - 1
                            If order(1) = rstObj.dtProcess(i)(1) Then
                                cmbProcess.Items.Remove(rstObj.dtProcess(i)(0))
                                'cmbProcess.Items.Remove(rstObj.dtProcess(i)(1))
                                row2(0) = rstObj.dtProcess(i)(0)
                                ' MsgBox(rstObj.dtProcess(i)(0))
                            End If
                        Next

                        row2(1) = order(1)
                        row2(2) = ProdTime(1)
                        row2(3) = MoveTime(1)
                    End If


                    Dim row3 As DataRow = rstObj.dt2.NewRow()
                    rstObj.dt2.Rows.Add(row3)
                    'row3(0) = "30"
                    If count >= 3 Then
                        For i = 0 To rstObj.dtProcess.Rows.Count - 1
                            If order(2) = rstObj.dtProcess(i)(1) Then
                                cmbProcess.Items.Remove(rstObj.dtProcess(i)(0))
                                'cmbProcess.Items.Remove(rstObj.dtProcess(i)(1))
                                row3(0) = rstObj.dtProcess(i)(0)
                                ' MsgBox(rstObj.dtProcess(i)(0))
                            End If
                        Next
                        row3(1) = order(2)
                        row3(2) = ProdTime(2)
                        row3(3) = MoveTime(2)
                    End If



                    Dim row4 As DataRow = rstObj.dt2.NewRow()
                    rstObj.dt2.Rows.Add(row4)
                    'row4(0) = "40"
                    If count >= 4 Then
                        For i = 0 To rstObj.dtProcess.Rows.Count - 1
                            If order(3) = rstObj.dtProcess(i)(1) Then
                                cmbProcess.Items.Remove(rstObj.dtProcess(i)(0))
                                'cmbProcess.Items.Remove(rstObj.dtProcess(i)(1))
                                row4(0) = rstObj.dtProcess(i)(0)
                                ' MsgBox(rstObj.dtProcess(i)(0))
                            End If
                        Next
                        row4(1) = order(3)
                        row4(2) = ProdTime(3)
                        row4(3) = MoveTime(3)
                    End If


                    Dim row5 As DataRow = rstObj.dt2.NewRow()
                    rstObj.dt2.Rows.Add(row5)
                    'row5(0) = "50"
                    If count >= 5 Then
                        For i = 0 To rstObj.dtProcess.Rows.Count - 1
                            If order(4) = rstObj.dtProcess(i)(1) Then
                                cmbProcess.Items.Remove(rstObj.dtProcess(i)(0))
                                'cmbProcess.Items.Remove(rstObj.dtProcess(i)(1))
                                row5(0) = rstObj.dtProcess(i)(0)
                                ' MsgBox(rstObj.dtProcess(i)(0))
                            End If
                        Next

                        row5(1) = order(4)
                        row5(2) = ProdTime(4)
                        row5(3) = MoveTime(4)
                    End If


                    Dim row6 As DataRow = rstObj.dt2.NewRow()
                    rstObj.dt2.Rows.Add(row6)
                    'row6(0) = "60"
                    If count >= 6 Then
                        For i = 0 To rstObj.dtProcess.Rows.Count - 1
                            If order(5) = rstObj.dtProcess(i)(1) Then
                                cmbProcess.Items.Remove(rstObj.dtProcess(i)(0))
                                'cmbProcess.Items.Remove(rstObj.dtProcess(i)(1))
                                row6(0) = rstObj.dtProcess(i)(0)
                                ' MsgBox(rstObj.dtProcess(i)(0))
                            End If
                        Next
                        row6(1) = order(5)
                        row6(2) = ProdTime(5)
                        row6(3) = MoveTime(5)
                    End If


                    Dim row7 As DataRow = rstObj.dt2.NewRow()
                    rstObj.dt2.Rows.Add(row7)
                    'row7(0) = "70"
                    If count >= 7 Then
                        For i = 0 To rstObj.dtProcess.Rows.Count - 1
                            If order(6) = rstObj.dtProcess(i)(1) Then
                                cmbProcess.Items.Remove(rstObj.dtProcess(i)(0))
                                'cmbProcess.Items.Remove(rstObj.dtProcess(i)(1))
                                row7(0) = rstObj.dtProcess(i)(0)
                                ' MsgBox(rstObj.dtProcess(i)(0))
                            End If
                        Next

                        row7(1) = order(6)
                        row7(2) = ProdTime(6)
                        row7(3) = MoveTime(6)
                    End If


                    Dim row8 As DataRow = rstObj.dt2.NewRow()
                    rstObj.dt2.Rows.Add(row8)
                    'row8(0) = "80"
                    If count >= 8 Then
                        For i = 0 To rstObj.dtProcess.Rows.Count - 1
                            If order(7) = rstObj.dtProcess(i)(1) Then
                                cmbProcess.Items.Remove(rstObj.dtProcess(i)(0))
                                'cmbProcess.Items.Remove(rstObj.dtProcess(i)(1))
                                row8(0) = rstObj.dtProcess(i)(0)
                                ' MsgBox(rstObj.dtProcess(i)(0))
                            End If
                        Next

                        row8(1) = order(7)
                        row8(2) = ProdTime(7)
                        row8(3) = MoveTime(7)
                    End If


                    Dim row9 As DataRow = rstObj.dt2.NewRow()
                    rstObj.dt2.Rows.Add(row9)
                    'row9(0) = "90"
                    If count >= 9 Then
                        For i = 0 To rstObj.dtProcess.Rows.Count - 1
                            If order(8) = rstObj.dtProcess(i)(1) Then
                                cmbProcess.Items.Remove(rstObj.dtProcess(i)(0))
                                'cmbProcess.Items.Remove(rstObj.dtProcess(i)(1))
                                row9(0) = rstObj.dtProcess(i)(0)
                                ' MsgBox(rstObj.dtProcess(i)(0))
                            End If
                        Next

                        row9(1) = order(8)
                        row9(2) = ProdTime(8)
                        row9(3) = MoveTime(8)
                    Else
                        row9(0) = ""
                        row9(1) = ""
                        row9(2) = ""
                        row9(3) = ""
                    End If




                    Dim row10 As DataRow = rstObj.dt2.NewRow()
                    rstObj.dt2.Rows.Add(row10)
                    'row10(0) = "100"
                    If count = 10 Then
                        For i = 0 To rstObj.dtProcess.Rows.Count - 1
                            If order(9) = rstObj.dtProcess(i)(1) Then
                                cmbProcess.Items.Remove(rstObj.dtProcess(i)(0))
                                'cmbProcess.Items.Remove(rstObj.dtProcess(i)(1))
                                row10(0) = rstObj.dtProcess(i)(0)
                                ' MsgBox(rstObj.dtProcess(i)(0))
                            End If
                        Next

                        row10(1) = order(9)
                        row10(2) = ProdTime(9)
                        row10(3) = MoveTime(9)
                    Else
                        row10(0) = ""
                        row10(1) = ""
                        row10(2) = ""
                        row10(3) = ""
                    End If

                End If

            Next

            bs.DataSource = rstObj.dt2

            dgvSub.DataSource = bs
            dgvSub.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
            dgvSub.Columns(0).Width = 100
        Catch ex As Exception
            MessageBox.Show($"Error While Setting DataTable2 Data {vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}", "Message")
            CustomLogUtil.Log("While Setting DataTable2 Data ", ex.Message, ex.StackTrace)
        End Try
    End Sub
    Public Sub ShowDgvSub2(rstObj As RountingSequenceClass)

        Try
            Dim CMB_HoleFeature As New DataGridViewComboBoxCell
            CMB_HoleFeature.Items.Add("TRUE")
            CMB_HoleFeature.Items.Add("FALSE")

            Dim CMB_Louvers As New DataGridViewComboBoxCell
            CMB_Louvers.Items.Add("TRUE")
            CMB_Louvers.Items.Add("FALSE")

            Dim CMB_HemBeadsGuesset As New DataGridViewComboBoxCell
            CMB_HemBeadsGuesset.Items.Add("TRUE")
            CMB_HemBeadsGuesset.Items.Add("FALSE")

            Dim CMB_PerforatedOrExpanded As New DataGridViewComboBoxCell
            CMB_PerforatedOrExpanded.Items.Add("TRUE")
            CMB_PerforatedOrExpanded.Items.Add("FALSE")

            Dim CMB_HoleFit As New DataGridViewComboBoxCell
            CMB_HoleFit.Items.Add("Loose")
            CMB_HoleFit.Items.Add("Close")
            CMB_HoleFit.Items.Add("Normal")
            CMB_HoleFit.Items.Add("Nominal")
            CMB_HoleFit.Items.Add("Clearance")
            CMB_HoleFit.Items.Add("Transitional")
            CMB_HoleFit.Items.Add("Press")
            CMB_HoleFit.Items.Add("Exact")


            dgvsub2.DataSource = rstObj.dt3
            dgvsub2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
            dgvsub2.Columns(0).Width = 100
            For i = 0 To dgvsub2.Rows.Count - 1
                If dgvsub2.Rows(i).Cells(0).Value = "Hole Feature" Then
                    dgvsub2.Rows(i).Cells(1) = CMB_HoleFeature
                End If
                If dgvsub2.Rows(i).Cells(0).Value = "Louvers" Then
                    dgvsub2.Rows(i).Cells(1) = CMB_Louvers
                End If
                If dgvsub2.Rows(i).Cells(0).Value = "Hem/beads/guesset" Then
                    dgvsub2.Rows(i).Cells(1) = CMB_HemBeadsGuesset
                End If
                If dgvsub2.Rows(i).Cells(0).Value = "Perforated or Expanded" Then
                    dgvsub2.Rows(i).Cells(1) = CMB_PerforatedOrExpanded
                End If
                If dgvsub2.Rows(i).Cells(0).Value = "Hole Fit" Then
                    dgvsub2.Rows(i).Cells(1) = CMB_HoleFit
                End If
            Next

        Catch ex As Exception
            MessageBox.Show($"Error in add and fill combo in grid {vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}", "Message")
            CustomLogUtil.Log("While adding combobox Items and filling data into grid", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Public Sub Lbldeclaration()
        lblMainGridTitle.Text = rstObj.ProjectName + " || CATEGORY-" + rstObj.CategoryName
        lblMainGridTitle.Font = New Font("Segoe UI", 8,
                    FontStyle.Bold)
        lblMainGridTitle.Size = New System.Drawing.Size(494, 13)

        lblSubGridTitle.Text = rstObj.PartName + " - " + rstObj.MaterialDescription
        lblSubGridTitle.Font = New Font("Segoe UI", 10,
                    FontStyle.Bold)
        lblSubGridTitle.Size = New System.Drawing.Size(494, 13)


        If rstObj.user = Nothing Then
            rstObj.user = ""
        End If
        lblUserName.Text = "USER : " + rstObj.user.ToUpper
        lblUserName.Font = New Font("Segoe UI", 10,
                   FontStyle.Bold)
        lblUserName.Size = New System.Drawing.Size(80, 13)




    End Sub
    Public Sub SelectCategory(rstObj As RountingSequenceClass)
        If cmbCategory.SelectedItem = "" Then
            MessageBox.Show("Please select  Category", "Message")
            'MsgBox("Please select directory")
        End If
    End Sub
    Public Sub SelectFile(rstObj As RountingSequenceClass)
        Try
            'If Not IO.File.Exists(txtFilename.Text) Then
            '    MessageBox.Show("Please select valid File")
            'End If

            Using dialog As New OpenFileDialog
                dialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
                If dialog.ShowDialog() = DialogResult.OK Then
                    txtFilename.Text = dialog.FileName

                    'If txtFoldername.Text = "" Then
                    '    MessageBox.Show("Please browse directory", "Message")
                    'End If

                End If
            End Using
            rstObj.excelFilepath = txtFilename.Text
        Catch ex As Exception
            MessageBox.Show("Error While selecting File")
        End Try
    End Sub

    Private Sub BtnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        BroseFolder(rstObj)
        Dim strFileName As String = rstObj.dir + "\MTC_RST_Report.xlsx"
        'If rstObj.dir = "" Then
        '    BroseFolder(rstObj)
        'ElseIf System.IO.File.Exists(strFileName) And rstObj.excelFilepath = "" Then
        '    ShowAllDGV()
        'ElseIf Not rstObj.dir = "" And Not rstObj.excelFilepath = "" Then
        '    mtcReport(rstObj)
        '    ShowAllDGV()
        'End If




    End Sub
    Public Sub BroseFolder(rstObj As RountingSequenceClass)


        'If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
        '    txtFoldername.Text = FolderBrowserDialog1.SelectedPath
        '    rstObj.dir = txtFoldername.Text
        'Else


        'End If


        Dim BetterFolderBrowser As New BetterFolderBrowser With {
            .Title = "Select Destination Folder",
            .Multiselect = False
        }
        If BetterFolderBrowser.ShowDialog() = DialogResult.OK Then
            txtFoldername.Text = BetterFolderBrowser.SelectedFolder
            rstObj.dir = txtFoldername.Text
        End If

    End Sub
    Public Sub MtcReport(rstObj As RountingSequenceClass)
        Try
            resSave.MtcReport(rstObj)
        Catch ex As Exception
            MessageBox.Show($"Error While Creating MTC Report {vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}", "Message")
            CustomLogUtil.Log("While Creating MTC Report", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Public Sub ShowAllDGV()
        Try
            ShowDgvMain(rstObj)
            ShowDgvSub(rstObj)
        Catch ex As Exception
            MessageBox.Show($"Error While Setting All Grid Data {vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}", "Message")
            CustomLogUtil.Log("While Setting All Grid Data", ex.Message, ex.StackTrace)
        End Try


        ' showDgvSub2(rstObj)
    End Sub
    Public Sub ShowExistingMtcReport(rstObj As RountingSequenceClass)
        resSave.ShowExistingMtcReport(rstObj)
    End Sub

    Public Sub Get_dt3data(rstObj As RountingSequenceClass)
        Try
            resSave.Get_dt3data(rstObj)
        Catch ex As Exception
            MessageBox.Show($"Error While Getting DataTable3 data {vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}", "Message")
            CustomLogUtil.Log("While Getting DataTable3 data", ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub DgvMain_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvMain.CellClick

        DGV_Calculator.Columns.Clear()
        DGV_Calculator.Rows.Clear()
        dgvSub.Enabled = False
        BtnApplyValues.Enabled = True
        BtnApplyValues.Visible = True
        BtnPreview.Enabled = True
        CmbProcessItems(rstObj)
        Try

            rstObj.dt2.Clear()
            Dim selectedGR As DataGridViewRow
            Dim ProdTimeGR As DataGridViewRow
            Dim MoveTimeGR As DataGridViewRow
            selectedGR = dgvMain.Rows(e.RowIndex)
            ProdTimeGR = dgvMain.Rows(e.RowIndex + 1)
            MoveTimeGR = dgvMain.Rows(e.RowIndex + 2)


            rstObj.dgvmainRowIndex = selectedGR
            rstObj.dgvmainRowIndexMoveTIME = MoveTimeGR
            rstObj.dgvmainRowIndexProdTIME = ProdTimeGR
            rstObj.PartName = selectedGR.Cells(0).Value.ToString
            'MsgBox(selectedGR.Cells(0).Value.ToString)

            order.Clear()
            ProdTime.Clear()
            MoveTime.Clear()



            If Not rstObj.dgvmainRowIndex.Cells(0).Value() = "ProdTime" And Not rstObj.dgvmainRowIndex.Cells(0).Value() = "MoveTime" Then


                Try


                    For i = 1 To rstObj.dgvmainRowIndex.Cells.Count - 1
                        If Not rstObj.dgvmainRowIndex.Cells(i).Value.ToString = "" Then
                            order.Add(rstObj.dgvmainRowIndex.Cells(i).Value)
                        End If
                    Next

                    For i = 1 To rstObj.dgvmainRowIndexProdTIME.Cells.Count - 1
                        If Not rstObj.dgvmainRowIndexProdTIME.Cells(i).Value.ToString = "" And Not rstObj.dgvmainRowIndexProdTIME.Cells(i).Value.ToString = "FALSE" Then
                            ProdTime.Add(rstObj.dgvmainRowIndexProdTIME.Cells(i).Value)
                        End If
                    Next

                    For i = 1 To rstObj.dgvmainRowIndexMoveTIME.Cells.Count - 1
                        If Not rstObj.dgvmainRowIndexMoveTIME.Cells(i).Value.ToString = "" And Not rstObj.dgvmainRowIndexMoveTIME.Cells(i).Value.ToString = "FALSE" Then
                            MoveTime.Add(rstObj.dgvmainRowIndexMoveTIME.Cells(i).Value)
                        End If
                    Next


                    Dim count = order.Count

                    Dim row1 As DataRow = rstObj.dt2.NewRow()
                    rstObj.dt2.Rows.Add(row1)
                    'row1(0) = "10"

                    If count >= 1 Then
                        For i = 0 To rstObj.dtProcess.Rows.Count - 1
                            If order(0) = rstObj.dtProcess(i)(1) Then
                                'cmbProcess.Items.Remove(rstObj.dtProcess(i)(1))
                                cmbProcess.Items.Remove(rstObj.dtProcess(i)(0))
                                row1(0) = rstObj.dtProcess(i)(0)
                                ' MsgBox(rstObj.dtProcess(i)(0))
                            End If
                        Next
#Region "Comment"
                        'If (order(0) = rstObj.ten) Then
                        '    cmbProcess.Items.Remove(rstObj.process_ten)
                        '    row1(0) = rstObj.process_ten
                        'ElseIf (order(0) = rstObj.twenty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_twenty)
                        '    row1(0) = rstObj.process_twenty
                        'ElseIf (order(0) = rstObj.thirty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_thirty)
                        '    row1(0) = rstObj.process_thirty
                        'ElseIf (order(0) = rstObj.fourty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fourty)
                        '    row1(0) = rstObj.process_fourty
                        'ElseIf (order(0) = rstObj.fifty1) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty1)
                        '    row1(0) = rstObj.process_fifty1
                        'ElseIf (order(0) = rstObj.fifty2) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty2)
                        '    row1(0) = rstObj.process_fifty2
                        'ElseIf (order(0) = rstObj.fifty3) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty3)
                        '    row1(0) = rstObj.process_fifty3
                        'ElseIf (order(0) = rstObj.fifty4) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty4)
                        '    row1(0) = rstObj.process_fifty4
                        'ElseIf (order(0) = rstObj.sixty1) Then
                        '    cmbProcess.Items.Remove(rstObj.process_sixty1)
                        '    row1(0) = rstObj.process_sixty1
                        'ElseIf (order(0) = rstObj.sixty2) Then
                        '    cmbProcess.Items.Remove(rstObj.process_sixty2)
                        '    row1(0) = rstObj.process_sixty2
                        'ElseIf (order(0) = rstObj.seventy) Then
                        '    cmbProcess.Items.Remove(rstObj.process_seventy)
                        '    row1(0) = rstObj.process_seventy
                        'ElseIf (order(0) = rstObj.eighty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_eighty)
                        '    row1(0) = rstObj.process_eighty
                        'End If
#End Region
                        row1(1) = order(0)
                        row1(2) = ProdTime(0)
                        row1(3) = MoveTime(0)
                    End If



                    Dim row2 As DataRow = rstObj.dt2.NewRow()
                    rstObj.dt2.Rows.Add(row2)
                    'row2(0) = "20"
                    If count >= 2 Then
                        For i = 0 To rstObj.dtProcess.Rows.Count - 1
                            If order(1) = rstObj.dtProcess(i)(1) Then
                                cmbProcess.Items.Remove(rstObj.dtProcess(i)(0))
                                'cmbProcess.Items.Remove(rstObj.dtProcess(i)(1))
                                row2(0) = rstObj.dtProcess(i)(0)
                                ' MsgBox(rstObj.dtProcess(i)(0))
                            End If
                        Next
#Region "comment"
                        'If (order(1) = rstObj.ten) Then
                        '    cmbProcess.Items.Remove(rstObj.process_ten)
                        '    row2(0) = rstObj.process_ten
                        'ElseIf (order(1) = rstObj.twenty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_twenty)
                        '    row2(0) = rstObj.process_twenty
                        'ElseIf (order(1) = rstObj.thirty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_thirty)
                        '    row2(0) = rstObj.process_thirty
                        'ElseIf (order(1) = rstObj.fourty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fourty)
                        '    row2(0) = rstObj.process_fourty
                        'ElseIf (order(1) = rstObj.fifty1) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty1)
                        '    row2(0) = rstObj.process_fifty1
                        'ElseIf (order(1) = rstObj.fifty2) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty2)
                        '    row2(0) = rstObj.process_fifty2
                        'ElseIf (order(1) = rstObj.fifty3) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty3)
                        '    row2(0) = rstObj.process_fifty3
                        'ElseIf (order(1) = rstObj.fifty4) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty4)
                        '    row2(0) = rstObj.process_fifty4
                        'ElseIf (order(1) = rstObj.sixty1) Then
                        '    cmbProcess.Items.Remove(rstObj.process_sixty1)
                        '    row2(0) = rstObj.process_sixty1
                        'ElseIf (order(1) = rstObj.sixty2) Then
                        '    cmbProcess.Items.Remove(rstObj.process_sixty2)
                        '    row2(0) = rstObj.process_sixty2
                        'ElseIf (order(1) = rstObj.seventy) Then
                        '    cmbProcess.Items.Remove(rstObj.process_seventy)
                        '    row2(0) = rstObj.process_seventy
                        'ElseIf (order(1) = rstObj.eighty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_eighty)
                        '    row2(0) = rstObj.process_eighty
                        'End If
#End Region

                        row2(1) = order(1)
                        row2(2) = ProdTime(1)
                        row2(3) = MoveTime(1)
                    End If


                    Dim row3 As DataRow = rstObj.dt2.NewRow()
                    rstObj.dt2.Rows.Add(row3)
                    'row3(0) = "30"
                    If count >= 3 Then
                        For i = 0 To rstObj.dtProcess.Rows.Count - 1
                            If order(2) = rstObj.dtProcess(i)(1) Then
                                cmbProcess.Items.Remove(rstObj.dtProcess(i)(0))
                                'cmbProcess.Items.Remove(rstObj.dtProcess(i)(1))
                                row3(0) = rstObj.dtProcess(i)(0)
                                ' MsgBox(rstObj.dtProcess(i)(0))
                            End If
                        Next
#Region "Comment"
                        'If (order(2) = rstObj.ten) Then
                        '    cmbProcess.Items.Remove(rstObj.process_ten)
                        '    row3(0) = rstObj.process_ten
                        'ElseIf (order(2) = rstObj.twenty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_twenty)
                        '    row3(0) = rstObj.process_twenty
                        'ElseIf (order(2) = rstObj.thirty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_thirty)
                        '    row3(0) = rstObj.process_thirty
                        'ElseIf (order(2) = rstObj.fourty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fourty)
                        '    row3(0) = rstObj.process_fourty
                        'ElseIf (order(2) = rstObj.fifty1) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty1)
                        '    row3(0) = rstObj.process_fifty1
                        'ElseIf (order(2) = rstObj.fifty2) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty2)
                        '    row3(0) = rstObj.process_fifty2
                        'ElseIf (order(2) = rstObj.fifty3) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty3)
                        '    row3(0) = rstObj.process_fifty3
                        'ElseIf (order(2) = rstObj.fifty4) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty4)
                        '    row3(0) = rstObj.process_fifty4
                        'ElseIf (order(2) = rstObj.sixty1) Then
                        '    cmbProcess.Items.Remove(rstObj.process_sixty1)
                        '    row3(0) = rstObj.process_sixty1
                        'ElseIf (order(2) = rstObj.sixty2) Then
                        '    cmbProcess.Items.Remove(rstObj.process_sixty2)
                        '    row3(0) = rstObj.process_sixty2
                        'ElseIf (order(2) = rstObj.seventy) Then
                        '    cmbProcess.Items.Remove(rstObj.process_seventy)
                        '    row3(0) = rstObj.process_seventy
                        'ElseIf (order(2) = rstObj.eighty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_eighty)
                        '    row3(0) = rstObj.process_eighty
                        'End If
#End Region
                        row3(1) = order(2)
                        row3(2) = ProdTime(2)
                        row3(3) = MoveTime(2)
                    End If



                    Dim row4 As DataRow = rstObj.dt2.NewRow()
                    rstObj.dt2.Rows.Add(row4)
                    'row4(0) = "40"
                    If count >= 4 Then
                        For i = 0 To rstObj.dtProcess.Rows.Count - 1
                            If order(3) = rstObj.dtProcess(i)(1) Then
                                cmbProcess.Items.Remove(rstObj.dtProcess(i)(0))
                                'cmbProcess.Items.Remove(rstObj.dtProcess(i)(1))
                                row4(0) = rstObj.dtProcess(i)(0)
                                ' MsgBox(rstObj.dtProcess(i)(0))
                            End If
                        Next
#Region "Comment"
                        'If (order(3) = rstObj.ten) Then
                        '    cmbProcess.Items.Remove(rstObj.process_ten)
                        '    row4(0) = rstObj.process_ten
                        'ElseIf (order(3) = rstObj.twenty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_twenty)
                        '    row4(0) = rstObj.process_twenty
                        'ElseIf (order(3) = rstObj.thirty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_thirty)
                        '    row4(0) = rstObj.process_thirty
                        'ElseIf (order(3) = rstObj.fourty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fourty)
                        '    row4(0) = rstObj.process_fourty
                        'ElseIf (order(3) = rstObj.fifty1) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty1)
                        '    row4(0) = rstObj.process_fifty1
                        'ElseIf (order(3) = rstObj.fifty2) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty2)
                        '    row4(0) = rstObj.process_fifty2
                        'ElseIf (order(3) = rstObj.fifty3) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty3)
                        '    row4(0) = rstObj.process_fifty3
                        'ElseIf (order(3) = rstObj.fifty4) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty4)
                        '    row4(0) = rstObj.process_fifty4
                        'ElseIf (order(3) = rstObj.sixty1) Then
                        '    cmbProcess.Items.Remove(rstObj.process_sixty1)
                        '    row4(0) = rstObj.process_sixty1
                        'ElseIf (order(3) = rstObj.sixty2) Then
                        '    cmbProcess.Items.Remove(rstObj.process_sixty2)
                        '    row4(0) = rstObj.process_sixty2
                        'ElseIf (order(3) = rstObj.seventy) Then
                        '    cmbProcess.Items.Remove(rstObj.process_seventy)
                        '    row4(0) = rstObj.process_seventy
                        'ElseIf (order(3) = rstObj.eighty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_eighty)
                        '    row4(0) = rstObj.process_eighty
                        'End If
#End Region
                        row4(1) = order(3)
                        row4(2) = ProdTime(3)
                        row4(3) = MoveTime(3)
                    End If


                    Dim row5 As DataRow = rstObj.dt2.NewRow()
                    rstObj.dt2.Rows.Add(row5)
                    'row5(0) = "50"
                    If count >= 5 Then
                        For i = 0 To rstObj.dtProcess.Rows.Count - 1
                            If order(4) = rstObj.dtProcess(i)(1) Then
                                cmbProcess.Items.Remove(rstObj.dtProcess(i)(0))
                                'cmbProcess.Items.Remove(rstObj.dtProcess(i)(1))
                                row5(0) = rstObj.dtProcess(i)(0)
                                ' MsgBox(rstObj.dtProcess(i)(0))
                            End If
                        Next
#Region "Comment"
                        'If (order(4) = rstObj.ten) Then
                        '    cmbProcess.Items.Remove(rstObj.process_ten)
                        '    row5(0) = rstObj.process_ten
                        'ElseIf (order(4) = rstObj.twenty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_twenty)
                        '    row5(0) = rstObj.process_twenty
                        'ElseIf (order(4) = rstObj.thirty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_thirty)
                        '    row5(0) = rstObj.process_thirty
                        'ElseIf (order(4) = rstObj.fourty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fourty)
                        '    row5(0) = rstObj.process_fourty
                        'ElseIf (order(4) = rstObj.fifty1) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty1)
                        '    row5(0) = rstObj.process_fifty1
                        'ElseIf (order(4) = rstObj.fifty2) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty2)
                        '    row5(0) = rstObj.process_fifty2
                        'ElseIf (order(4) = rstObj.fifty3) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty3)
                        '    row5(0) = rstObj.process_fifty3
                        'ElseIf (order(4) = rstObj.fifty4) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty4)
                        '    row5(0) = rstObj.process_fifty4
                        'ElseIf (order(4) = rstObj.sixty1) Then
                        '    cmbProcess.Items.Remove(rstObj.process_sixty1)
                        '    row5(0) = rstObj.process_sixty1
                        'ElseIf (order(4) = rstObj.sixty2) Then
                        '    cmbProcess.Items.Remove(rstObj.process_sixty2)
                        '    row5(0) = rstObj.process_sixty2
                        'ElseIf (order(4) = rstObj.seventy) Then
                        '    cmbProcess.Items.Remove(rstObj.process_seventy)
                        '    row5(0) = rstObj.process_seventy
                        'ElseIf (order(4) = rstObj.eighty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_eighty)
                        '    row5(0) = rstObj.process_eighty
                        'End If
#End Region
                        row5(1) = order(4)
                        row5(2) = ProdTime(4)
                        row5(3) = MoveTime(4)
                    End If


                    Dim row6 As DataRow = rstObj.dt2.NewRow()
                    rstObj.dt2.Rows.Add(row6)
                    'row6(0) = "60"
                    If count >= 6 Then
                        For i = 0 To rstObj.dtProcess.Rows.Count - 1
                            If order(5) = rstObj.dtProcess(i)(1) Then
                                cmbProcess.Items.Remove(rstObj.dtProcess(i)(0))
                                'cmbProcess.Items.Remove(rstObj.dtProcess(i)(1))
                                row6(0) = rstObj.dtProcess(i)(0)
                                ' MsgBox(rstObj.dtProcess(i)(0))
                            End If
                        Next
#Region "Comment"
                        'If (order(5) = rstObj.ten) Then
                        '    cmbProcess.Items.Remove(rstObj.process_ten)
                        '    row6(0) = rstObj.process_ten
                        'ElseIf (order(5) = rstObj.twenty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_twenty)
                        '    row6(0) = rstObj.process_twenty
                        'ElseIf (order(5) = rstObj.thirty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_thirty)
                        '    row6(0) = rstObj.process_thirty
                        'ElseIf (order(5) = rstObj.fourty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fourty)
                        '    row6(0) = rstObj.process_fourty
                        'ElseIf (order(5) = rstObj.fifty1) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty1)
                        '    row6(0) = rstObj.process_fifty1
                        'ElseIf (order(5) = rstObj.fifty2) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty2)
                        '    row6(0) = rstObj.process_fifty2
                        'ElseIf (order(5) = rstObj.fifty3) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty3)
                        '    row6(0) = rstObj.process_fifty3
                        'ElseIf (order(5) = rstObj.fifty4) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty4)
                        '    row6(0) = rstObj.process_fifty4
                        'ElseIf (order(5) = rstObj.sixty1) Then
                        '    cmbProcess.Items.Remove(rstObj.process_sixty1)
                        '    row6(0) = rstObj.process_sixty1
                        'ElseIf (order(5) = rstObj.sixty2) Then
                        '    cmbProcess.Items.Remove(rstObj.process_sixty2)
                        '    row6(0) = rstObj.process_sixty2
                        'ElseIf (order(5) = rstObj.seventy) Then
                        '    cmbProcess.Items.Remove(rstObj.process_seventy)
                        '    row6(0) = rstObj.process_seventy
                        'ElseIf (order(5) = rstObj.eighty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_eighty)
                        '    row6(0) = rstObj.process_eighty
                        'End If
#End Region
                        row6(1) = order(5)
                        row6(2) = ProdTime(5)
                        row6(3) = MoveTime(5)
                    End If


                    Dim row7 As DataRow = rstObj.dt2.NewRow()
                    rstObj.dt2.Rows.Add(row7)
                    'row7(0) = "70"
                    If count >= 7 Then
                        For i = 0 To rstObj.dtProcess.Rows.Count - 1
                            If order(6) = rstObj.dtProcess(i)(1) Then
                                cmbProcess.Items.Remove(rstObj.dtProcess(i)(0))
                                'cmbProcess.Items.Remove(rstObj.dtProcess(i)(1))
                                row7(0) = rstObj.dtProcess(i)(0)
                                ' MsgBox(rstObj.dtProcess(i)(0))
                            End If
                        Next
#Region "Comment"
                        'If (order(6) = rstObj.ten) Then
                        '    cmbProcess.Items.Remove(rstObj.process_ten)
                        '    row7(0) = rstObj.process_ten
                        'ElseIf (order(6) = rstObj.twenty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_twenty)
                        '    row7(0) = rstObj.process_twenty
                        'ElseIf (order(6) = rstObj.thirty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_thirty)
                        '    row7(0) = rstObj.process_thirty
                        'ElseIf (order(6) = rstObj.fourty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fourty)
                        '    row7(0) = rstObj.process_fourty
                        'ElseIf (order(6) = rstObj.fifty1) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty1)
                        '    row7(0) = rstObj.process_fifty1
                        'ElseIf (order(6) = rstObj.fifty2) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty2)
                        '    row7(0) = rstObj.process_fifty2
                        'ElseIf (order(6) = rstObj.fifty3) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty3)
                        '    row7(0) = rstObj.process_fifty3
                        'ElseIf (order(6) = rstObj.fifty4) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty4)
                        '    row7(0) = rstObj.process_fifty4
                        'ElseIf (order(6) = rstObj.sixty1) Then
                        '    cmbProcess.Items.Remove(rstObj.process_sixty1)
                        '    row7(0) = rstObj.process_sixty1
                        'ElseIf (order(6) = rstObj.sixty2) Then
                        '    cmbProcess.Items.Remove(rstObj.process_sixty2)
                        '    row7(0) = rstObj.process_sixty2
                        'ElseIf (order(6) = rstObj.seventy) Then
                        '    cmbProcess.Items.Remove(rstObj.seventy)
                        '    row7(0) = rstObj.process_seventy
                        'ElseIf (order(6) = rstObj.eighty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_eighty)
                        '    row7(0) = rstObj.process_eighty
                        'End If
#End Region

                        row7(1) = order(6)
                        row7(2) = ProdTime(6)
                        row7(3) = MoveTime(6)
                    End If


                    Dim row8 As DataRow = rstObj.dt2.NewRow()
                    rstObj.dt2.Rows.Add(row8)
                    'row8(0) = "80"
                    If count >= 8 Then
                        For i = 0 To rstObj.dtProcess.Rows.Count - 1
                            If order(7) = rstObj.dtProcess(i)(1) Then
                                cmbProcess.Items.Remove(rstObj.dtProcess(i)(0))
                                'cmbProcess.Items.Remove(rstObj.dtProcess(i)(1))
                                row8(0) = rstObj.dtProcess(i)(0)
                                ' MsgBox(rstObj.dtProcess(i)(0))
                            End If
                        Next
#Region "Comment"
                        'Next
                        'If (order(7) = rstObj.ten) Then
                        '    cmbProcess.Items.Remove(rstObj.process_ten)
                        '    row8(0) = rstObj.process_ten
                        'ElseIf (order(7) = rstObj.twenty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_twenty)
                        '    row8(0) = rstObj.process_twenty
                        'ElseIf (order(7) = rstObj.thirty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_thirty)
                        '    row8(0) = rstObj.process_thirty
                        'ElseIf (order(7) = rstObj.fourty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fourty)
                        '    row8(0) = rstObj.process_fourty
                        'ElseIf (order(7) = rstObj.fifty1) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty1)
                        '    row8(0) = rstObj.process_fifty1
                        'ElseIf (order(7) = rstObj.fifty2) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty2)
                        '    row8(0) = rstObj.process_fifty2
                        'ElseIf (order(7) = rstObj.fifty3) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty3)
                        '    row8(0) = rstObj.process_fifty3
                        'ElseIf (order(7) = rstObj.fifty4) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty4)
                        '    row8(0) = rstObj.process_fifty4
                        'ElseIf (order(7) = rstObj.sixty1) Then
                        '    cmbProcess.Items.Remove(rstObj.process_sixty1)
                        '    row8(0) = rstObj.process_sixty1
                        'ElseIf (order(7) = rstObj.sixty2) Then
                        '    cmbProcess.Items.Remove(rstObj.process_sixty2)
                        '    row8(0) = rstObj.process_sixty2
                        'ElseIf (order(7) = rstObj.seventy) Then
                        '    cmbProcess.Items.Remove(rstObj.process_seventy)
                        '    row8(0) = rstObj.process_seventy
                        'ElseIf (order(7) = rstObj.eighty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_eighty)
                        '    row8(0) = rstObj.process_eighty
                        'End If
#End Region
                        row8(1) = order(7)
                        row8(2) = ProdTime(7)
                        row8(3) = MoveTime(7)
                    End If


                    Dim row9 As DataRow = rstObj.dt2.NewRow()
                    rstObj.dt2.Rows.Add(row9)
                    'row9(0) = "90"
                    If count >= 9 Then
                        For i = 0 To rstObj.dtProcess.Rows.Count - 1
                            If order(8) = rstObj.dtProcess(i)(1) Then
                                cmbProcess.Items.Remove(rstObj.dtProcess(i)(0))
                                'cmbProcess.Items.Remove(rstObj.dtProcess(i)(1))
                                row9(0) = rstObj.dtProcess(i)(0)
                                ' MsgBox(rstObj.dtProcess(i)(0))
                            End If
                        Next
#Region "Comment"
                        'If (order(8) = rstObj.ten) Then
                        '    cmbProcess.Items.Remove(rstObj.process_ten)
                        '    row9(0) = rstObj.process_ten
                        'ElseIf (order(8) = rstObj.twenty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_twenty)
                        '    row9(0) = rstObj.process_twenty
                        'ElseIf (order(8) = rstObj.thirty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_thirty)
                        '    row9(0) = rstObj.process_thirty
                        'ElseIf (order(8) = rstObj.fourty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fourty)
                        '    row9(0) = rstObj.process_fourty
                        'ElseIf (order(8) = rstObj.fifty1) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty1)
                        '    row9(0) = rstObj.process_fifty1
                        'ElseIf (order(8) = rstObj.fifty2) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty2)
                        '    row9(0) = rstObj.process_fifty2
                        'ElseIf (order(8) = rstObj.fifty3) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty3)
                        '    row9(0) = rstObj.process_fifty3
                        'ElseIf (order(8) = rstObj.fifty4) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty4)
                        '    row9(0) = rstObj.process_fifty4
                        'ElseIf (order(8) = rstObj.sixty1) Then
                        '    cmbProcess.Items.Remove(rstObj.process_sixty1)
                        '    row9(0) = rstObj.process_sixty1
                        'ElseIf (order(8) = rstObj.sixty2) Then
                        '    cmbProcess.Items.Remove(rstObj.process_sixty2)
                        '    row9(0) = rstObj.process_sixty2
                        'ElseIf (order(8) = rstObj.seventy) Then
                        '    cmbProcess.Items.Remove(rstObj.process_seventy)
                        '    row9(0) = rstObj.process_seventy
                        'ElseIf (order(8) = rstObj.eighty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_eighty)
                        '    row9(0) = rstObj.process_eighty
                        'End If
#End Region
                        row9(1) = order(8)
                        row9(2) = ProdTime(8)
                        row9(3) = MoveTime(8)
                    Else
                        row9(0) = ""
                        row9(1) = ""
                        row9(2) = ""
                        row9(3) = ""
                    End If




                    Dim row10 As DataRow = rstObj.dt2.NewRow()
                    rstObj.dt2.Rows.Add(row10)
                    'row10(0) = "100"
                    If count = 10 Then
                        For i = 0 To rstObj.dtProcess.Rows.Count - 1
                            If order(9) = rstObj.dtProcess(i)(1) Then
                                cmbProcess.Items.Remove(rstObj.dtProcess(i)(0))
                                'cmbProcess.Items.Remove(rstObj.dtProcess(i)(1))
                                row10(0) = rstObj.dtProcess(i)(0)
                                ' MsgBox(rstObj.dtProcess(i)(0))
                            End If
                        Next

#Region "Comment"
                        'If (order(9) = rstObj.ten) Then
                        '    cmbProcess.Items.Remove(rstObj.process_ten)
                        '    row10(0) = rstObj.process_ten
                        'ElseIf (order(9) = rstObj.twenty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_twenty)
                        '    row10(0) = rstObj.process_twenty
                        'ElseIf (order(9) = rstObj.thirty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_thirty)
                        '    row10(0) = rstObj.process_thirty
                        'ElseIf (order(9) = rstObj.fourty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fourty)
                        '    row10(0) = rstObj.process_fourty
                        'ElseIf (order(9) = rstObj.fifty1) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty1)
                        '    row10(0) = rstObj.process_fifty1
                        'ElseIf (order(9) = rstObj.fifty2) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty2)
                        '    row10(0) = rstObj.process_fifty2
                        'ElseIf (order(9) = rstObj.fifty3) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty3)
                        '    row10(0) = rstObj.process_fifty3
                        'ElseIf (order(9) = rstObj.fifty4) Then
                        '    cmbProcess.Items.Remove(rstObj.process_fifty4)
                        '    row10(0) = rstObj.process_fifty4
                        'ElseIf (order(9) = rstObj.sixty1) Then
                        '    cmbProcess.Items.Remove(rstObj.process_sixty1)
                        '    row10(0) = rstObj.process_sixty1
                        'ElseIf (order(9) = rstObj.sixty2) Then
                        '    cmbProcess.Items.Remove(rstObj.process_sixty2)
                        '    row10(0) = rstObj.process_sixty2
                        'ElseIf (order(9) = rstObj.seventy) Then
                        '    cmbProcess.Items.Remove(rstObj.process_seventy)
                        '    row10(0) = rstObj.process_seventy
                        'ElseIf (order(9) = rstObj.eighty) Then
                        '    cmbProcess.Items.Remove(rstObj.process_eighty)
                        '    row10(0) = rstObj.process_eighty
                        'End If
#End Region
                        row10(1) = order(9)
                        row10(2) = ProdTime(9)
                        row10(3) = MoveTime(9)
                    Else
                        row10(0) = ""
                        row10(1) = ""
                        row10(2) = ""
                        row10(3) = ""
                    End If

                Catch ex As Exception
                    CustomLogUtil.Log($"While clicking on MainGrid", ex.Message, ex.StackTrace)
                End Try



                Get_dt3data(rstObj)

                Lbldeclaration()
                ShowDgvSub2(rstObj)


                'Dim mypath As New System.IO.DirectoryInfo(System.IO.Path.Combine(Environment.CurrentDirectory, ".."))
                'Dim sJpgFile As String = Nothing
                'sJpgFile = mypath.FullName + "\Thumbnail\" + rstObj.PartName + ".jpg"
                'File.Delete(sJpgFile)
                'DeleteDir()
            Else
                MessageBox.Show("Please select valid Part Name", "Message")
            End If
        Catch ex As Exception

            MessageBox.Show($"Error While Clicking On Main Grid{vbNewLine}{ex.Message}{ex.StackTrace}", "Message")
            CustomLogUtil.Log("While Clicking On Main Grid", ex.Message, ex.StackTrace)
        End Try

    End Sub

    Public Sub CmbCategoryItems()
        cmbCategory.Items.Add("Assembly")
        cmbCategory.Items.Add("SheetMetal")
        cmbCategory.Items.Add("Structure")
        cmbCategory.Items.Add("Misc. Parts")
    End Sub
    Public Sub CmbProcessItems(rstObj As RountingSequenceClass)
        cmbProcess.Items.Clear()
        For i = 0 To rstObj.dtProcess.Rows.Count - 1
            'MsgBox(rstObj.dtProcess(i)(0))
            cmbProcess.Items.Add(rstObj.dtProcess(i)(0))
        Next

    End Sub
    Private Sub DgvSub_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvSub.CellClick
        Try
            Dim selectedGR As DataGridViewRow
            selectedGR = dgvSub.Rows(e.RowIndex)
            rstObj.dgvCalculatorRowIndex = selectedGR
            If cmbProcess.SelectedItem IsNot Nothing Then
                If Not selectedGR.Cells(0).Value.ToString = "" Then
                    cmbProcess.Items.Add(selectedGR.Cells(0).Value)
                End If
                selectedGR.Cells(0).Value = cmbProcess.SelectedItem
            End If
            cmbProcess.Enabled = True


            'MsgBox(dgvSub.Rows(index).Cells(0).Value.ToString())
            For i = 0 To rstObj.dtProcess.Rows.Count - 1

                If cmbProcess.SelectedItem = rstObj.dtProcess(i)(0) Then
                    cmbProcess.Items.Remove(rstObj.dtProcess(i)(0))
                    selectedGR.Cells(1).Value = rstObj.dtProcess(i)(1)
                    selectedGR.Cells(2).Value = rstObj.dtProcess(i)(3)
                    selectedGR.Cells(3).Value = rstObj.dtProcess(i)(2)
                End If
            Next
        Catch ex As Exception
            CustomLogUtil.Log($"While clicking on the SubGrid", ex.Message, ex.StackTrace)
            MessageBox.Show($"Error While Clicking On The SubGrid {vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}", "Message")
        End Try
        DGV_calculatorData()

        'cmbProcess.Enabled = False
        BtnCalculateAndSave.Enabled = True

    End Sub
    Public Sub DGV_calculatorData()
        Try
            If DGV_Calculator.RowCount > 0 Then
                DGV_Calculator.Columns.Clear()
                DGV_Calculator.Rows.Clear()
            End If

            Dim checkCol As New DataGridViewCheckBoxColumn()
            Dim ProdTime As String = String.Empty
            checkCol.HeaderText = " "
            checkCol.Width = "30"
            DGV_Calculator.Columns.Add(checkCol)

            Dim processname As String = "P# " + rstObj.dgvCalculatorRowIndex.Cells(0).Value
            DGV_Calculator.Columns.Add("ProcessName", processname)

            Dim wc As String = ("WC# " + rstObj.dgvCalculatorRowIndex.Cells(1).Value)
            DGV_Calculator.Columns.Add("WC", wc)
            DGV_Calculator.Columns.Add("ProdTime", "PT#")
            If Not rstObj.dgvCalculatorRowIndex.Cells(0).Value.ToString = "" Then


                For i = 0 To rstObj.dtProcess.Rows.Count - 1
                    If rstObj.dtProcess(i)(0) = rstObj.dgvCalculatorRowIndex.Cells(0).Value Then
                        ProdTime = rstObj.dtProcess(i)(3)
                        Exit For
                    End If
                Next


            End If
            Dim count As Integer = 0
#Region "Comment"
            'If rstObj.dt3(i)(0) = "Material Thickness" Then
            '    If Not rstObj.dt3(i)(1) = Nothing Or rstObj.dt3(i)(1) = "" Then
            '        With DGV_Calculator
            '            .Rows.Add()
            '            .Rows(count).Cells(1).Value = rstObj.dt3(i)(0)
            '            .Rows(count).Cells(2).Value = rstObj.dt3(i)(1)
            '            .Rows(count).Cells(3).Value = ProdTime
            '        End With
            '        count += 1
            '    End If
            'End If
            'If rstObj.dt3(i)(0) = "Mass" Then
            '    If Not rstObj.dt3(i)(1) = Nothing Or rstObj.dt3(i)(1) = "" Then

            '        With DGV_Calculator
            '            .Rows.Add()
            '            .Rows(count).Cells(1).Value = rstObj.dt3(i)(0)
            '            .Rows(count).Cells(2).Value = rstObj.dt3(i)(1)
            '            .Rows(count).Cells(3).Value = ProdTime
            '        End With
            '        count += 1
            '    End If
            'End If
            'If rstObj.dt3(i)(0) = "Bend Radius" Then
            '    If Not rstObj.dt3(i)(1) = Nothing Or rstObj.dt3(i)(1) = "" Then

            '        With DGV_Calculator
            '            .Rows.Add()
            '            .Rows(count).Cells(1).Value = rstObj.dt3(i)(0)
            '            .Rows(count).Cells(2).Value = rstObj.dt3(i)(1)
            '            .Rows(count).Cells(3).Value = ProdTime
            '        End With
            '        count += 1
            '    End If
            'End If

            'If rstObj.dt3(i)(0) = "Flat_Pattern_Model_CutSizeX" Then
            '    If Not rstObj.dt3(i)(1) = Nothing Or rstObj.dt3(i)(1) = "" Then

            '        With DGV_Calculator
            '            .Rows.Add()
            '            .Rows(count).Cells(1).Value = rstObj.dt3(i)(0)
            '            .Rows(count).Cells(2).Value = rstObj.dt3(i)(1)
            '            .Rows(count).Cells(3).Value = ProdTime
            '        End With
            '        count += 1
            '    End If
            'End If
            'If rstObj.dt3(i)(0) = "Flat_Pattern_Model_CutSizeY" Then
            '    If Not rstObj.dt3(i)(1) = Nothing Or rstObj.dt3(i)(1) = "" Then

            '        With DGV_Calculator
            '            .Rows.Add()
            '            .Rows(count).Cells(1).Value = rstObj.dt3(i)(0)
            '            .Rows(count).Cells(2).Value = rstObj.dt3(i)(1)
            '            .Rows(count).Cells(3).Value = ProdTime
            '        End With
            '        count += 1
            '    End If
            'End If

#End Region
            'assembly data for calculator
#Region "Assembly"

            If rstObj.CategoryName = "Assembly".ToUpper Then
                For i = 0 To rstObj.dt3.Rows.Count - 1
                    If rstObj.dt3(i)(0) = "Part Qty" Then
                        If Not rstObj.dt3(i)(1) = Nothing Or rstObj.dt3(i)(1) = "" Then

                            With DGV_Calculator
                                .Rows.Add()
                                .Rows(count).Cells(1).Value = rstObj.dt3(i)(0)
                                .Rows(count).Cells(2).Value = rstObj.dt3(i)(1)
                                .Rows(count).Cells(3).Value = ProdTime
                            End With
                            count = 0
                        End If
                    End If
                Next
            End If

#End Region
            'sheetmetal data for calculator
#Region "SheetMetal"

            If rstObj.CategoryName = "SheetMetal".ToUpper Then
                For i = 0 To rstObj.dt3.Rows.Count - 1
                    If rstObj.dt3(i)(0) = "Bend Qty" Then
                        If Not rstObj.dt3(i)(1) = Nothing Or rstObj.dt3(i)(1) = "" Then

                            With DGV_Calculator
                                .Rows.Add()
                                .Rows(count).Cells(1).Value = rstObj.dt3(i)(0)
                                .Rows(count).Cells(2).Value = rstObj.dt3(i)(1)
                                .Rows(count).Cells(3).Value = ProdTime
                            End With
                            count += 1
                        End If
                    End If
                    If rstObj.dt3(i)(0) = "Hole Qty" Then
                        If Not rstObj.dt3(i)(1) = Nothing Or rstObj.dt3(i)(1) = "" Then

                            With DGV_Calculator
                                .Rows.Add()
                                .Rows(count).Cells(1).Value = rstObj.dt3(i)(0)
                                .Rows(count).Cells(2).Value = rstObj.dt3(i)(1)
                                .Rows(count).Cells(3).Value = ProdTime
                            End With
                            count += 1
                        End If
                    End If



                    If rstObj.dt3(i)(0) = "Perimeter" Then
                        If Not rstObj.dt3(i)(1) = Nothing Or rstObj.dt3(i)(1) = "" Then

                            With DGV_Calculator
                                .Rows.Add()
                                .Rows(count).Cells(1).Value = rstObj.dt3(i)(0)
                                .Rows(count).Cells(2).Value = rstObj.dt3(i)(1)
                                .Rows(count).Cells(3).Value = ProdTime
                            End With
                            count = 0
                        End If
                    End If
                Next
            End If

#End Region
            'structure data for calculator
#Region "Structure"
            If rstObj.CategoryName = "Structure".ToUpper Then
                For i = 0 To rstObj.dt3.Rows.Count - 1
                    If rstObj.dt3(i)(0) = "Hole Qty" Then
                        If Not rstObj.dt3(i)(1) = Nothing Or rstObj.dt3(i)(1) = "" Then

                            With DGV_Calculator
                                .Rows.Add()
                                .Rows(count).Cells(1).Value = rstObj.dt3(i)(0)
                                .Rows(count).Cells(2).Value = rstObj.dt3(i)(1)
                                .Rows(count).Cells(3).Value = ProdTime
                            End With
                            count = 0
                        End If
                    End If
                Next
            End If

#End Region
            'misc. parts data for calculator
#Region "Misc. Parts"
            If rstObj.CategoryName = "Misc. Parts".ToUpper Then
                For i = 0 To rstObj.dt3.Rows.Count - 1
                    If rstObj.dt3(i)(0) = "Hole Qty" Then
                        If Not rstObj.dt3(i)(1) = Nothing Or rstObj.dt3(i)(1) = "" Then

                            With DGV_Calculator
                                .Rows.Add()
                                .Rows(count).Cells(1).Value = rstObj.dt3(i)(0)
                                .Rows(count).Cells(2).Value = rstObj.dt3(i)(1)
                                .Rows(count).Cells(3).Value = ProdTime
                            End With
                            count = 0
                        End If
                    End If
                Next
            End If
#End Region
            DGV_Calculator.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
            DGV_Calculator.Columns(0).Width = 25
            DGV_Calculator.Columns(1).Width = 100
        Catch ex As Exception
            CustomLogUtil.Log("While Calculating the prodTime ", ex.Message, ex.StackTrace)
            MessageBox.Show($"Error While Calculating the prodTime {vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}", "Message")

        End Try

    End Sub
    Public Sub BtnUpData()
        Try
            Dim rowIndex = dgvSub.SelectedCells(0).OwningRow.Index
            If rowIndex > 0 Then
                Dim row As DataRow = rstObj.dt2.NewRow
                ' row(0) = Convert.ToInt32(dgvSub.Rows(rowIndex).Cells(0).Value.ToString)
                row(0) = (dgvSub.Rows(rowIndex).Cells(0).Value.ToString)
                row(1) = Convert.ToInt32(dgvSub.Rows(rowIndex).Cells(1).Value.ToString)
                row(2) = Convert.ToDouble(dgvSub.Rows(rowIndex).Cells(2).Value.ToString)
                row(3) = Convert.ToInt32(dgvSub.Rows(rowIndex).Cells(3).Value.ToString)
                rstObj.dt2.Rows.RemoveAt(rowIndex)
                rstObj.dt2.Rows.InsertAt(row, rowIndex - 1)
                dgvSub.Rows(rowIndex - 1).Selected = True
                dgvSub.DataSource = rstObj.dt2
            End If
        Catch ex As Exception
            CustomLogUtil.Log($"While clicking on UP BUTTON ..", ex.Message, ex.StackTrace)
            MessageBox.Show($"Error While clicking on UP BUTTON ..{vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}", "Message")
        End Try

    End Sub
    Private Sub BtnUp_Click(sender As Object, e As EventArgs) Handles BtnUp.Click
        BtnUpData()
    End Sub

    Private Sub BtnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
        BtnEnable()
        ApplyData()
        rstObj.dt2.Clear()
        rstObj.dt3.Clear()
        DGV_Calculator.Rows.Clear()
        DGV_Calculator.Columns.Clear()
        'rstObj.dt2.Columns.Clear()
    End Sub
    Public Sub BtnEnable()
        If BtnUp.Enabled = True And BtnDown.Enabled = True And BtnDelete.Enabled = True Then
            ' dgvSub.Enabled = False
            dgvSub.ReadOnly = True



            BtnUp.Enabled = False
            BtnDown.Enabled = False
            BtnDelete.Enabled = False
            btnApply.Enabled = False
            btnApply.Visible = False

            'BtnApplyValues.Enabled = True
            'BtnApplyValues.Visible = True


            dgvsub2.Enabled = True
            dgvsub2.ReadOnly = False


        Else

            dgvSub.Enabled = True
            dgvSub.ReadOnly = False

            BtnUp.Enabled = True
            BtnDown.Enabled = True
            BtnDelete.Enabled = True
            btnApply.Enabled = True
            btnApply.Visible = True

            BtnApplyValues.Enabled = False
            BtnApplyValues.Visible = False


            'dgvsub2.Enabled = False
            dgvsub2.ReadOnly = True


        End If

        If Not btnApply.Enabled = False And BtnApplyValues.Enabled = False Then
            dgvMain.Enabled = False
            dgvMain.ReadOnly = True
        Else
            dgvMain.Enabled = True
            dgvMain.ReadOnly = False
        End If

    End Sub
    Public Sub Set_dt3data(rstObj As RountingSequenceClass)
        Try
            resSave.Set_dt3data(rstObj)
        Catch ex As Exception
            CustomLogUtil.Log($"While Setting dt3Data ..", ex.Message, ex.StackTrace)
            MessageBox.Show($"Error While Setting  Databalet3Data ..{vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}", "Message")
        End Try

    End Sub

    Public Sub Get_Maindt_data(rstObj As RountingSequenceClass)
        Try
            resSave.Get_Maindt_data(rstObj)
            ShowDgvMain(rstObj)

            Set_dt2data()
            'showDgvSub(rstObj)
        Catch ex As Exception
            CustomLogUtil.Log($"While Getting Maindt data..", ex.Message, ex.StackTrace)
            MessageBox.Show($"Error While Getting Data From Datatable ..{vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}", "Message")
        End Try

    End Sub



    Public Sub ApplyData()
        Try
            Dim ProdTotal As Decimal = Nothing
            Dim MoveTotal As Integer = Nothing
            For i = 0 To 9
                'sMsgBox(rstObj.dgvmainRowIndex.Cells(i + 1).Value.ToString)

                If Not dgvSub.Rows(i).Cells(2).Value.ToString = "" And Not dgvSub.Rows(i).Cells(3).Value.ToString = "" Then

                    ProdTotal += rstObj.dt2.Rows(i)(2)
                    MoveTotal += rstObj.dt2.Rows(i)(3)
                End If
                rstObj.dgvmainRowIndex.Cells(i + 1).Value = rstObj.dt2.Rows(i)(1)
                rstObj.dgvmainRowIndexProdTIME.Cells(i + 1).Value = rstObj.dt2.Rows(i)(2)
                rstObj.dgvmainRowIndexMoveTIME.Cells(i + 1).Value = rstObj.dt2.Rows(i)(3)
            Next
            rstObj.dgvmainRowIndexProdTIME.Cells(11).Value = ProdTotal
            rstObj.dgvmainRowIndexMoveTIME.Cells(11).Value = MoveTotal

        Catch ex As Exception
            CustomLogUtil.Log($"While clicking on Apply Sequence BUTTON ..", ex.Message, ex.StackTrace)
            MessageBox.Show($"Error While clicking on Apply Sequence BUTTON ..{vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}", "Message")
        End Try

    End Sub
    Private Sub BtnAproveSequence_Click(sender As Object, e As EventArgs) Handles btnAproveSequence.Click


        If Not txtFilename.Text = "" And Not txtFoldername.Text = "" Then
            WaitStartSave()
            ApproveSequence(rstObj)
            ' reset()
            If BtnGetDataAndReset.Text = "Reset" Then
                cmbProcess.Enabled = False
                BtnPreview.Enabled = False
                'If cmbProcess.Enabled = True And BtnUp.Enabled = True And BtnDown.Enabled = True And BtnDelete.Enabled = True And BtnPreview.Enabled = True Then
                '    BtnApplyValues.Enabled = False
                '    cmbProcess.Enabled = False
                '    BtnUp.Enabled = False
                '    BtnDown.Enabled = False
                '    BtnDelete.Enabled = False
                '    BtnPreview.Enabled = False
                '    BtnApplyValues.Visible = True

                'End If

                BtnEnable()
                If BtnUp.Enabled = True And BtnDown.Enabled = True And BtnDelete.Enabled = True Then
                    BtnEnable()
                End If
                Reset()
                DGV_Calculator.Columns.Clear()
                rstObj.dtProcess.Clear()
                rstObj.dtProcess.Columns.Clear()
                ReleaseObject(rstObj)

                '  ResetAll()
                BtnGetDataAndReset.Text = "Get Data"

            End If
            WaitEndSave()

        Else
            MessageBox.Show("Please select File for approve Sequence", "Message")

        End If

    End Sub

    Public Sub ApproveSequence(rstObj As RountingSequenceClass)
        Try
            resSave.ApproveSequence(rstObj)
            MessageBox.Show("Routing Sequence Report successfully created", "Message")
            CustomLogUtil.Heading($"Routing Sequence Report successfully created")

            If IO.Directory.Exists(rstObj.dir) Then
                Process.Start(rstObj.dir)
            End If
        Catch ex As Exception
            MessageBox.Show("Error While Creating Routing Sequence Report", "Message")
            CustomLogUtil.Log("While Creating Routing Sequence Report", ex.Message, ex.StackTrace)
        End Try

    End Sub

    Public Sub Reset()

        lblMainGridTitle.Text = ""
        lblSubGridTitle.Text = ""
        lblUserName.Text = ""
        txtFilename.Clear()
        txtFoldername.Clear()
        rstObj.Maindt.Clear()
        rstObj.Maindt.Columns.Clear()
        rstObj.dt2.Clear()
        rstObj.dt2.Columns.Clear()
        rstObj.dt3.Clear()
        rstObj.dt3.Columns.Clear()

        dgvMain.DataSource = rstObj.Maindt
        dgvSub.DataSource = rstObj.dt2

        cmbCategory.Enabled = True
        cmbCategory.SelectedItem = ""
        cmbCategory.Text = "CATEGORY"
        rstObj.CategoryName = Nothing
    End Sub

    Private Sub BtnDown_Click(sender As Object, e As EventArgs) Handles BtnDown.Click
        BtnDownData()
    End Sub
    Public Sub BtnDownData()
        Try
            Dim rowIndex = dgvSub.SelectedCells(0).OwningRow.Index
            If rowIndex < dgvSub.Rows.Count - 2 Then
                Dim row As DataRow = rstObj.dt2.NewRow
                'row(0) = Convert.ToInt32(dgvSub.Rows(rowIndex).Cells(0).Value.ToString)
                row(0) = (dgvSub.Rows(rowIndex).Cells(0).Value.ToString)
                row(1) = Convert.ToInt32(dgvSub.Rows(rowIndex).Cells(1).Value.ToString)
                row(2) = Convert.ToDouble(dgvSub.Rows(rowIndex).Cells(2).Value.ToString)
                row(3) = Convert.ToInt32(dgvSub.Rows(rowIndex).Cells(3).Value.ToString)
                rstObj.dt2.Rows.RemoveAt(rowIndex)
                rstObj.dt2.Rows.InsertAt(row, rowIndex + 1)
                dgvSub.Rows(rowIndex + 1).Selected = True
                dgvSub.DataSource = rstObj.dt2
            End If
        Catch ex As Exception
            CustomLogUtil.Log("While clicking on DOWN BUTTON ..", ex.Message, ex.StackTrace)
            MessageBox.Show($"Error While clicking on DOWN BUTTON ..{vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}", "Message")
        End Try
    End Sub
    Public Sub BtnDeleteData()
        Try
            Dim rowIndex = dgvSub.SelectedCells(0).OwningRow.Index

            Dim row As DataRow = rstObj.dt2.NewRow
            'MsgBox(dgvSub.SelectedCells(0).Value)
            cmbProcess.Items.Add(dgvSub.SelectedCells(0).Value)

            ' row(0) = Convert.ToInt32(dgvSub.Rows(rowIndex).Cells(0).Value.ToString)
            row(0) = ""
            row(1) = ""
            row(2) = ""
            row(3) = ""
            rstObj.dt2.Rows.RemoveAt(rowIndex)
            rstObj.dt2.Rows.InsertAt(row, rowIndex + 9)

            dgvSub.Rows(rowIndex + 1).Selected = True

            dgvSub.DataSource = rstObj.dt2

        Catch ex As Exception
            CustomLogUtil.Log($"While clicking on DELETE BUTTON ..", ex.Message, ex.StackTrace)
            MessageBox.Show($"Error While clicking on DELETE BUTTON ..{vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}", "Message")
        End Try
    End Sub

    Private Sub BtnDelete_Click(sender As Object, e As EventArgs) Handles BtnDelete.Click
        BtnDeleteData()
    End Sub

    Private Sub BtnApplyValues_Click(sender As Object, e As EventArgs) Handles BtnApplyValues.Click
        Try
            Set_dt3data(rstObj)

            Get_Maindt_data(rstObj)
            BtnEnable()
        Catch ex As Exception
            CustomLogUtil.Log($"While clicking on ApplyValues BUTTON ..", ex.Message, ex.StackTrace)
            MessageBox.Show($"Error While clicking on ApplyValues BUTTON ..{vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}", "Message")
        End Try


        'rstObj.dt3.Clear()
        ' rstObj.dt3.Columns.Clear()

    End Sub

    Public Sub OpenPreview(rstObj As RountingSequenceClass)
        Dim a As New RST_Preview
        a.PictureBox1.Image = rstObj.image
        a.Label1.Text = rstObj.FilePath
        a.Show()
    End Sub
    Private Sub BtnPreview_Click(sender As Object, e As EventArgs) Handles BtnPreview.Click
        Preview()
    End Sub
    Public Sub Preview()
        Try
            BtnPreview.Enabled = False
            WaitStartSave()
            resSave.OpenSEDocument(rstObj)
            WaitEndSave()
            OpenPreview(rstObj)
        Catch ex As Exception
            CustomLogUtil.Log($"While clicking on Preview BUTTON ..", ex.Message, ex.StackTrace)
            MessageBox.Show($"Error While clicking on Preview BUTTON ..{vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}", "Message")
        End Try
        'PictureBox1.Image.Dispose()
    End Sub

    Private Sub BtnCalculateAndSave_Click(sender As Object, e As EventArgs) Handles BtnCalculateAndSave.Click
        CalculateAndSaveData()
    End Sub
    Public Sub CalculateAndSaveData()
        Try
            Dim mul As Decimal
            Dim add As Decimal = 0
            If BtnCalculateAndSave.Text = "Calculate" Then

                If rstObj.CategoryName = "Assembly".ToUpper Then
                    For i = 0 To DGV_Calculator.Rows.Count - 1
                        If DGV_Calculator.Rows(i).Cells(0).Value = "True" And Not DGV_Calculator.Rows(i).Cells(1).Value = Nothing Then
                            If Not DGV_Calculator.Rows(i).Cells(2).Value = "" And Not DGV_Calculator.Rows(i).Cells(3).Value = "" Then
                                Dim s As String = DGV_Calculator.Rows(i).Cells(3).Value
                                If s.Contains("/") Then
                                    Dim arrStr() As String = s.Split("/")
                                    Dim a As Decimal = arrStr(0)
                                    Dim b As Decimal = arrStr(1)

                                    mul = a + ((DGV_Calculator.Rows(i).Cells(2).Value - 1) * b)
                                    If add = 0 Then
                                        add = mul
                                    Else
                                        add += mul
                                    End If

                                Else
                                    mul = DGV_Calculator.Rows(i).Cells(2).Value * DGV_Calculator.Rows(i).Cells(3).Value
                                    If add = 0 Then
                                        add = mul
                                    Else
                                        add += mul
                                    End If
                                End If
                            End If
                        End If
                        mul = Nothing
                    Next
                Else

                    For i = 0 To DGV_Calculator.Rows.Count - 1
                        If DGV_Calculator.Rows(i).Cells(0).Value = "True" And Not DGV_Calculator.Rows(i).Cells(1).Value = Nothing Then
                            If Not DGV_Calculator.Rows(i).Cells(2).Value = "" And Not DGV_Calculator.Rows(i).Cells(3).Value Then


                                mul = DGV_Calculator.Rows(i).Cells(2).Value * DGV_Calculator.Rows(i).Cells(3).Value
                                If add = 0 Then
                                    add = mul
                                Else
                                    add += mul
                                End If
                            End If
                        End If
                        mul = Nothing
                    Next
                End If
                If rstObj.CategoryName = "Assembly".ToUpper Then
                    DGV_Calculator.Rows(1).Cells(1).Value = "Final PTime"
                    DGV_Calculator.Rows(1).Cells(3).Value = Math.Round(add, 2)
                ElseIf rstObj.CategoryName = "SheetMetal".ToUpper Then
                    DGV_Calculator.Rows(3).Cells(1).Value = "Final PTime"
                    DGV_Calculator.Rows(3).Cells(3).Value = Math.Round(add, 2)
                ElseIf rstObj.CategoryName = "Structure".ToUpper Then
                    DGV_Calculator.Rows(1).Cells(1).Value = "Final PTime"
                    DGV_Calculator.Rows(1).Cells(3).Value = Math.Round(add, 2)
                ElseIf rstObj.CategoryName = "Misc. Parts".ToUpper Then
                    DGV_Calculator.Rows(1).Cells(1).Value = "Final PTime"
                    DGV_Calculator.Rows(1).Cells(3).Value = Math.Round(add, 2)
                End If
                BtnReset.Enabled = True
            End If
            If BtnCalculateAndSave.Text = "Save" Then
                BtnCalculateAndSave.Text = "Calculate"
                BtnCalculateAndSave.Enabled = False
                BtnReset.Enabled = False
                If rstObj.CategoryName = "Assembly".ToUpper Then
                    rstObj.dgvCalculatorRowIndex.Cells(2).Value = DGV_Calculator.Rows(1).Cells(3).Value
                ElseIf rstObj.CategoryName = "SheetMetal".ToUpper Then
                    rstObj.dgvCalculatorRowIndex.Cells(2).Value = DGV_Calculator.Rows(3).Cells(3).Value
                ElseIf rstObj.CategoryName = "Structure".ToUpper Then
                    rstObj.dgvCalculatorRowIndex.Cells(2).Value = DGV_Calculator.Rows(1).Cells(3).Value
                ElseIf rstObj.CategoryName = "Misc. Parts".ToUpper Then
                    rstObj.dgvCalculatorRowIndex.Cells(2).Value = DGV_Calculator.Rows(1).Cells(3).Value
                End If

                DGV_Calculator.Columns.Clear()
                DGV_Calculator.Rows.Clear()
                Exit Sub
            End If
            BtnCalculateAndSave.Text = "Save"
        Catch ex As Exception
            CustomLogUtil.Log($"While Clicking on {BtnCalculateAndSave.Text} BUTTON... ", ex.Message, ex.StackTrace)
            MessageBox.Show($"While Clicking on {BtnCalculateAndSave.Text} BUTTON{vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}", "Message")
        End Try
    End Sub

    Private Sub BtnReset_Click(sender As Object, e As EventArgs) Handles BtnReset.Click
        ResetData()
    End Sub
    Public Sub ResetData()
        For i = 0 To DGV_Calculator.Rows.Count - 1
            DGV_Calculator.Rows(i).Cells(0).Value = Nothing
        Next
        BtnCalculateAndSave.Text = "Calculate"

        If rstObj.CategoryName = "Assembly".ToUpper Then
            DGV_Calculator.Rows(1).Cells(1).Value = ""
            DGV_Calculator.Rows(1).Cells(3).Value = ""
        ElseIf rstObj.CategoryName = "SheetMetal".ToUpper Then
            DGV_Calculator.Rows(3).Cells(1).Value = ""
            DGV_Calculator.Rows(3).Cells(3).Value = ""
        ElseIf rstObj.CategoryName = "Structure".ToUpper Then
            DGV_Calculator.Rows(1).Cells(1).Value = ""
            DGV_Calculator.Rows(1).Cells(3).Value = ""
        ElseIf rstObj.CategoryName = "Misc. Parts".ToUpper Then
            DGV_Calculator.Rows(1).Cells(1).Value = ""
            DGV_Calculator.Rows(1).Cells(3).Value = ""
        End If
        BtnReset.Enabled = False
    End Sub

    Private Sub RST_Design1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        DeleteDir()
    End Sub
    Public Sub DeleteDir()
        'PictureBox1.Image.Dispose()
        'PictureBox1.Image = Nothing
        Dim mypath As New System.IO.DirectoryInfo(System.IO.Path.Combine(Environment.CurrentDirectory, ".."))
        If Directory.Exists(mypath.FullName + "\Thumbnail") Then
            Dim files As String() = Directory.GetFiles(mypath.FullName + "\Thumbnail")
            Dim dirs As String() = Directory.GetDirectories(mypath.FullName + "\Thumbnail")
            For Each file As String In files

                System.IO.File.SetAttributes(file, FileAttributes.Normal)
                System.IO.File.Delete(file)
            Next
            Directory.Delete(mypath.FullName + "\Thumbnail")
        End If
    End Sub

    Private Sub BtnGetDataAndReset_Click(sender As Object, e As EventArgs) Handles BtnGetDataAndReset.Click
        GetDataValidation()

        GetDataAndReset()
    End Sub
    Public Sub GetDataAndReset()
        If BtnGetDataAndReset.Text = "Get Data" Then
            WaitStartSave()
            MtcReport(rstObj)
            ShowAllDGV()
            Lbldeclaration()

            WaitEndSave()

        ElseIf BtnGetDataAndReset.Text = "Reset" Then
            cmbProcess.Enabled = False
            BtnPreview.Enabled = False
            'If cmbProcess.Enabled = True And BtnUp.Enabled = True And BtnDown.Enabled = True And BtnDelete.Enabled = True And BtnPreview.Enabled = True Then
            '    BtnApplyValues.Enabled = False
            '    cmbProcess.Enabled = False
            '    BtnUp.Enabled = False
            '    BtnDown.Enabled = False
            '    BtnDelete.Enabled = False
            '    BtnPreview.Enabled = False
            '    BtnApplyValues.Visible = True

            'End If

            BtnEnable()
            If BtnUp.Enabled = True And BtnDown.Enabled = True And BtnDelete.Enabled = True Then
                BtnEnable()
            End If
            Reset()
            DGV_Calculator.Columns.Clear()
            rstObj.dtProcess.Clear()
            rstObj.dtProcess.Columns.Clear()
            ReleaseObject(rstObj)

            '  ResetAll()
            BtnGetDataAndReset.Text = "Get Data"
            Exit Sub
        End If
        BtnGetDataAndReset.Text = "Reset"
    End Sub
    Public Sub GetDataValidation()
        If Not cmbCategory.SelectedItem = "" Then
            rstObj.CategoryName = cmbCategory.SelectedItem
            cmbCategory.Enabled = False
        End If

        If txtFilename.Text = "" Or txtFoldername.Text = "" Or cmbCategory.SelectedItem = "" Then
            MsgBox("Select Filename or FolderName or Category")
            'SelectCategory(rstObj)
            Reset()
            Exit Sub
        End If
    End Sub
    Public Sub ResetAll()
        rstObj.i = Nothing
        rstObj.excelFilepath = Nothing
        rstObj.dir = Nothing
        rstObj.PartName = Nothing
        rstObj.MaterialDescription = Nothing
        rstObj.FilePath = Nothing

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

End Class