Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports ExcelDataReader
Imports Microsoft.Office.Interop.Excel
Imports NLog
Imports SolidEdgeAssembly
Imports SolidEdgeDraft
Imports SolidEdgeFramework
Imports SolidEdgePart
Imports WK.Libraries.BetterFolderBrowserNS

Public Class MTC_MTR_ReviewForm2

    Public Shared log As Logger = LogManager.GetCurrentClassLogger()

    Dim objApp As SolidEdgeFramework.Application = Nothing
    Dim objAssemblyDocument As SolidEdgeAssembly.AssemblyDocument = Nothing
    Dim dtAssemblyData As Data.DataTable
    Dim dtfilter As Data.DataTable
    Dim projectnamelst As New List(Of String)
    Dim authorsList As New List(Of String)
    Dim dicData As New Dictionary(Of String, DataSet)()

    Dim dtM2M As New Data.DataTable("M2M")

    Dim excelcol As Integer = 3

    'Dim columnName As Char = "C"c
    Dim columnChar As Char = "Z"

    Dim color1 As Color = Color.FromArgb(226, 239, 218)
    Dim color2 As Color = Color.FromArgb(252, 228, 214)
    Dim MainAsmBomCount As Integer = 0
    Dim mtcMtrModelObj As New MTC_MTR_Model()
    Dim dt As New Data.DataTable

    '17th Sep 2024
    Dim activeDocument As SolidEdgeFramework.SolidEdgeDocument = Nothing
    Dim objPartDocument As SolidEdgePart.PartDocument = Nothing
    Dim objSheetMetalDocument As SolidEdgePart.SheetMetalDocument = Nothing

#Region "Decide report selection"

    Private Sub ChkAll_CheckedChanged(sender As Object, e As EventArgs) Handles chkAll.CheckedChanged
        chkAssembly.Checked = False
        chkPart.Checked = False
        chkSheetMetal.Checked = False
    End Sub

    Private Sub ChkAssembly_CheckedChanged(sender As Object, e As EventArgs) Handles chkAssembly.CheckedChanged
        ReportSelection()
    End Sub

    Private Sub ChkPart_CheckedChanged(sender As Object, e As EventArgs) Handles chkPart.CheckedChanged
        ReportSelection()
    End Sub

    Private Sub ChkSheetMetal_CheckedChanged(sender As Object, e As EventArgs) Handles chkSheetMetal.CheckedChanged
        ReportSelection()
    End Sub

    Private Sub ReportSelection()

        If chkAssembly.Checked = True Or chkPart.Checked = True Or chkSheetMetal.Checked = True Then

            chkAll.Checked = False

        End If

    End Sub

#End Region

#Region "Select Baseline Directory"

    Private Sub BtnBrowseDirPath_Click(sender As Object, e As EventArgs) Handles BtnBrowseBaselineDirPath.Click
        Try
            Dim folderpath As String = ""
            folderpath = BrowseFolderAdvanced()
            If Not folderpath = String.Empty Then
                txtBaseLineDirectoryPath.Text = folderpath
            End If
        Catch ex1 As Exception
            Try
                Dim MyFolderBrowser As New Ookii.Dialogs.VistaFolderBrowserDialog
                Dim dlgResult As DialogResult = MyFolderBrowser.ShowDialog
                If dlgResult = DialogResult.OK Then
                    txtBaseLineDirectoryPath.Text = MyFolderBrowser.SelectedPath
                End If
            Catch ex As Exception
                If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
                    Dim path As String = FolderBrowserDialog1.SelectedPath
                    txtBaseLineDirectoryPath.Text = path
                End If
            End Try

        End Try
    End Sub

    Public Shared Function BrowseFolderAdvanced() As String

        Dim folderpath As String = ""
        Try
            Dim BetterFolderBrowser As New BetterFolderBrowser With {
                .Title = "Select folders",
                .RootFolder = "C:\\",
                .Multiselect = False
            }
            If BetterFolderBrowser.ShowDialog() = DialogResult.OK Then
                folderpath = BetterFolderBrowser.SelectedFolder
            End If
        Catch ex As Exception
            MsgBox("Advanced Browse folder" + ex.Message + vbNewLine + ex.StackTrace)
            CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
            MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
        End Try

        Return folderpath
    End Function

#End Region

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

    Private Sub MTC_MTR_ReviewForm2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        log.Info("============")
        log.Info("MTC/MTR Form Open.....")
        log.Info("============")
        CustomLogUtil.Heading("MTC/MTR Form Open.....")
        SolidEdgeCommunity.OleMessageFilter.Register()

        objApp = SolidEdgeCommunity.SolidEdgeUtils.Connect(True, False)
        'objApp = Marshal.GetActiveObject("SolidEdge.Application")

        objApp.DisplayAlerts = False

        SetControls(dgvDocumentDetails)
        chkAll.Checked = True
        Me.Text = $"{Me.Text} ({GlobalEntity.Version}) "

        txtBaseLineDirectoryPath.Text = Config.configObj.baselineDirectoryPath
        txtExportDirLocationMTR.Text = Config.configObj.mtcMtrReportsExportDirLocation
        '2nd Sep 2024
        txtExportDirLocationRouting.Text = Config.configObj.RoutingSequenceOutputDirectory
    End Sub

    Private Sub SetControls(ByRef DataGridViewComments As DataGridView)
        For Each col As DataGridViewColumn In DataGridViewComments.Columns
            col.SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        DataGridViewComments.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.GhostWhite
        DataGridViewComments.AlternatingRowsDefaultCellStyle.ForeColor = System.Drawing.Color.DarkBlue

        DataGridViewComments.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(6, 150, 215) ' Drawing.Color.WhiteSmoke
        DataGridViewComments.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.WhiteSmoke 'Drawing.Color.OrangeRed 'FromArgb(6, 150, 215) '

        DataGridViewComments.AllowUserToAddRows = False
        DataGridViewComments.AllowUserToDeleteRows = False
        DataGridViewComments.AllowUserToResizeRows = False
        DataGridViewComments.AllowUserToResizeColumns = False
        DataGridViewComments.RowHeadersVisible = False
        DataGridViewComments.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect 'CellSelect
        DataGridViewComments.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        'DataGridViewComments.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        DataGridViewComments.EnableHeadersVisualStyles = True
        DataGridViewComments.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
        DataGridViewComments.MultiSelect = False
        DataGridViewComments.ReadOnly = False
    End Sub

    Private Function GetM2MData(ByVal dirPath As String) As System.Data.DataTable
        '9th Sep 2024 ' Added try..catch
        Try
            'dirPath is skipped because m2mfile get throgh config object
            'Dim mtcExcelPath As String = IO.Path.Combine(dirPath, "M2MData.csv")
            Dim mtcExcelPath As String = Config.configObj.m2MFile
            dicData = ExcelUtil.ReadM2Mfile_CSV(mtcExcelPath)
            Dim ds As DataSet = dicData("Sheet1")
            Return ds.Tables(0)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

    Public Function Readfile(ByVal file As String) As List(Of String)
        '9th Sep 2024 'Added try..catch
        Try
            Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(file)
            Dim a As String

            Dim lst As New List(Of String)
            Do
                a = reader.ReadLine
                If lst.Contains("Begin Project") Then
                    Dim split As String() = a.Split(";")

                    lst.Add(split(0))
                End If
                If a = "Begin Project" Then
                    Dim split As String() = a.Split(";")

                    lst.Add(split(0))
                End If

                '
                ' Code here
                '
            Loop Until a = "End Project"

            reader.Close()
            Return lst
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function


    Public Function ReadfileAuthor(ByVal file As String) As List(Of String)
        '9th Sep 2024 'Added try..catch

        Try
            Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(file)
            Dim a As String

            Dim lst As New List(Of String)
            Do
                a = reader.ReadLine
                If lst.Contains("Begin Author") Then
                    Dim split As String() = a.Split(";")

                    lst.Add(split(0))
                End If
                If a = "Begin Author" Then
                    Dim split As String() = a.Split(";")

                    lst.Add(split(0))
                End If

                '
                ' Code here
                '
            Loop Until a = "End Author"

            reader.Close()
            Return lst
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Function

    Private Function GetPropValue(ByVal prop1 As [Property])
        Dim value As String = String.Empty
        If prop1.Value IsNot Nothing Then
            value = prop1.Value.ToString().Trim()
        End If
        Return value
    End Function

    Private Function GetMainAssemblyData(ByVal mtcMtrModelObj As MTC_MTR_Model) As Dictionary(Of String, String)
        '9th Sep 2024 'Added try..catch
        Try

            Dim dicProperties As New Dictionary(Of String, String)()

            objAssemblyDocument = objApp.ActiveDocument

            'Custom
            If objAssemblyDocument IsNot Nothing Then

                Debug.Print("Custom")

                Dim propSets As SolidEdgeFramework.PropertySets = objAssemblyDocument.Properties

                Dim custProps As SolidEdgeFramework.Properties = propSets.Item("Custom")

                For Each prop1 As [Property] In custProps

                    Try
                        Debug.Print($"{prop1.Name} > {GetPropValue(prop1)}")
                        dicProperties.Add(prop1.Name, GetPropValue(prop1))
                    Catch ex As Exception
                        MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
                        CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
                    End Try

                Next
            End If

            'Document Summary
            If objAssemblyDocument IsNot Nothing Then

                Debug.Print("DocumentSummaryInformation")

                Dim propSets As SolidEdgeFramework.PropertySets = objAssemblyDocument.Properties

                Dim custProps As SolidEdgeFramework.Properties = propSets.Item("DocumentSummaryInformation")

                For Each prop1 As [Property] In custProps

                    Try
                        Debug.Print($"{prop1.Name} > {GetPropValue(prop1)}")
                        dicProperties.Add(prop1.Name, GetPropValue(prop1))
                    Catch ex As Exception
                        MTC_MTR_ReviewForm2.log.Error($"Create documentsummary property collection: {ex.Message} {vbNewLine} {ex.StackTrace}")
                    End Try

                Next
            End If

            'Project Info
            If objAssemblyDocument IsNot Nothing Then

                Debug.Print("ProjectInformation")

                Dim propSets As SolidEdgeFramework.PropertySets = objAssemblyDocument.Properties

                Dim custProps As SolidEdgeFramework.Properties = propSets.Item("ProjectInformation")

                For Each prop1 As [Property] In custProps

                    Try
                        Debug.Print($"{prop1.Name} > {GetPropValue(prop1)}")
                        dicProperties.Add(prop1.Name, GetPropValue(prop1))
                    Catch ex As Exception
                        MTC_MTR_ReviewForm2.log.Error($"Create project information property collection: {ex.Message}{vbNewLine}{ex.StackTrace}")
                    End Try

                Next
            End If

            'Summary Info
            If objAssemblyDocument IsNot Nothing Then

                Debug.Print("SummaryInformation")

                Dim propSets As SolidEdgeFramework.PropertySets = objAssemblyDocument.Properties

                Dim custProps As SolidEdgeFramework.Properties = propSets.Item("SummaryInformation")

                For Each prop1 As [Property] In custProps

                    Try




                        If (prop1.Name = "Title") Then
                            Dim TitleValue As String = Strings.Left(prop1.Value.ToString(), 35)
                            dicProperties.Add(prop1.Name, TitleValue)
                            Debug.Print($"{prop1.Name} > {TitleValue}")
                        Else
                            dicProperties.Add(prop1.Name, GetPropValue(prop1))
                            Debug.Print($"{prop1.Name} > {GetPropValue(prop1)}")
                        End If

                        'If (prop1.Name = "Author") Then
                        'Dim dv As New DataView(mtcMtrModelObj.dtAuthorData)
                        'Dim WindowsAuthorName As String = "Windows Usernames"
                        ''dv.RowFilter = $"{partno}='{mainAssemblyDocModel.fileNameWithoutExt}'"
                        'Dim OldValue = GetPropValue(prop1)

                        'Dim filter As String = $"Convert([{WindowsAuthorName}], 'System.String') = '{GetPropValue(prop1)}'"
                        'dv.RowFilter = filter '$"{materialUsedCol}='{cmbMaterialUsed2_Mw.Text}'"

                        'For Each drv As DataRowView In dv
                        '    'mtcMtrModelObj.authorList.Remove(OldValue)
                        '    prop1.Value = If(drv("Full Name") = Nothing, "", drv("Full Name"))
                        '    Dim NewValue = prop1.Value
                        '    custProps.Save()

                        '    dicProperties.Add(prop1.Name, NewValue)
                        '    If (Not mtcMtrModelObj.authorList.Contains(prop1.Value)) Then
                        '        mtcMtrModelObj.authorList.Add(prop1.Value)
                        '    End If


                        '    Exit For
                        'Next

                        'custProps.Save()
                        'objAssemblyDocument.Save()


                        'End If





                    Catch ex As Exception
                        MTC_MTR_ReviewForm2.log.Error($"Create summary information collection: {ex.Message}{vbNewLine}{ex.StackTrace}")
                        CustomLogUtil.Log("Create summary information collection:", ex.Message, ex.StackTrace)
                    End Try

                Next
#Region "Author Update for all parts"
                'Dim objDoc As SolidEdgeAssembly.AssemblyDocument
                'Dim objParts As SolidEdgeAssembly.Occurrences


                'objDoc = objApp.ActiveDocument

                '' Getting the parts objects of the AssemblyDocument object.
                'objParts = objDoc.Occurrences
                'Dim a As Integer
                'Dim listofparts As New List(Of String)
                'For a = 1 To objParts.Count
                '    If (Not listofparts.Contains(objParts.Item(a).PartFileName)) Then
                '        listofparts.Add(objParts.Item(a).PartFileName)

                '    End If
                'Next
                'For a = 0 To listofparts.Count - 1


                '    Dim objApp As SolidEdgeFramework.Application = System.Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application")

                '    Dim objDocument As SolidEdgeFramework.SolidEdgeDocument = objApp.Documents.Open(listofparts.Item(a))



                '    propSets = objDocument.Properties

                '    custProps = propSets.Item("SummaryInformation")

                '    For Each prop2 As [Property] In custProps

                '        Try



                '            If (prop2.Name = "Author") Then


                '                Dim dv As New DataView(mtcMtrModelObj.dtAuthorData)
                '                Dim WindowsAuthorName As String = "Windows Usernames"
                '                'dv.RowFilter = $"{partno}='{mainAssemblyDocModel.fileNameWithoutExt}'"

                '                Dim OldValue = GetPropValue(prop2)
                '                Dim filter As String = $"Convert([{WindowsAuthorName}], 'System.String') = '{GetPropValue(prop2)}'"
                '                dv.RowFilter = filter '$"{materialUsedCol}='{cmbMaterialUsed2_Mw.Text}'"

                '                For Each drv As DataRowView In dv
                '                    'mtcMtrModelObj.authorList.Remove(OldValue)
                '                    prop2.Value = If(drv("Full Name") = Nothing, "", drv("Full Name"))
                '                    custProps.Save()
                '                    Dim NewValue = prop2.Value
                '                    'dicProperties.Add(prop2.Name, NewValue)
                '                    If (Not mtcMtrModelObj.authorList.Contains(prop2.Value)) Then
                '                        mtcMtrModelObj.authorList.Add(prop2.Value)
                '                    End If


                '                    Exit For
                '                Next
                '                custProps.Save()




                '            End If

                '        Catch ex As Exception

                '        End Try
                '    Next
                '    objDocument.Save()
                '    objDocument.Close()

                'Next

#End Region


            End If

            Return dicProperties
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Function


    '17th Sep 2024
    Private Function GetPartPropertyData(ByVal mtcMtrModelObj As MTC_MTR_Model) As Dictionary(Of String, String)
        '9th Sep 2024 'Added try..catch
        Try

            Dim dicProperties As New Dictionary(Of String, String)()

            objPartDocument = objApp.ActiveDocument

            'Custom
            If objPartDocument IsNot Nothing Then

                Debug.Print("Custom")

                Dim propSets As SolidEdgeFramework.PropertySets = objPartDocument.Properties

                Dim custProps As SolidEdgeFramework.Properties = propSets.Item("Custom")

                For Each prop1 As [Property] In custProps

                    Try
                        Debug.Print($"{prop1.Name} > {GetPropValue(prop1)}")
                        dicProperties.Add(prop1.Name, GetPropValue(prop1))
                    Catch ex As Exception
                        MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
                        CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
                    End Try

                Next
            End If

            'Document Summary
            If objPartDocument IsNot Nothing Then

                Debug.Print("DocumentSummaryInformation")

                Dim propSets As SolidEdgeFramework.PropertySets = objPartDocument.Properties

                Dim custProps As SolidEdgeFramework.Properties = propSets.Item("DocumentSummaryInformation")

                For Each prop1 As [Property] In custProps

                    Try
                        Debug.Print($"{prop1.Name} > {GetPropValue(prop1)}")
                        dicProperties.Add(prop1.Name, GetPropValue(prop1))
                    Catch ex As Exception
                        MTC_MTR_ReviewForm2.log.Error($"Create documentsummary property collection: {ex.Message} {vbNewLine} {ex.StackTrace}")
                    End Try

                Next
            End If

            'Project Info
            If objPartDocument IsNot Nothing Then

                Debug.Print("ProjectInformation")

                Dim propSets As SolidEdgeFramework.PropertySets = objPartDocument.Properties

                Dim custProps As SolidEdgeFramework.Properties = propSets.Item("ProjectInformation")

                For Each prop1 As [Property] In custProps

                    Try
                        Debug.Print($"{prop1.Name} > {GetPropValue(prop1)}")
                        dicProperties.Add(prop1.Name, GetPropValue(prop1))
                    Catch ex As Exception
                        MTC_MTR_ReviewForm2.log.Error($"Create project information property collection: {ex.Message}{vbNewLine}{ex.StackTrace}")
                    End Try

                Next
            End If

            'Summary Info
            If objPartDocument IsNot Nothing Then

                Debug.Print("SummaryInformation")

                Dim propSets As SolidEdgeFramework.PropertySets = objPartDocument.Properties

                Dim custProps As SolidEdgeFramework.Properties = propSets.Item("SummaryInformation")

                For Each prop1 As [Property] In custProps

                    Try




                        If (prop1.Name = "Title") Then
                            Dim TitleValue As String = Strings.Left(prop1.Value.ToString(), 35)
                            dicProperties.Add(prop1.Name, TitleValue)
                            Debug.Print($"{prop1.Name} > {TitleValue}")
                        Else
                            dicProperties.Add(prop1.Name, GetPropValue(prop1))
                            Debug.Print($"{prop1.Name} > {GetPropValue(prop1)}")
                        End If

                        'If (prop1.Name = "Author") Then
                        'Dim dv As New DataView(mtcMtrModelObj.dtAuthorData)
                        'Dim WindowsAuthorName As String = "Windows Usernames"
                        ''dv.RowFilter = $"{partno}='{mainAssemblyDocModel.fileNameWithoutExt}'"
                        'Dim OldValue = GetPropValue(prop1)

                        'Dim filter As String = $"Convert([{WindowsAuthorName}], 'System.String') = '{GetPropValue(prop1)}'"
                        'dv.RowFilter = filter '$"{materialUsedCol}='{cmbMaterialUsed2_Mw.Text}'"

                        'For Each drv As DataRowView In dv
                        '    'mtcMtrModelObj.authorList.Remove(OldValue)
                        '    prop1.Value = If(drv("Full Name") = Nothing, "", drv("Full Name"))
                        '    Dim NewValue = prop1.Value
                        '    custProps.Save()

                        '    dicProperties.Add(prop1.Name, NewValue)
                        '    If (Not mtcMtrModelObj.authorList.Contains(prop1.Value)) Then
                        '        mtcMtrModelObj.authorList.Add(prop1.Value)
                        '    End If


                        '    Exit For
                        'Next

                        'custProps.Save()
                        'objAssemblyDocument.Save()


                        'End If





                    Catch ex As Exception
                        MTC_MTR_ReviewForm2.log.Error($"Create summary information collection: {ex.Message}{vbNewLine}{ex.StackTrace}")
                        CustomLogUtil.Log("Create summary information collection:", ex.Message, ex.StackTrace)
                    End Try

                Next
#Region "Author Update for all parts"
                'Dim objDoc As SolidEdgeAssembly.AssemblyDocument
                'Dim objParts As SolidEdgeAssembly.Occurrences


                'objDoc = objApp.ActiveDocument

                '' Getting the parts objects of the AssemblyDocument object.
                'objParts = objDoc.Occurrences
                'Dim a As Integer
                'Dim listofparts As New List(Of String)
                'For a = 1 To objParts.Count
                '    If (Not listofparts.Contains(objParts.Item(a).PartFileName)) Then
                '        listofparts.Add(objParts.Item(a).PartFileName)

                '    End If
                'Next
                'For a = 0 To listofparts.Count - 1


                '    Dim objApp As SolidEdgeFramework.Application = System.Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application")

                '    Dim objDocument As SolidEdgeFramework.SolidEdgeDocument = objApp.Documents.Open(listofparts.Item(a))



                '    propSets = objDocument.Properties

                '    custProps = propSets.Item("SummaryInformation")

                '    For Each prop2 As [Property] In custProps

                '        Try



                '            If (prop2.Name = "Author") Then


                '                Dim dv As New DataView(mtcMtrModelObj.dtAuthorData)
                '                Dim WindowsAuthorName As String = "Windows Usernames"
                '                'dv.RowFilter = $"{partno}='{mainAssemblyDocModel.fileNameWithoutExt}'"

                '                Dim OldValue = GetPropValue(prop2)
                '                Dim filter As String = $"Convert([{WindowsAuthorName}], 'System.String') = '{GetPropValue(prop2)}'"
                '                dv.RowFilter = filter '$"{materialUsedCol}='{cmbMaterialUsed2_Mw.Text}'"

                '                For Each drv As DataRowView In dv
                '                    'mtcMtrModelObj.authorList.Remove(OldValue)
                '                    prop2.Value = If(drv("Full Name") = Nothing, "", drv("Full Name"))
                '                    custProps.Save()
                '                    Dim NewValue = prop2.Value
                '                    'dicProperties.Add(prop2.Name, NewValue)
                '                    If (Not mtcMtrModelObj.authorList.Contains(prop2.Value)) Then
                '                        mtcMtrModelObj.authorList.Add(prop2.Value)
                '                    End If


                '                    Exit For
                '                Next
                '                custProps.Save()




                '            End If

                '        Catch ex As Exception

                '        End Try
                '    Next
                '    objDocument.Save()
                '    objDocument.Close()

                'Next

#End Region


            End If

            Return dicProperties
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Function


    '17th Sep 2024
    Private Function GetSheetMetalPropertyData(ByVal mtcMtrModelObj As MTC_MTR_Model) As Dictionary(Of String, String)
        '9th Sep 2024 'Added try..catch
        Try

            Dim dicProperties As New Dictionary(Of String, String)()

            objSheetMetalDocument = objApp.ActiveDocument

            'Custom
            If objSheetMetalDocument IsNot Nothing Then

                Debug.Print("Custom")

                Dim propSets As SolidEdgeFramework.PropertySets = objSheetMetalDocument.Properties

                Dim custProps As SolidEdgeFramework.Properties = propSets.Item("Custom")

                For Each prop1 As [Property] In custProps

                    Try
                        Debug.Print($"{prop1.Name} > {GetPropValue(prop1)}")
                        dicProperties.Add(prop1.Name, GetPropValue(prop1))
                    Catch ex As Exception
                        MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
                        CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
                    End Try

                Next
            End If

            'Document Summary
            If objSheetMetalDocument IsNot Nothing Then

                Debug.Print("DocumentSummaryInformation")

                Dim propSets As SolidEdgeFramework.PropertySets = objSheetMetalDocument.Properties

                Dim custProps As SolidEdgeFramework.Properties = propSets.Item("DocumentSummaryInformation")

                For Each prop1 As [Property] In custProps

                    Try
                        Debug.Print($"{prop1.Name} > {GetPropValue(prop1)}")
                        dicProperties.Add(prop1.Name, GetPropValue(prop1))
                    Catch ex As Exception
                        MTC_MTR_ReviewForm2.log.Error($"Create documentsummary property collection: {ex.Message} {vbNewLine} {ex.StackTrace}")
                    End Try

                Next
            End If

            'Project Info
            If objSheetMetalDocument IsNot Nothing Then

                Debug.Print("ProjectInformation")

                Dim propSets As SolidEdgeFramework.PropertySets = objSheetMetalDocument.Properties

                Dim custProps As SolidEdgeFramework.Properties = propSets.Item("ProjectInformation")

                For Each prop1 As [Property] In custProps

                    Try
                        Debug.Print($"{prop1.Name} > {GetPropValue(prop1)}")
                        dicProperties.Add(prop1.Name, GetPropValue(prop1))
                    Catch ex As Exception
                        MTC_MTR_ReviewForm2.log.Error($"Create project information property collection: {ex.Message}{vbNewLine}{ex.StackTrace}")
                    End Try

                Next
            End If

            'Summary Info
            If objSheetMetalDocument IsNot Nothing Then

                Debug.Print("SummaryInformation")

                Dim propSets As SolidEdgeFramework.PropertySets = objSheetMetalDocument.Properties

                Dim custProps As SolidEdgeFramework.Properties = propSets.Item("SummaryInformation")

                For Each prop1 As [Property] In custProps

                    Try

                        If (prop1.Name = "Title") Then
                            Dim TitleValue As String = Strings.Left(prop1.Value.ToString(), 35)
                            dicProperties.Add(prop1.Name, TitleValue)
                            Debug.Print($"{prop1.Name} > {TitleValue}")
                        Else
                            dicProperties.Add(prop1.Name, GetPropValue(prop1))
                            Debug.Print($"{prop1.Name} > {GetPropValue(prop1)}")
                        End If

                        'If (prop1.Name = "Author") Then
                        'Dim dv As New DataView(mtcMtrModelObj.dtAuthorData)
                        'Dim WindowsAuthorName As String = "Windows Usernames"
                        ''dv.RowFilter = $"{partno}='{mainAssemblyDocModel.fileNameWithoutExt}'"
                        'Dim OldValue = GetPropValue(prop1)

                        'Dim filter As String = $"Convert([{WindowsAuthorName}], 'System.String') = '{GetPropValue(prop1)}'"
                        'dv.RowFilter = filter '$"{materialUsedCol}='{cmbMaterialUsed2_Mw.Text}'"

                        'For Each drv As DataRowView In dv
                        '    'mtcMtrModelObj.authorList.Remove(OldValue)
                        '    prop1.Value = If(drv("Full Name") = Nothing, "", drv("Full Name"))
                        '    Dim NewValue = prop1.Value
                        '    custProps.Save()

                        '    dicProperties.Add(prop1.Name, NewValue)
                        '    If (Not mtcMtrModelObj.authorList.Contains(prop1.Value)) Then
                        '        mtcMtrModelObj.authorList.Add(prop1.Value)
                        '    End If


                        '    Exit For
                        'Next

                        'custProps.Save()
                        'objAssemblyDocument.Save()


                        'End If

                    Catch ex As Exception
                        MTC_MTR_ReviewForm2.log.Error($"Create summary information collection: {ex.Message}{vbNewLine}{ex.StackTrace}")
                        CustomLogUtil.Log("Create summary information collection:", ex.Message, ex.StackTrace)
                    End Try

                Next
#Region "Author Update for all parts"
                'Dim objDoc As SolidEdgeAssembly.AssemblyDocument
                'Dim objParts As SolidEdgeAssembly.Occurrences


                'objDoc = objApp.ActiveDocument

                '' Getting the parts objects of the AssemblyDocument object.
                'objParts = objDoc.Occurrences
                'Dim a As Integer
                'Dim listofparts As New List(Of String)
                'For a = 1 To objParts.Count
                '    If (Not listofparts.Contains(objParts.Item(a).PartFileName)) Then
                '        listofparts.Add(objParts.Item(a).PartFileName)

                '    End If
                'Next
                'For a = 0 To listofparts.Count - 1


                '    Dim objApp As SolidEdgeFramework.Application = System.Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application")

                '    Dim objDocument As SolidEdgeFramework.SolidEdgeDocument = objApp.Documents.Open(listofparts.Item(a))



                '    propSets = objDocument.Properties

                '    custProps = propSets.Item("SummaryInformation")

                '    For Each prop2 As [Property] In custProps

                '        Try



                '            If (prop2.Name = "Author") Then


                '                Dim dv As New DataView(mtcMtrModelObj.dtAuthorData)
                '                Dim WindowsAuthorName As String = "Windows Usernames"
                '                'dv.RowFilter = $"{partno}='{mainAssemblyDocModel.fileNameWithoutExt}'"

                '                Dim OldValue = GetPropValue(prop2)
                '                Dim filter As String = $"Convert([{WindowsAuthorName}], 'System.String') = '{GetPropValue(prop2)}'"
                '                dv.RowFilter = filter '$"{materialUsedCol}='{cmbMaterialUsed2_Mw.Text}'"

                '                For Each drv As DataRowView In dv
                '                    'mtcMtrModelObj.authorList.Remove(OldValue)
                '                    prop2.Value = If(drv("Full Name") = Nothing, "", drv("Full Name"))
                '                    custProps.Save()
                '                    Dim NewValue = prop2.Value
                '                    'dicProperties.Add(prop2.Name, NewValue)
                '                    If (Not mtcMtrModelObj.authorList.Contains(prop2.Value)) Then
                '                        mtcMtrModelObj.authorList.Add(prop2.Value)
                '                    End If


                '                    Exit For
                '                Next
                '                custProps.Save()




                '            End If

                '        Catch ex As Exception

                '        End Try
                '    Next
                '    objDocument.Save()
                '    objDocument.Close()

                'Next

#End Region


            End If

            Return dicProperties
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

    Private Function SetMainAssemblyData(ByVal mtcMtrModelObj As MTC_MTR_Model, ByVal dicProperties As Dictionary(Of String, String)) As DocumentModel
        Dim mtcReviewObj As New DocumentModel With {
            .fileNameWithoutExt = IO.Path.GetFileNameWithoutExtension(objAssemblyDocument.FullName) 'Active document path
            }
        mtcReviewObj.fileName = mtcReviewObj.fileNameWithoutExt
        mtcReviewObj.revisionNumber_Prop = If(dicProperties.ContainsKey("Revision"), dicProperties("Revision"), "")
        mtcReviewObj.Title = If(dicProperties.ContainsKey("Title"), dicProperties("Title"), "")
        mtcReviewObj.author = If(dicProperties.ContainsKey("Author"), dicProperties("Author"), "") 'dicProperties("Author")
        mtcReviewObj.documentno = If(dicProperties.ContainsKey("Document Number"), dicProperties("Document Number"), "") ' dicProperties("Document Number")
        mtcReviewObj.comments = If(dicProperties.ContainsKey("Comments"), dicProperties("Comments"), "") ' dicProperties("Comments")
        mtcReviewObj.category = If(dicProperties.ContainsKey("Category"), dicProperties("Category"), "") ' dicProperties("Category")
        mtcReviewObj.isElectrical = False
        mtcReviewObj.materialused = If(dicProperties.ContainsKey("Material Used"), dicProperties("Material Used"), "") ' dicProperties("Material Used")
        mtcReviewObj.matlspec = If(dicProperties.ContainsKey("MATL SPEC"), dicProperties("MATL SPEC"), "") ' dicProperties("MATL SPEC")
        mtcReviewObj.isBaseline = False
        mtcReviewObj.fullpath = objAssemblyDocument.FullName
        mtcReviewObj.density = If(dicProperties.ContainsKey("Category"), dicProperties("Category"), "") ' dicProperties("Category")
        mtcReviewObj.projectname = If(dicProperties.ContainsKey("Project"), dicProperties("Project"), "") 'dicProperties("Project")
        If mtcReviewObj.projectname = "BROOKVILLE EQUIPMENT CORP" Then
            mtcReviewObj.isBrookVilleProject_Baseline = "Yes"
        End If
        mtcReviewObj.Documenttype = If(dicProperties.ContainsKey("Status Text"), dicProperties("Status Text"), "") ' dicProperties("Status Text")
        mtcReviewObj.UomProperty = If(dicProperties.ContainsKey("UOM"), dicProperties("UOM"), "") ' dicProperties("UOM")
        mtcReviewObj.keywords = If(dicProperties.ContainsKey("Keywords"), dicProperties("Keywords"), "") ' dicProperties("Keywords")
        mtcReviewObj.ECO = If(dicProperties.ContainsKey("ECO/SOW"), dicProperties("ECO/SOW"), "") ' dicProperties("ECO/SOW")
        mtcReviewObj.projectname = If(dicProperties.ContainsKey("Project Name"), dicProperties("Project Name"), "") ' dicProperties("ECO/SOW")
        mtcReviewObj.lastAuthor = If(dicProperties.ContainsKey("Last Author"), dicProperties("Last Author"), "")

        '13th Sep 2024

        If mtcReviewObj.fileName.Contains("_") Or mtcReviewObj.fileName.Contains("-") Then
            mtcReviewObj.revisionNumber_FileName = GetRevisionLevel(mtcReviewObj.fileName, mtcReviewObj.revisionNumber_Prop)
        Else
            mtcReviewObj.fileName = mtcReviewObj.fileNameWithoutExt
        End If


        'Project Name
        'mtcReviewObj.modifiedDate = If(dicProperties.ContainsKey("Modified"), dicProperties("Modified"), "") ' dicProperties("ECO/SOW")

        'We could not find the Modified property in assembly document
        'So, we have used last saved date.
        'Last Save Date
        Try
            mtcReviewObj.modifiedDate = If(dicProperties.ContainsKey("Last Save Date"), dicProperties("Last Save Date"), "") ' dicProperties("ECO/SOW")
            mtcReviewObj.modifiedDate = DateTime.Parse(mtcReviewObj.modifiedDate).ToShortDateString()
        Catch ex As Exception
            '9th Sep 2024 
            MsgBox(ex.Message)
        End Try

        Dim documents As SolidEdgeFramework.Documents = objApp.Documents

        Dim asmdoc As SolidEdgeAssembly.AssemblyDocument = objApp.ActiveDocument
        Try
            Dim occurrences As SolidEdgeAssembly.Occurrences = asmdoc.Occurrences
            Dim partlst As New List(Of String)

            For Each occur As SolidEdgeAssembly.Occurrence In occurrences
                If Not partlst.Contains(occur.Name) Then
                    partlst.Add(occur.Name)
                    mtcReviewObj.partlistcount += 1
                End If

            Next



            GetPartList()

            mtcReviewObj.partlistcount = GetBomCount()

            mtcReviewObj.checkAssemblyFeature = CheckAssemblyFeatureExistence(asmdoc)

            mtcReviewObj.interPartLink = CheckInterPartLinksPSM()

            mtcReviewObj.isGeometryBroken = IsGeomtryBroken_assembly(asmdoc)

            mtcReviewObj.qAQC = If(dicProperties.ContainsKey("QAQC"), dicProperties("QAQC"), "")

            mtcReviewObj.quantity = "1"

        Catch ex As Exception
            MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
            CustomLogUtil.Log("While Setting Main Assembly Data", ex.Message, ex.StackTrace)
            '9th Sep 2024 
            MsgBox(ex.Message)
        End Try

        Return mtcReviewObj
    End Function


    '17th Sep 2024
    Private Function SetPartPropertyData(ByVal mtcMtrModelObj As MTC_MTR_Model, ByVal dicProperties As Dictionary(Of String, String)) As DocumentModel
        Dim mtcReviewObj As New DocumentModel With {
            .fileNameWithoutExt = IO.Path.GetFileNameWithoutExtension(objPartDocument.FullName) 'Active document path
            }
        mtcReviewObj.fileName = mtcReviewObj.fileNameWithoutExt
        mtcReviewObj.revisionNumber_Prop = If(dicProperties.ContainsKey("Revision"), dicProperties("Revision"), "")
        mtcReviewObj.Title = If(dicProperties.ContainsKey("Title"), dicProperties("Title"), "")
        mtcReviewObj.author = If(dicProperties.ContainsKey("Author"), dicProperties("Author"), "") 'dicProperties("Author")
        mtcReviewObj.documentno = If(dicProperties.ContainsKey("Document Number"), dicProperties("Document Number"), "") ' dicProperties("Document Number")
        mtcReviewObj.comments = If(dicProperties.ContainsKey("Comments"), dicProperties("Comments"), "") ' dicProperties("Comments")
        mtcReviewObj.category = If(dicProperties.ContainsKey("Category"), dicProperties("Category"), "") ' dicProperties("Category")
        mtcReviewObj.isElectrical = False
        mtcReviewObj.materialused = If(dicProperties.ContainsKey("Material Used"), dicProperties("Material Used"), "") ' dicProperties("Material Used")
        mtcReviewObj.matlspec = If(dicProperties.ContainsKey("MATL SPEC"), dicProperties("MATL SPEC"), "") ' dicProperties("MATL SPEC")
        mtcReviewObj.isBaseline = False
        mtcReviewObj.fullpath = objPartDocument.FullName
        mtcReviewObj.density = If(dicProperties.ContainsKey("Category"), dicProperties("Category"), "") ' dicProperties("Category")
        mtcReviewObj.projectname = If(dicProperties.ContainsKey("Project"), dicProperties("Project"), "") 'dicProperties("Project")
        If mtcReviewObj.projectname = "BROOKVILLE EQUIPMENT CORP" Then
            mtcReviewObj.isBrookVilleProject_Baseline = "Yes"
        End If
        mtcReviewObj.Documenttype = If(dicProperties.ContainsKey("Status Text"), dicProperties("Status Text"), "") ' dicProperties("Status Text")
        mtcReviewObj.UomProperty = If(dicProperties.ContainsKey("UOM"), dicProperties("UOM"), "") ' dicProperties("UOM")
        mtcReviewObj.keywords = If(dicProperties.ContainsKey("Keywords"), dicProperties("Keywords"), "") ' dicProperties("Keywords")
        mtcReviewObj.ECO = If(dicProperties.ContainsKey("ECO/SOW"), dicProperties("ECO/SOW"), "") ' dicProperties("ECO/SOW")
        mtcReviewObj.projectname = If(dicProperties.ContainsKey("Project Name"), dicProperties("Project Name"), "") ' dicProperties("ECO/SOW")
        mtcReviewObj.lastAuthor = If(dicProperties.ContainsKey("Last Author"), dicProperties("Last Author"), "")

        '13th Sep 2024

        If mtcReviewObj.fileName.Contains("_") Or mtcReviewObj.fileName.Contains("-") Then
            mtcReviewObj.revisionNumber_FileName = GetRevisionLevel(mtcReviewObj.fileName, mtcReviewObj.revisionNumber_Prop)
        Else
            mtcReviewObj.fileName = mtcReviewObj.fileNameWithoutExt
        End If


        'Project Name
        'mtcReviewObj.modifiedDate = If(dicProperties.ContainsKey("Modified"), dicProperties("Modified"), "") ' dicProperties("ECO/SOW")

        'We could not find the Modified property in assembly document
        'So, we have used last saved date.
        'Last Save Date
        Try
            mtcReviewObj.modifiedDate = If(dicProperties.ContainsKey("Last Save Date"), dicProperties("Last Save Date"), "") ' dicProperties("ECO/SOW")
            mtcReviewObj.modifiedDate = DateTime.Parse(mtcReviewObj.modifiedDate).ToShortDateString()
        Catch ex As Exception
            '9th Sep 2024 
            MsgBox(ex.Message)
        End Try

        'Dim documents As SolidEdgeFramework.Documents = objApp.Documents

        'Dim partdoc As SolidEdgePart.PartDocument = objApp.ActiveDocument

        Try
            'Dim occurrences As SolidEdgeAssembly.Occurrences = partdoc.Occurrences
            'Dim partlst As New List(Of String)

            'For Each occur As SolidEdgeAssembly.Occurrence In occurrences
            '    If Not partlst.Contains(occur.Name) Then
            '        partlst.Add(occur.Name)
            '        mtcReviewObj.partlistcount += 1
            '    End If

            'Next



            GetPartListPartDoc()

            mtcReviewObj.partlistcount = GetBomCount()

            'mtcReviewObj.checkAssemblyFeature = CheckPartFeatureExistence(partdoc)

            'mtcReviewObj.interPartLink = CheckInterPartLinksPSM()

            'mtcReviewObj.isGeometryBroken = IsGeomtryBroken_assembly(partdoc)

            mtcReviewObj.qAQC = If(dicProperties.ContainsKey("QAQC"), dicProperties("QAQC"), "")

            mtcReviewObj.quantity = "1"

        Catch ex As Exception
            MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
            CustomLogUtil.Log("While Setting Main Assembly Data", ex.Message, ex.StackTrace)
            '9th Sep 2024 
            MsgBox(ex.Message)
        End Try

        Return mtcReviewObj
    End Function


    '17th Sep 2024
    Private Function SetSheetMetalPropertyData(ByVal mtcMtrModelObj As MTC_MTR_Model, ByVal dicProperties As Dictionary(Of String, String)) As DocumentModel
        Dim mtcReviewObj As New DocumentModel With {
            .fileNameWithoutExt = IO.Path.GetFileNameWithoutExtension(objSheetMetalDocument.FullName) 'Active document path
            }
        mtcReviewObj.fileName = mtcReviewObj.fileNameWithoutExt
        mtcReviewObj.revisionNumber_Prop = If(dicProperties.ContainsKey("Revision"), dicProperties("Revision"), "")
        mtcReviewObj.Title = If(dicProperties.ContainsKey("Title"), dicProperties("Title"), "")
        mtcReviewObj.author = If(dicProperties.ContainsKey("Author"), dicProperties("Author"), "") 'dicProperties("Author")
        mtcReviewObj.documentno = If(dicProperties.ContainsKey("Document Number"), dicProperties("Document Number"), "") ' dicProperties("Document Number")
        mtcReviewObj.comments = If(dicProperties.ContainsKey("Comments"), dicProperties("Comments"), "") ' dicProperties("Comments")
        mtcReviewObj.category = If(dicProperties.ContainsKey("Category"), dicProperties("Category"), "") ' dicProperties("Category")
        mtcReviewObj.isElectrical = False
        mtcReviewObj.materialused = If(dicProperties.ContainsKey("Material Used"), dicProperties("Material Used"), "") ' dicProperties("Material Used")
        mtcReviewObj.matlspec = If(dicProperties.ContainsKey("MATL SPEC"), dicProperties("MATL SPEC"), "") ' dicProperties("MATL SPEC")
        mtcReviewObj.isBaseline = False
        mtcReviewObj.fullpath = objSheetMetalDocument.FullName
        mtcReviewObj.density = If(dicProperties.ContainsKey("Category"), dicProperties("Category"), "") ' dicProperties("Category")
        mtcReviewObj.projectname = If(dicProperties.ContainsKey("Project"), dicProperties("Project"), "") 'dicProperties("Project")
        If mtcReviewObj.projectname = "BROOKVILLE EQUIPMENT CORP" Then
            mtcReviewObj.isBrookVilleProject_Baseline = "Yes"
        End If
        mtcReviewObj.Documenttype = If(dicProperties.ContainsKey("Status Text"), dicProperties("Status Text"), "") ' dicProperties("Status Text")
        mtcReviewObj.UomProperty = If(dicProperties.ContainsKey("UOM"), dicProperties("UOM"), "") ' dicProperties("UOM")
        mtcReviewObj.keywords = If(dicProperties.ContainsKey("Keywords"), dicProperties("Keywords"), "") ' dicProperties("Keywords")
        mtcReviewObj.ECO = If(dicProperties.ContainsKey("ECO/SOW"), dicProperties("ECO/SOW"), "") ' dicProperties("ECO/SOW")
        mtcReviewObj.projectname = If(dicProperties.ContainsKey("Project Name"), dicProperties("Project Name"), "") ' dicProperties("ECO/SOW")
        mtcReviewObj.lastAuthor = If(dicProperties.ContainsKey("Last Author"), dicProperties("Last Author"), "")

        '13th Sep 2024

        If mtcReviewObj.fileName.Contains("_") Or mtcReviewObj.fileName.Contains("-") Then
            mtcReviewObj.revisionNumber_FileName = GetRevisionLevel(mtcReviewObj.fileName, mtcReviewObj.revisionNumber_Prop)
        Else
            mtcReviewObj.fileName = mtcReviewObj.fileNameWithoutExt
        End If


        'Project Name
        'mtcReviewObj.modifiedDate = If(dicProperties.ContainsKey("Modified"), dicProperties("Modified"), "") ' dicProperties("ECO/SOW")

        'We could not find the Modified property in assembly document
        'So, we have used last saved date.
        'Last Save Date
        Try
            mtcReviewObj.modifiedDate = If(dicProperties.ContainsKey("Last Save Date"), dicProperties("Last Save Date"), "") ' dicProperties("ECO/SOW")
            mtcReviewObj.modifiedDate = DateTime.Parse(mtcReviewObj.modifiedDate).ToShortDateString()
        Catch ex As Exception
            '9th Sep 2024 
            MsgBox(ex.Message)
        End Try

        'Dim documents As SolidEdgeFramework.Documents = objApp.Documents

        'Dim sheetmetaldoc As SolidEdgePart.SheetMetalDocument = objApp.ActiveDocument

        Try
            'Dim occurrences As SolidEdgeAssembly.Occurrences = partdoc.Occurrences
            'Dim partlst As New List(Of String)

            'For Each occur As SolidEdgeAssembly.Occurrence In occurrences
            '    If Not partlst.Contains(occur.Name) Then
            '        partlst.Add(occur.Name)
            '        mtcReviewObj.partlistcount += 1
            '    End If

            'Next



            GetPartListSheetMetalDoc()

            mtcReviewObj.partlistcount = GetBomCount()

            'mtcReviewObj.checkAssemblyFeature = CheckPartFeatureExistence(partdoc)

            'mtcReviewObj.interPartLink = CheckInterPartLinksPSM()

            'mtcReviewObj.isGeometryBroken = IsGeomtryBroken_assembly(partdoc)

            mtcReviewObj.qAQC = If(dicProperties.ContainsKey("QAQC"), dicProperties("QAQC"), "")

            mtcReviewObj.quantity = "1"

        Catch ex As Exception
            MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
            CustomLogUtil.Log("While Setting Main Assembly Data", ex.Message, ex.StackTrace)
            '9th Sep 2024 
            MsgBox(ex.Message)
        End Try

        Return mtcReviewObj
    End Function

    Public Function GetBomCount() As Integer

        Dim Count As Integer = dt.Rows.Count() - 1
        Dim value As String
        Dim BOMCount As Integer
        Dim Max As Integer = 0
        For i = 0 To Count
            value = dt.Rows(i)(0)
            If value.ToString.Contains("*") Then
                value = value.ToString.Replace("*", "")
                value.ToString.Trim()
            End If
            BOMCount = Convert.ToInt32(value)
            If BOMCount > Max Then
                Max = BOMCount
            End If
        Next

        BOMCount = Max
        Return BOMCount
    End Function

    Public Sub GetPartList()
        Dim objApp As SolidEdgeFramework.Application = Nothing
        Dim Objdocuemnts As SolidEdgeFramework.Documents = Nothing
        Dim ObjdraftDoc As SolidEdgeDraft.DraftDocument = Nothing
        Dim ObjAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = Nothing
        Dim ObjPartdoc As SolidEdgePart.PartDocument = Nothing
        Dim ObjSheetMetalDoc As SolidEdgePart.SheetMetalDocument = Nothing
        Dim Objsheet As SolidEdgeDraft.Sheet = Nothing
        Dim ObjdrwViews As SolidEdgeDraft.DrawingViews = Nothing
        Dim ObjdrwView As SolidEdgeDraft.DrawingView = Nothing
        Dim ObjpartLists As SolidEdgeDraft.PartsLists = Nothing
        Dim ObjpartList As SolidEdgeDraft.PartsList = Nothing
        Dim objModelLink As SolidEdgeDraft.ModelLink = Nothing
        Dim objModelLinks As SolidEdgeDraft.ModelLinks = Nothing
        Dim ObjTableCell As SolidEdgeDraft.TableCell = Nothing
        Dim ObjCols As SolidEdgeDraft.TableColumns = Nothing
        Dim ObjRows As SolidEdgeDraft.TableRows = Nothing

        Try
            objApp = Marshal.GetActiveObject("SolidEdge.Application")

            ObjAssemblyDoc = objApp.ActiveDocument
            Objdocuemnts = objApp.Documents
            ObjdraftDoc = Objdocuemnts.Add("SolidEdge.DraftDocument")
            Objsheet = ObjdraftDoc.ActiveSheet
            objModelLinks = ObjdraftDoc.ModelLinks
            Dim filename As String
            Dim file As String = objAssemblyDocument.FullName
            filename = System.IO.Path.Combine(SolidEdgeCommunity.SolidEdgeUtils.GetTrainingFolderPath(), file)
            objModelLink = objModelLinks.Add(filename)
            ObjdrwViews = Objsheet.DrawingViews
            ObjdrwView = ObjdrwViews.AddAssemblyView(From:=objModelLink, Orientation:=SolidEdgeDraft.ViewOrientationConstants.igDimetricTopBackLeftView, Scale:=1.0, x:=0.4, y:=0.4, ViewType:=SolidEdgeDraft.AssemblyDrawingViewTypeConstants.seAssemblyDesignedView)
            ObjdrwView.Caption = "New Drawing View"
            ObjdrwView.DisplayCaption = False
        Catch ex As Exception
            MessageBox.Show("Error While Getting Part Lists", ex.Message)
        End Try

        ObjdraftDoc = objApp.ActiveDocument
        ObjpartLists = ObjdraftDoc.PartsLists
        ObjpartList = ObjpartLists.Add(ObjdrwView, "BEC", 1, 1)
        ObjpartList = ObjpartLists.Item(1)
        ObjCols = ObjpartList.Columns
        ObjRows = ObjpartList.Rows


        For Each tableColumn In ObjCols.OfType(Of SolidEdgeDraft.TableColumn)()
            Dim dtcolums As DataColumn = New DataColumn()
            dtcolums.ColumnName = tableColumn.Header
            dt.Columns.Add(dtcolums.ColumnName)
            Debug.Print(tableColumn.Header)

        Next tableColumn

        Dim dtrows As DataRow = dt.NewRow()
        dt.Rows.Add(dtrows)

        For Each tableRow In ObjRows.OfType(Of SolidEdgeDraft.TableRow)()

            For Each tableColumn In ObjCols.OfType(Of SolidEdgeDraft.TableColumn)()

                If tableColumn.Show Then
                    ObjTableCell = ObjpartList.Cell(tableRow.Index, tableColumn.Index)
                    Dim rowindex As Integer = tableRow.Index - 1
                    Dim colindex As Integer = tableColumn.Index - 1
                    Dim tablevalue As String = ObjTableCell.value
                    dt.Rows(rowindex)(colindex) = tablevalue
                End If
            Next

            dtrows = dt.NewRow()
            dt.Rows.Add(dtrows)
        Next

        Dim rowCnt As Integer = dt.Rows.Count

        If dt.Rows.Count > 0 Then
            dt.Rows.RemoveAt(rowCnt - 1)
        End If

        objApp.Documents.CloseDocument(ObjdraftDoc.FullName, False, "", False, False)


    End Sub

    '17th Sep 2024
    Public Sub GetPartListPartDoc()

        Dim objApp As SolidEdgeFramework.Application = Nothing
        Dim Objdocuemnts As SolidEdgeFramework.Documents = Nothing
        Dim ObjdraftDoc As SolidEdgeDraft.DraftDocument = Nothing
        Dim ObjAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = Nothing
        Dim ObjPartdoc As SolidEdgePart.PartDocument = Nothing
        Dim ObjSheetMetalDoc As SolidEdgePart.SheetMetalDocument = Nothing
        Dim Objsheet As SolidEdgeDraft.Sheet = Nothing
        Dim ObjdrwViews As SolidEdgeDraft.DrawingViews = Nothing
        Dim ObjdrwView As SolidEdgeDraft.DrawingView = Nothing
        Dim ObjpartLists As SolidEdgeDraft.PartsLists = Nothing
        Dim ObjpartList As SolidEdgeDraft.PartsList = Nothing
        Dim objModelLink As SolidEdgeDraft.ModelLink = Nothing
        Dim objModelLinks As SolidEdgeDraft.ModelLinks = Nothing
        Dim ObjTableCell As SolidEdgeDraft.TableCell = Nothing
        Dim ObjCols As SolidEdgeDraft.TableColumns = Nothing
        Dim ObjRows As SolidEdgeDraft.TableRows = Nothing

        Try
            objApp = Marshal.GetActiveObject("SolidEdge.Application")

            ObjPartdoc = objApp.ActiveDocument
            Objdocuemnts = objApp.Documents
            ObjdraftDoc = Objdocuemnts.Add("SolidEdge.DraftDocument")
            Objsheet = ObjdraftDoc.ActiveSheet
            objModelLinks = ObjdraftDoc.ModelLinks
            Dim filename As String
            Dim file As String = objPartDocument.FullName
            filename = System.IO.Path.Combine(SolidEdgeCommunity.SolidEdgeUtils.GetTrainingFolderPath(), file)
            objModelLink = objModelLinks.Add(filename)
            ObjdrwViews = Objsheet.DrawingViews  'AddPartView 'AddAssemblyView
            ObjdrwView = ObjdrwViews.AddPartView(From:=objModelLink, Orientation:=SolidEdgeDraft.ViewOrientationConstants.igDimetricTopBackLeftView, Scale:=1.0, x:=0.4, y:=0.4, ViewType:=SolidEdgeDraft.AssemblyDrawingViewTypeConstants.seAssemblyDesignedView)
            ObjdrwView.Caption = "New Drawing View"
            ObjdrwView.DisplayCaption = False
        Catch ex As Exception
            MessageBox.Show("Error While Getting Part Lists", ex.Message)
        End Try

        ObjdraftDoc = objApp.ActiveDocument
        ObjpartLists = ObjdraftDoc.PartsLists
        ObjpartList = ObjpartLists.Add(ObjdrwView, "BEC", 1, 1)
        ObjpartList = ObjpartLists.Item(1)
        ObjCols = ObjpartList.Columns
        ObjRows = ObjpartList.Rows


        For Each tableColumn In ObjCols.OfType(Of SolidEdgeDraft.TableColumn)()
            Dim dtcolums As DataColumn = New DataColumn()
            dtcolums.ColumnName = tableColumn.Header
            dt.Columns.Add(dtcolums.ColumnName)
            Debug.Print(tableColumn.Header)

        Next tableColumn

        Dim dtrows As DataRow = dt.NewRow()
        dt.Rows.Add(dtrows)

        For Each tableRow In ObjRows.OfType(Of SolidEdgeDraft.TableRow)()

            For Each tableColumn In ObjCols.OfType(Of SolidEdgeDraft.TableColumn)()

                If tableColumn.Show Then
                    ObjTableCell = ObjpartList.Cell(tableRow.Index, tableColumn.Index)
                    Dim rowindex As Integer = tableRow.Index - 1
                    Dim colindex As Integer = tableColumn.Index - 1
                    Dim tablevalue As String = ObjTableCell.value
                    dt.Rows(rowindex)(colindex) = tablevalue
                End If
            Next

            dtrows = dt.NewRow()
            dt.Rows.Add(dtrows)
        Next

        Dim rowCnt As Integer = dt.Rows.Count

        If dt.Rows.Count > 0 Then
            dt.Rows.RemoveAt(rowCnt - 1)
        End If

        objApp.Documents.CloseDocument(ObjdraftDoc.FullName, False, "", False, False)


    End Sub


    '17th Sep 2024
    Public Sub GetPartListSheetMetalDoc()

        Dim objApp As SolidEdgeFramework.Application = Nothing
        Dim Objdocuemnts As SolidEdgeFramework.Documents = Nothing
        Dim ObjdraftDoc As SolidEdgeDraft.DraftDocument = Nothing
        Dim ObjAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = Nothing
        Dim ObjPartdoc As SolidEdgePart.PartDocument = Nothing
        Dim ObjSheetMetalDoc As SolidEdgePart.SheetMetalDocument = Nothing
        Dim Objsheet As SolidEdgeDraft.Sheet = Nothing
        Dim ObjdrwViews As SolidEdgeDraft.DrawingViews = Nothing
        Dim ObjdrwView As SolidEdgeDraft.DrawingView = Nothing
        Dim ObjpartLists As SolidEdgeDraft.PartsLists = Nothing
        Dim ObjpartList As SolidEdgeDraft.PartsList = Nothing
        Dim objModelLink As SolidEdgeDraft.ModelLink = Nothing
        Dim objModelLinks As SolidEdgeDraft.ModelLinks = Nothing
        Dim ObjTableCell As SolidEdgeDraft.TableCell = Nothing
        Dim ObjCols As SolidEdgeDraft.TableColumns = Nothing
        Dim ObjRows As SolidEdgeDraft.TableRows = Nothing

        Try
            objApp = Marshal.GetActiveObject("SolidEdge.Application")

            ObjSheetMetalDoc = objApp.ActiveDocument
            Objdocuemnts = objApp.Documents
            ObjdraftDoc = Objdocuemnts.Add("SolidEdge.DraftDocument")
            Objsheet = ObjdraftDoc.ActiveSheet
            objModelLinks = ObjdraftDoc.ModelLinks
            Dim filename As String
            Dim file As String = objSheetMetalDocument.FullName
            filename = System.IO.Path.Combine(SolidEdgeCommunity.SolidEdgeUtils.GetTrainingFolderPath(), file)
            objModelLink = objModelLinks.Add(filename)
            ObjdrwViews = Objsheet.DrawingViews 'AddSheetMetalView 'AddPartView 'AddAssemblyView
            ObjdrwView = ObjdrwViews.AddSheetMetalView(From:=objModelLink, Orientation:=SolidEdgeDraft.ViewOrientationConstants.igDimetricTopBackLeftView, Scale:=1.0, x:=0.4, y:=0.4, ViewType:=SolidEdgeDraft.AssemblyDrawingViewTypeConstants.seAssemblyDesignedView)
            ObjdrwView.Caption = "New Drawing View"
            ObjdrwView.DisplayCaption = False
        Catch ex As Exception
            MessageBox.Show("Error While Getting Part Lists", ex.Message)
        End Try

        ObjdraftDoc = objApp.ActiveDocument
        ObjpartLists = ObjdraftDoc.PartsLists
        ObjpartList = ObjpartLists.Add(ObjdrwView, "BEC", 1, 1)
        ObjpartList = ObjpartLists.Item(1)
        ObjCols = ObjpartList.Columns
        ObjRows = ObjpartList.Rows


        For Each tableColumn In ObjCols.OfType(Of SolidEdgeDraft.TableColumn)()
            Dim dtcolums As DataColumn = New DataColumn()
            dtcolums.ColumnName = tableColumn.Header
            dt.Columns.Add(dtcolums.ColumnName)
            Debug.Print(tableColumn.Header)

        Next tableColumn

        Dim dtrows As DataRow = dt.NewRow()
        dt.Rows.Add(dtrows)

        For Each tableRow In ObjRows.OfType(Of SolidEdgeDraft.TableRow)()

            For Each tableColumn In ObjCols.OfType(Of SolidEdgeDraft.TableColumn)()

                If tableColumn.Show Then
                    ObjTableCell = ObjpartList.Cell(tableRow.Index, tableColumn.Index)
                    Dim rowindex As Integer = tableRow.Index - 1
                    Dim colindex As Integer = tableColumn.Index - 1
                    Dim tablevalue As String = ObjTableCell.value
                    dt.Rows(rowindex)(colindex) = tablevalue
                End If
            Next

            dtrows = dt.NewRow()
            dt.Rows.Add(dtrows)
        Next

        Dim rowCnt As Integer = dt.Rows.Count

        If dt.Rows.Count > 0 Then
            dt.Rows.RemoveAt(rowCnt - 1)
        End If

        objApp.Documents.CloseDocument(ObjdraftDoc.FullName, False, "", False, False)


    End Sub


    Private Function GetMainAssemblyDoc(ByVal mtcMtrModelObj As MTC_MTR_Model, ByVal mainAssemblyDocModel As Object) As MTC_MTR_Model
        '9th Sep 2024 
        Try
            Dim dv As New DataView(mtcMtrModelObj.dtM2M)
            Dim partno As String = ExcelUtil.ExcelMtcReview.fpartno.ToString()
            'dv.RowFilter = $"{partno}='{mainAssemblyDocModel.fileNameWithoutExt}'"


            Dim filter As String = $"Convert([{partno}], 'System.String') = '{mainAssemblyDocModel.fileNameWithoutExt}'"
            dv.RowFilter = filter '$"{materialUsedCol}='{cmbMaterialUsed2_Mw.Text}'"

            Dim m2MDataObj As New M2MDataModel()
            m2MDataObj = SetM2mData(m2MDataObj, dv)

            If dv.Count = 0 Then
                mainAssemblyDocModel.ispartfound = False
            Else
                mainAssemblyDocModel.ispartfound = True
            End If

            Dim mtcAssemblyObj As MTC_Assembly = GetMTCAssembly(mainAssemblyDocModel, mtcMtrModelObj, m2MDataObj)
            Dim mtrAssemblyObj As MTR_Assembly = GetMTRAssembly(mainAssemblyDocModel, mtcMtrModelObj, m2MDataObj)

            Dim routingSeqAssemblyObj As RoutingSequence_Assembly = GetRoutingSequenceAssembly(mainAssemblyDocModel, mtcMtrModelObj, m2MDataObj)

            For i As Integer = 0 To mtcMtrModelObj.authorList.Count - 1
                mtcMtrModelObj.authorList(i) = mtcMtrModelObj.authorList(i).Trim().ToUpper
            Next
            If mtcMtrModelObj.authorList.Contains(mainAssemblyDocModel.author.ToString.ToUpper) Then

                mtcMtrModelObj.mtcAssemblyList_BEC.Add(mtcAssemblyObj)
                mtcMtrModelObj.mtrAssemblyList_BEC.Add(mtrAssemblyObj)

                mtcMtrModelObj.routingSequenceAssemblyList_BEC.Add(routingSeqAssemblyObj)


            Else
                'mtcAssemblyObj.author = "DGS"
                mtcMtrModelObj.mtcAssemblyList_DGS.Add(mtcAssemblyObj)
                mtcMtrModelObj.mtrAssemblyList_DGS.Add(mtrAssemblyObj)

                mtcMtrModelObj.routingSequenceAssemblyList_DGS.Add(routingSeqAssemblyObj)
            End If

            Return mtcMtrModelObj
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function


    '17th Sep 2024
    Private Function GetMainPartDoc(ByVal mtcMtrModelObj As MTC_MTR_Model, ByVal mainPartDocModel As Object) As MTC_MTR_Model
        '9th Sep 2024 
        Try
            Dim dv As New DataView(mtcMtrModelObj.dtM2M)
            Dim partno As String = ExcelUtil.ExcelMtcReview.fpartno.ToString()
            'dv.RowFilter = $"{partno}='{mainAssemblyDocModel.fileNameWithoutExt}'"


            Dim filter As String = $"Convert([{partno}], 'System.String') = '{mainPartDocModel.fileNameWithoutExt}'"
            dv.RowFilter = filter '$"{materialUsedCol}='{cmbMaterialUsed2_Mw.Text}'"

            Dim m2MDataObj As New M2MDataModel()
            m2MDataObj = SetM2mData(m2MDataObj, dv)

            If dv.Count = 0 Then
                mainPartDocModel.ispartfound = False
            Else
                mainPartDocModel.ispartfound = True
            End If

            Dim mtcAssemblyObj As MTC_Assembly = GetMTCAssembly(mainPartDocModel, mtcMtrModelObj, m2MDataObj)
            Dim mtrAssemblyObj As MTR_Assembly = GetMTRAssembly(mainPartDocModel, mtcMtrModelObj, m2MDataObj)

            Dim routingSeqAssemblyObj As RoutingSequence_Assembly = GetRoutingSequenceAssembly(mainPartDocModel, mtcMtrModelObj, m2MDataObj)

            For i As Integer = 0 To mtcMtrModelObj.authorList.Count - 1
                mtcMtrModelObj.authorList(i) = mtcMtrModelObj.authorList(i).Trim().ToUpper
            Next
            If mtcMtrModelObj.authorList.Contains(mainPartDocModel.author.ToString.ToUpper) Then

                mtcMtrModelObj.mtcAssemblyList_BEC.Add(mtcAssemblyObj)
                mtcMtrModelObj.mtrAssemblyList_BEC.Add(mtrAssemblyObj)

                mtcMtrModelObj.routingSequenceAssemblyList_BEC.Add(routingSeqAssemblyObj)


            Else
                'mtcAssemblyObj.author = "DGS"
                mtcMtrModelObj.mtcAssemblyList_DGS.Add(mtcAssemblyObj)
                mtcMtrModelObj.mtrAssemblyList_DGS.Add(mtrAssemblyObj)

                mtcMtrModelObj.routingSequenceAssemblyList_DGS.Add(routingSeqAssemblyObj)
            End If

            Return mtcMtrModelObj
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function


    '17th Sep 2024
    Private Function GetMainSheetMetalDoc(ByVal mtcMtrModelObj As MTC_MTR_Model, ByVal mainSheetMetalDocModel As Object) As MTC_MTR_Model
        '9th Sep 2024 
        Try
            Dim dv As New DataView(mtcMtrModelObj.dtM2M)
            Dim partno As String = ExcelUtil.ExcelMtcReview.fpartno.ToString()
            'dv.RowFilter = $"{partno}='{mainAssemblyDocModel.fileNameWithoutExt}'"


            Dim filter As String = $"Convert([{partno}], 'System.String') = '{mainSheetMetalDocModel.fileNameWithoutExt}'"
            dv.RowFilter = filter '$"{materialUsedCol}='{cmbMaterialUsed2_Mw.Text}'"

            Dim m2MDataObj As New M2MDataModel()
            m2MDataObj = SetM2mData(m2MDataObj, dv)

            If dv.Count = 0 Then
                mainSheetMetalDocModel.ispartfound = False
            Else
                mainSheetMetalDocModel.ispartfound = True
            End If

            Dim mtcAssemblyObj As MTC_Assembly = GetMTCAssembly(mainSheetMetalDocModel, mtcMtrModelObj, m2MDataObj)
            Dim mtrAssemblyObj As MTR_Assembly = GetMTRAssembly(mainSheetMetalDocModel, mtcMtrModelObj, m2MDataObj)

            Dim routingSeqAssemblyObj As RoutingSequence_Assembly = GetRoutingSequenceAssembly(mainSheetMetalDocModel, mtcMtrModelObj, m2MDataObj)

            For i As Integer = 0 To mtcMtrModelObj.authorList.Count - 1
                mtcMtrModelObj.authorList(i) = mtcMtrModelObj.authorList(i).Trim().ToUpper
            Next
            If mtcMtrModelObj.authorList.Contains(mainSheetMetalDocModel.author.ToString.ToUpper) Then

                mtcMtrModelObj.mtcAssemblyList_BEC.Add(mtcAssemblyObj)
                mtcMtrModelObj.mtrAssemblyList_BEC.Add(mtrAssemblyObj)

                mtcMtrModelObj.routingSequenceAssemblyList_BEC.Add(routingSeqAssemblyObj)


            Else
                'mtcAssemblyObj.author = "DGS"
                mtcMtrModelObj.mtcAssemblyList_DGS.Add(mtcAssemblyObj)
                mtcMtrModelObj.mtrAssemblyList_DGS.Add(mtrAssemblyObj)

                mtcMtrModelObj.routingSequenceAssemblyList_DGS.Add(routingSeqAssemblyObj)
            End If

            Return mtcMtrModelObj
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

    Private Sub BtnExportMTC_MTR_Report_Click(sender As Object, e As EventArgs) Handles BtnExportMTC_MTR_RoutingSeq_Report.Click

        WaitStartSave()
        Try

            Dim mtcMtrBL As New MTC_MTR_BL()

            mtcMtrModelObj = New MTC_MTR_Model()

            '17th Sep 2024
            activeDocument = objApp.ActiveDocument

            '17th Sep 2024
            If activeDocument.Application.ActiveEnvironment.ToLower = "assembly" Then

                objAssemblyDocument = activeDocument

                mtcMtrModelObj.assemblyPath = objAssemblyDocument.FullName

                log.Info("============")
                log.Info($"Report generation of {mtcMtrModelObj.assemblyPath} started.")
                log.Info("============")
                CustomLogUtil.Heading($"Report generation of {mtcMtrModelObj.assemblyPath} started.")

                Dim aPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location

                Dim dirPath As String = IO.Path.GetDirectoryName(aPath)

                waitFormObj.SetProgressInformationMessage("Read M2M data")

                mtcMtrModelObj.dtM2M = GetM2MData(dirPath)

                log.Info("Read M2MData completed.")
                CustomLogUtil.Log("Read M2MData completed.")

                waitFormObj.SetProgressInformationMessage("Read propseed file")

                'Dim propseedfile As String = IO.Path.Combine(dirPath, "propseed.txt")
                Dim propseedfile As String = Config.configObj.propseedFile

                mtcMtrModelObj.projectNameList = Readfile(propseedfile)

                log.Info("Read propseed details completed.")
                CustomLogUtil.Log("Read propseed details completed.")

                waitFormObj.SetProgressInformationMessage("Read author file")

                mtcMtrModelObj.authorList = ReadfileAuthor(propseedfile)

                log.Info("Read author details completed.")
                CustomLogUtil.Log("Read author details completed.")

                waitFormObj.SetProgressInformationMessage("Read active assembly properties")
                'mtcMtrModelObj.dtAuthorData = mtcMtrBL.ReadAuthorExcelData(mtcMtrModelObj.dtAuthorData)

                Dim dicMainAssemblyProperties As Dictionary(Of String, String) = GetMainAssemblyData(mtcMtrModelObj)

                log.Info("Read active assembly properties completed.")
                CustomLogUtil.Log("Read active assembly properties completed.")

                Dim mainAssemblyDocModel As DocumentModel = SetMainAssemblyData(mtcMtrModelObj, dicMainAssemblyProperties)

                '4th Sep 2024
                Dim MTC_Author_Name As String = System.Environment.UserName 'mainAssemblyDocModel.author
                '28th Oct 2024
                'Dim MTC_Author_Name As String = System.Security.Principal.WindowsIdentity.GetCurrent().Name 'Environment.UserName

                mtcMtrModelObj = GetMainAssemblyDoc(mtcMtrModelObj, mainAssemblyDocModel)

                log.Info("Create main assembly document data completed.")
                CustomLogUtil.Log("Create main assembly document data completed.")

                waitFormObj.SetProgressInformationMessage("Read current assembly part list data")

                mtcMtrModelObj.dtCurrentAssemblyData = mtcMtrBL.GetCurrentAssemblyData()

                'mtcMtrModelObj.BOMCount = mtcMtrBL.GetBOMCount(mtcMtrModelObj.dtCurrentAssemblyData)

                'MainAsmBomCount = mtcMtrModelObj.BOMCount
                log.Info("Read partlist data completed.")
                CustomLogUtil.Log("Read partlist data completed.")

                dgvDocumentDetails.RowHeadersVisible = False

                dgvDocumentDetails.DataSource = mtcMtrModelObj.dtCurrentAssemblyData

                waitFormObj.SetProgressInformationMessage("Filter current assembly part list data")

                If chkAll.Checked = False Then

                    mtcMtrModelObj.dtFilteredAssemblyData = FilterCurrentAssemblyData(mtcMtrModelObj)
                Else

                    mtcMtrModelObj.dtFilteredAssemblyData = mtcMtrModelObj.dtCurrentAssemblyData.Copy()

                End If

                log.Info("Filter assembly data completed.")
                CustomLogUtil.Log("Filter assembly data completed.")

                dgvDocumentDetails.DataSource = mtcMtrModelObj.dtFilteredAssemblyData

                mtcMtrModelObj.dtBECAuthorAssemblyData = mtcMtrBL.GetBECAuthorAssemblyData(mtcMtrModelObj)

                mtcMtrModelObj.dtNonBECAuthorAssemblyData = mtcMtrBL.GetNonBECAuthorAssemblyData(mtcMtrModelObj)

                mtcMtrModelObj.exportDirectoryLocation = txtExportDirLocationMTR.Text

                '2nd Sep 2024
                mtcMtrModelObj.export_MTR_Report_DirectoryLocation = txtExportDirLocationMTR.Text
                mtcMtrModelObj.export_Routing_Report_DirectoryLocation = txtExportDirLocationRouting.Text

                mtcMtrModelObj.baseLineDirectoryLocation = txtBaseLineDirectoryPath.Text

                objAssemblyDocument.Close(False)

                waitFormObj.SetProgressInformationMessage("Create assembly report data")

                mtcMtrModelObj = GetMTC_MTR_ReportData(mtcMtrModelObj)

                log.Info("get report data collection completed.")
                CustomLogUtil.Log("get report data collection completed.")


                '-------------------------------------------------------------------------

                mtcMtrModelObj.exportDirectoryLocation = IO.Path.Combine(mtcMtrModelObj.exportDirectoryLocation, IO.Path.Combine(IO.Path.GetFileNameWithoutExtension(mtcMtrModelObj.assemblyPath)))

                '2nd Sep 2024
                mtcMtrModelObj.export_MTR_Report_DirectoryLocation = IO.Path.Combine(mtcMtrModelObj.export_MTR_Report_DirectoryLocation, IO.Path.Combine(IO.Path.GetFileNameWithoutExtension(mtcMtrModelObj.assemblyPath)))
                mtcMtrModelObj.export_Routing_Report_DirectoryLocation = IO.Path.Combine(mtcMtrModelObj.export_Routing_Report_DirectoryLocation, IO.Path.Combine(IO.Path.GetFileNameWithoutExtension(mtcMtrModelObj.assemblyPath)))

                '2nd Sep 2024
                'If Not IO.Directory.Exists(mtcMtrModelObj.exportDirectoryLocation) Then
                '    IO.Directory.CreateDirectory(mtcMtrModelObj.exportDirectoryLocation)
                'End If

                '2nd Sep 2024
                If Not IO.Directory.Exists(mtcMtrModelObj.export_MTR_Report_DirectoryLocation) Then
                    IO.Directory.CreateDirectory(mtcMtrModelObj.export_MTR_Report_DirectoryLocation)
                End If

                '26th Oct 2024
                ''2nd Sep 2024
                'If Not IO.Directory.Exists(mtcMtrModelObj.export_Routing_Report_DirectoryLocation) Then
                '    IO.Directory.CreateDirectory(mtcMtrModelObj.export_Routing_Report_DirectoryLocation)
                'End If

                waitFormObj.SetProgressInformationMessage("Create excel report")

                'temp 20Feb2024
                Dim mtcMtrExcelObj As New MTC_MTR_ExcelUtil()
                mtcMtrExcelObj.SaveAsMTCExcel(mtcMtrModelObj, MTC_Author_Name)

                log.Info("MTC export completed")
                CustomLogUtil.Log("MTC export completed")



                ''temp 20Feb2024
                'mtcMtrExcelObj.SaveAsMTCExcel1(mtcMtrModelObj, "BEC", ExcelFileName)

                'mtcMtrExcelObj.SaveAsMTCExcel(mtcMtrModelObj, "DGS")

                'log.Info("DGS MTC export completed")
                'CustomLogUtil.Log("DGS MTC export completed")

                ''temp 19Feb2024------------
#Region "MTR Comment"
                'mtcMtrExcelObj.SaveAsMTRExcel(mtcMtrModelObj, "BEC")

                'log.Info("BEC MTR export completed")
                'CustomLogUtil.Log("BEC MTR export completed")

                'mtcMtrExcelObj.SaveAsMTRExcel(mtcMtrModelObj, "DGS")

                'log.Info("DGS MTR export completed")
                'CustomLogUtil.Log("DGS MTR export completed")

#End Region
                '26th Oct 2024
                'mtcMtrExcelObj.SaveAsRoutingSequenceExcel(mtcMtrModelObj)

                objAssemblyDocument = DirectCast(objApp.Documents.Open(mtcMtrModelObj.assemblyPath), SolidEdgeFramework.SolidEdgeDocument)

                objApp.DisplayAlerts = True

                SolidEdgeCommunity.OleMessageFilter.Unregister()

                'Me.Close()

                log.Info($"Report generation completed.")
                CustomLogUtil.Heading($"Report generation completed.")

                '4th Sep 2024
                waitFormObj.SetProgressInformationMessage("Excel report Created!")

                '2nd Sep 2024
                'If IO.Directory.Exists(mtcMtrModelObj.exportDirectoryLocation) Then
                '    Process.Start(mtcMtrModelObj.exportDirectoryLocation)
                'End If

                '2nd Sep 2024
                If IO.Directory.Exists(mtcMtrModelObj.export_MTR_Report_DirectoryLocation) Then
                    Process.Start(mtcMtrModelObj.export_MTR_Report_DirectoryLocation)
                End If

                '26th Oct 2024
                ''2nd Sep 2024
                'If IO.Directory.Exists(mtcMtrModelObj.export_Routing_Report_DirectoryLocation) Then
                '    Process.Start(mtcMtrModelObj.export_Routing_Report_DirectoryLocation)
                'End If



                '17th Sep 2024
            ElseIf activeDocument.Application.ActiveEnvironment.ToLower = "part" Then

                objPartDocument = activeDocument

                mtcMtrModelObj.partPath = objPartDocument.FullName

                log.Info("============")
                log.Info($"Report generation of {mtcMtrModelObj.partPath} started.")
                log.Info("============")
                CustomLogUtil.Heading($"Report generation of {mtcMtrModelObj.partPath} started.")

                Dim aPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location

                Dim dirPath As String = IO.Path.GetDirectoryName(aPath)

                waitFormObj.SetProgressInformationMessage("Read M2M data")

                mtcMtrModelObj.dtM2M = GetM2MData(dirPath)

                log.Info("Read M2MData completed.")
                CustomLogUtil.Log("Read M2MData completed.")

                waitFormObj.SetProgressInformationMessage("Read propseed file")

                Dim propseedfile As String = Config.configObj.propseedFile

                mtcMtrModelObj.projectNameList = Readfile(propseedfile)

                log.Info("Read propseed details completed.")
                CustomLogUtil.Log("Read propseed details completed.")

                waitFormObj.SetProgressInformationMessage("Read author file")

                mtcMtrModelObj.authorList = ReadfileAuthor(propseedfile)

                log.Info("Read author details completed.")
                CustomLogUtil.Log("Read author details completed.")

                waitFormObj.SetProgressInformationMessage("Read active part properties")

                Dim dicMainPartProperties As Dictionary(Of String, String) = GetPartPropertyData(mtcMtrModelObj)

                log.Info("Read active part properties completed.")
                CustomLogUtil.Log("Read active part properties completed.")

                Dim mainPartDocModel As DocumentModel = SetPartPropertyData(mtcMtrModelObj, dicMainPartProperties)

                '13th Nov 2024
                'Dim MTC_Author_Name As String = mainPartDocModel.author
                Dim MTC_Author_Name As String = System.Environment.UserName

                mtcMtrModelObj = GetMainPartDoc(mtcMtrModelObj, mainPartDocModel)

                log.Info("Create part document data completed.")
                CustomLogUtil.Log("Create part document data completed.")

                waitFormObj.SetProgressInformationMessage("Read current part list data")

                mtcMtrModelObj.dtCurrentPartData = mtcMtrBL.GetCurrentPartData()

                log.Info("Read partlist data completed.")
                CustomLogUtil.Log("Read partlist data completed.")

                dgvDocumentDetails.RowHeadersVisible = False

                dgvDocumentDetails.DataSource = mtcMtrModelObj.dtCurrentPartData

                waitFormObj.SetProgressInformationMessage("Filter current part list data")

                If chkAll.Checked = False Then

                    mtcMtrModelObj.dtFilteredAssemblyData = FilterCurrentPartData(mtcMtrModelObj)
                Else

                    mtcMtrModelObj.dtFilteredAssemblyData = mtcMtrModelObj.dtCurrentPartData.Copy()

                End If

                log.Info("Filter part data completed.")
                CustomLogUtil.Log("Filter part data completed.")

                dgvDocumentDetails.DataSource = mtcMtrModelObj.dtFilteredAssemblyData

                mtcMtrModelObj.dtBECAuthorAssemblyData = mtcMtrBL.GetBECAuthorAssemblyData(mtcMtrModelObj)

                mtcMtrModelObj.dtNonBECAuthorAssemblyData = mtcMtrBL.GetNonBECAuthorAssemblyData(mtcMtrModelObj)

                mtcMtrModelObj.exportDirectoryLocation = txtExportDirLocationMTR.Text

                'objAssemblyDocument.Close(False)

                waitFormObj.SetProgressInformationMessage("Create part report data")

                mtcMtrModelObj = GetMTC_MTR_ReportDataPartDoc(mtcMtrModelObj)

                log.Info("get report data collection completed.")
                CustomLogUtil.Log("get report data collection completed.")

                '13th Nov 2024
                'mtcMtrModelObj.exportDirectoryLocation = IO.Path.Combine(mtcMtrModelObj.exportDirectoryLocation, "Part Files")

                mtcMtrModelObj.exportDirectoryLocation = IO.Path.Combine(mtcMtrModelObj.exportDirectoryLocation, IO.Path.Combine(IO.Path.GetFileNameWithoutExtension(mtcMtrModelObj.partPath)))

                If Not IO.Directory.Exists(mtcMtrModelObj.exportDirectoryLocation) Then
                    IO.Directory.CreateDirectory(mtcMtrModelObj.exportDirectoryLocation)
                End If

                waitFormObj.SetProgressInformationMessage("Create excel report")

                Dim mtcMtrExcelObj As New MTC_MTR_ExcelUtil()
                mtcMtrExcelObj.SaveAsMTCExcelPartDoc(mtcMtrModelObj, MTC_Author_Name)

                log.Info("MTC export completed")
                CustomLogUtil.Log("MTC export completed")

                waitFormObj.SetProgressInformationMessage("Excel report Created!")

                objPartDocument = DirectCast(objApp.Documents.Open(mtcMtrModelObj.partPath), SolidEdgeFramework.SolidEdgeDocument)

                objApp.DisplayAlerts = True

                SolidEdgeCommunity.OleMessageFilter.Unregister()

                log.Info($"Report generation completed.")
                CustomLogUtil.Heading($"Report generation completed.")

                If IO.Directory.Exists(mtcMtrModelObj.exportDirectoryLocation) Then
                    Process.Start(mtcMtrModelObj.exportDirectoryLocation)
                End If



                '17th Sep 2024
            ElseIf activeDocument.Application.ActiveEnvironment.ToLower = "sheetmetal" Then

                objSheetMetalDocument = activeDocument

                mtcMtrModelObj.sheetMetalPath = objSheetMetalDocument.FullName

                log.Info("============")
                log.Info($"Report generation of {mtcMtrModelObj.partPath} started.")
                log.Info("============")
                CustomLogUtil.Heading($"Report generation of {mtcMtrModelObj.partPath} started.")

                Dim aPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location

                Dim dirPath As String = IO.Path.GetDirectoryName(aPath)

                waitFormObj.SetProgressInformationMessage("Read M2M data")

                mtcMtrModelObj.dtM2M = GetM2MData(dirPath)

                log.Info("Read M2MData completed.")
                CustomLogUtil.Log("Read M2MData completed.")

                waitFormObj.SetProgressInformationMessage("Read propseed file")

                Dim propseedfile As String = Config.configObj.propseedFile

                mtcMtrModelObj.projectNameList = Readfile(propseedfile)

                log.Info("Read propseed details completed.")
                CustomLogUtil.Log("Read propseed details completed.")

                waitFormObj.SetProgressInformationMessage("Read author file")

                mtcMtrModelObj.authorList = ReadfileAuthor(propseedfile)

                log.Info("Read author details completed.")
                CustomLogUtil.Log("Read author details completed.")

                waitFormObj.SetProgressInformationMessage("Read active SheetMetal properties")

                Dim dicMainSheetMetalProperties As Dictionary(Of String, String) = GetSheetMetalPropertyData(mtcMtrModelObj)

                log.Info("Read active SheetMetal properties completed.")
                CustomLogUtil.Log("Read active SheetMetal properties completed.")

                Dim mainSheetMetalDocModel As DocumentModel = SetSheetMetalPropertyData(mtcMtrModelObj, dicMainSheetMetalProperties)

                '13th Nov 2024
                'Dim MTC_Author_Name As String = mainSheetMetalDocModel.author
                Dim MTC_Author_Name As String = System.Environment.UserName

                mtcMtrModelObj = GetMainSheetMetalDoc(mtcMtrModelObj, mainSheetMetalDocModel)

                log.Info("Create SheetMetal document data completed.")
                CustomLogUtil.Log("Create SheetMetal document data completed.")

                waitFormObj.SetProgressInformationMessage("Read current partlist data")

                mtcMtrModelObj.dtCurrentSheetMetalData = mtcMtrBL.GetCurrentSheetMetalData()  'dtCurrentSheetMetalData  'dtCurrentPartData

                log.Info("Read partlist data completed.")
                CustomLogUtil.Log("Read partlist data completed.")

                dgvDocumentDetails.RowHeadersVisible = False

                dgvDocumentDetails.DataSource = mtcMtrModelObj.dtCurrentSheetMetalData

                waitFormObj.SetProgressInformationMessage("Filter current partlist data")

                If chkAll.Checked = False Then

                    mtcMtrModelObj.dtFilteredAssemblyData = FilterCurrentSheetMetalData(mtcMtrModelObj)
                Else

                    mtcMtrModelObj.dtFilteredAssemblyData = mtcMtrModelObj.dtCurrentSheetMetalData.Copy()

                End If

                log.Info("Filter SheetMetal data completed.")
                CustomLogUtil.Log("Filter SheetMetal data completed.")

                dgvDocumentDetails.DataSource = mtcMtrModelObj.dtFilteredAssemblyData

                mtcMtrModelObj.dtBECAuthorAssemblyData = mtcMtrBL.GetBECAuthorAssemblyData(mtcMtrModelObj)

                mtcMtrModelObj.dtNonBECAuthorAssemblyData = mtcMtrBL.GetNonBECAuthorAssemblyData(mtcMtrModelObj)

                mtcMtrModelObj.exportDirectoryLocation = txtExportDirLocationMTR.Text

                'objAssemblyDocument.Close(False)

                waitFormObj.SetProgressInformationMessage("Create SheetMetal report data")

                mtcMtrModelObj = GetMTC_MTR_ReportDataSheetMetalDoc(mtcMtrModelObj)

                log.Info("get report data collection completed.")
                CustomLogUtil.Log("get report data collection completed.")

                '13th Nov 2024
                'mtcMtrModelObj.exportDirectoryLocation = IO.Path.Combine(mtcMtrModelObj.exportDirectoryLocation, "SheetMetal Files")

                mtcMtrModelObj.exportDirectoryLocation = IO.Path.Combine(mtcMtrModelObj.exportDirectoryLocation, IO.Path.Combine(IO.Path.GetFileNameWithoutExtension(mtcMtrModelObj.sheetMetalPath)))

                If Not IO.Directory.Exists(mtcMtrModelObj.exportDirectoryLocation) Then
                    IO.Directory.CreateDirectory(mtcMtrModelObj.exportDirectoryLocation)
                End If

                waitFormObj.SetProgressInformationMessage("Create excel report")

                Dim mtcMtrExcelObj As New MTC_MTR_ExcelUtil()
                mtcMtrExcelObj.SaveAsMTCExcelSheetMetalDoc(mtcMtrModelObj, MTC_Author_Name)

                log.Info("MTC export completed")
                CustomLogUtil.Log("MTC export completed")

                waitFormObj.SetProgressInformationMessage("Excel report Created!")

                objSheetMetalDocument = DirectCast(objApp.Documents.Open(mtcMtrModelObj.sheetMetalPath), SolidEdgeFramework.SolidEdgeDocument)

                objApp.DisplayAlerts = True

                SolidEdgeCommunity.OleMessageFilter.Unregister()

                log.Info($"Report generation completed.")
                CustomLogUtil.Heading($"Report generation completed.")

                If IO.Directory.Exists(mtcMtrModelObj.exportDirectoryLocation) Then
                    Process.Start(mtcMtrModelObj.exportDirectoryLocation)
                End If

            End If

                WaitEndSave()
        Catch ex As Exception
            SolidEdgeCommunity.OleMessageFilter.Register()
            CustomLogUtil.Log("While Generating MTC Report", ex.Message, ex.StackTrace)
        End Try


#Region "Old Code 17th Sep 2024"

        '        WaitStartSave()
        '        Try

        '            Dim mtcMtrBL As New MTC_MTR_BL()

        '            mtcMtrModelObj = New MTC_MTR_Model()

        '            objAssemblyDocument = objApp.ActiveDocument

        '            mtcMtrModelObj.assemblyPath = objAssemblyDocument.FullName

        '            log.Info("============")
        '            log.Info($"Report generation of {mtcMtrModelObj.assemblyPath} started.")
        '            log.Info("============")
        '            CustomLogUtil.Heading($"Report generation of {mtcMtrModelObj.assemblyPath} started.")

        '            Dim aPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location

        '            Dim dirPath As String = IO.Path.GetDirectoryName(aPath)

        '            waitFormObj.SetProgressInformationMessage("Read M2M data")

        '            mtcMtrModelObj.dtM2M = GetM2MData(dirPath)

        '            log.Info("Read M2MData completed.")
        '            CustomLogUtil.Log("Read M2MData completed.")

        '            waitFormObj.SetProgressInformationMessage("Read propseed file")

        '            'Dim propseedfile As String = IO.Path.Combine(dirPath, "propseed.txt")
        '            Dim propseedfile As String = Config.configObj.propseedFile

        '            mtcMtrModelObj.projectNameList = Readfile(propseedfile)

        '            log.Info("Read propseed details completed.")
        '            CustomLogUtil.Log("Read propseed details completed.")

        '            waitFormObj.SetProgressInformationMessage("Read author file")

        '            mtcMtrModelObj.authorList = ReadfileAuthor(propseedfile)

        '            log.Info("Read author details completed.")
        '            CustomLogUtil.Log("Read author details completed.")

        '            waitFormObj.SetProgressInformationMessage("Read active assembly properties")
        '            'mtcMtrModelObj.dtAuthorData = mtcMtrBL.ReadAuthorExcelData(mtcMtrModelObj.dtAuthorData)

        '            Dim dicMainAssemblyProperties As Dictionary(Of String, String) = GetMainAssemblyData(mtcMtrModelObj)

        '            log.Info("Read active assembly properties completed.")
        '            CustomLogUtil.Log("Read active assembly properties completed.")

        '            Dim mainAssemblyDocModel As DocumentModel = SetMainAssemblyData(mtcMtrModelObj, dicMainAssemblyProperties)

        '            '4th Sep 2024
        '            Dim MTC_Author_Name As String = mainAssemblyDocModel.author

        '            mtcMtrModelObj = GetMainAssemblyDoc(mtcMtrModelObj, mainAssemblyDocModel)

        '            log.Info("Create main assembly document data completed.")
        '            CustomLogUtil.Log("Create main assembly document data completed.")

        '            waitFormObj.SetProgressInformationMessage("Read current assembly part list data")

        '            mtcMtrModelObj.dtCurrentAssemblyData = mtcMtrBL.GetCurrentAssemblyData()

        '            'mtcMtrModelObj.BOMCount = mtcMtrBL.GetBOMCount(mtcMtrModelObj.dtCurrentAssemblyData)

        '            'MainAsmBomCount = mtcMtrModelObj.BOMCount
        '            log.Info("Read partlist data completed.")
        '            CustomLogUtil.Log("Read partlist data completed.")

        '            dgvDocumentDetails.RowHeadersVisible = False

        '            dgvDocumentDetails.DataSource = mtcMtrModelObj.dtCurrentAssemblyData

        '            waitFormObj.SetProgressInformationMessage("Filter current assembly part list data")

        '            If chkAll.Checked = False Then

        '                mtcMtrModelObj.dtFilteredAssemblyData = FilterCurrentAssemblyData(mtcMtrModelObj)
        '            Else

        '                mtcMtrModelObj.dtFilteredAssemblyData = mtcMtrModelObj.dtCurrentAssemblyData.Copy()

        '            End If

        '            log.Info("Filter assembly data completed.")
        '            CustomLogUtil.Log("Filter assembly data completed.")

        '            dgvDocumentDetails.DataSource = mtcMtrModelObj.dtFilteredAssemblyData

        '            mtcMtrModelObj.dtBECAuthorAssemblyData = mtcMtrBL.GetBECAuthorAssemblyData(mtcMtrModelObj)

        '            mtcMtrModelObj.dtNonBECAuthorAssemblyData = mtcMtrBL.GetNonBECAuthorAssemblyData(mtcMtrModelObj)

        '            mtcMtrModelObj.exportDirectoryLocation = txtExportDirLocationMTR.Text

        '            '2nd Sep 2024
        '            mtcMtrModelObj.export_MTR_Report_DirectoryLocation = txtExportDirLocationMTR.Text
        '            mtcMtrModelObj.export_Routing_Report_DirectoryLocation = txtExportDirLocationRouting.Text

        '            mtcMtrModelObj.baseLineDirectoryLocation = txtBaseLineDirectoryPath.Text

        '            objAssemblyDocument.Close(False)

        '            waitFormObj.SetProgressInformationMessage("Create assembly report data")

        '            mtcMtrModelObj = GetMTC_MTR_ReportData(mtcMtrModelObj)

        '            log.Info("get report data collection completed.")
        '            CustomLogUtil.Log("get report data collection completed.")

        '            mtcMtrModelObj.exportDirectoryLocation = IO.Path.Combine(mtcMtrModelObj.exportDirectoryLocation, IO.Path.Combine(IO.Path.GetFileNameWithoutExtension(mtcMtrModelObj.assemblyPath)))

        '            '2nd Sep 2024
        '            mtcMtrModelObj.export_MTR_Report_DirectoryLocation = IO.Path.Combine(mtcMtrModelObj.export_MTR_Report_DirectoryLocation, IO.Path.Combine(IO.Path.GetFileNameWithoutExtension(mtcMtrModelObj.assemblyPath)))
        '            mtcMtrModelObj.export_Routing_Report_DirectoryLocation = IO.Path.Combine(mtcMtrModelObj.export_Routing_Report_DirectoryLocation, IO.Path.Combine(IO.Path.GetFileNameWithoutExtension(mtcMtrModelObj.assemblyPath)))

        '            '2nd Sep 2024
        '            'If Not IO.Directory.Exists(mtcMtrModelObj.exportDirectoryLocation) Then
        '            '    IO.Directory.CreateDirectory(mtcMtrModelObj.exportDirectoryLocation)
        '            'End If

        '            '2nd Sep 2024
        '            If Not IO.Directory.Exists(mtcMtrModelObj.export_MTR_Report_DirectoryLocation) Then
        '                IO.Directory.CreateDirectory(mtcMtrModelObj.export_MTR_Report_DirectoryLocation)
        '            End If

        '            '2nd Sep 2024
        '            If Not IO.Directory.Exists(mtcMtrModelObj.export_Routing_Report_DirectoryLocation) Then
        '                IO.Directory.CreateDirectory(mtcMtrModelObj.export_Routing_Report_DirectoryLocation)
        '            End If

        '            waitFormObj.SetProgressInformationMessage("Create excel report")

        '            'temp 20Feb2024
        '            Dim mtcMtrExcelObj As New MTC_MTR_ExcelUtil()
        '            mtcMtrExcelObj.SaveAsMTCExcel(mtcMtrModelObj, MTC_Author_Name)

        '            log.Info("MTC export completed")
        '            CustomLogUtil.Log("MTC export completed")

        '            ''temp 20Feb2024
        '            'mtcMtrExcelObj.SaveAsMTCExcel1(mtcMtrModelObj, "BEC", ExcelFileName)

        '            'mtcMtrExcelObj.SaveAsMTCExcel(mtcMtrModelObj, "DGS")

        '            'log.Info("DGS MTC export completed")
        '            'CustomLogUtil.Log("DGS MTC export completed")

        '            ''temp 19Feb2024------------
        '#Region "MTR Comment"
        '            'mtcMtrExcelObj.SaveAsMTRExcel(mtcMtrModelObj, "BEC")

        '            'log.Info("BEC MTR export completed")
        '            'CustomLogUtil.Log("BEC MTR export completed")

        '            'mtcMtrExcelObj.SaveAsMTRExcel(mtcMtrModelObj, "DGS")

        '            'log.Info("DGS MTR export completed")
        '            'CustomLogUtil.Log("DGS MTR export completed")

        '#End Region

        '            mtcMtrExcelObj.SaveAsRoutingSequenceExcel(mtcMtrModelObj)

        '            objAssemblyDocument = DirectCast(objApp.Documents.Open(mtcMtrModelObj.assemblyPath), SolidEdgeFramework.SolidEdgeDocument)

        '            objApp.DisplayAlerts = True



        '            SolidEdgeCommunity.OleMessageFilter.Unregister()

        '            'Me.Close()

        '            log.Info($"Report generation completed.")
        '            CustomLogUtil.Heading($"Report generation completed.")

        '            '2nd Sep 2024
        '            'If IO.Directory.Exists(mtcMtrModelObj.exportDirectoryLocation) Then
        '            '    Process.Start(mtcMtrModelObj.exportDirectoryLocation)
        '            'End If

        '            '4th Sep 2024
        '            waitFormObj.SetProgressInformationMessage("Excel report Created!")

        '            '2nd Sep 2024
        '            If IO.Directory.Exists(mtcMtrModelObj.export_MTR_Report_DirectoryLocation) Then
        '                Process.Start(mtcMtrModelObj.export_MTR_Report_DirectoryLocation)
        '            End If

        '            '2nd Sep 2024
        '            If IO.Directory.Exists(mtcMtrModelObj.export_Routing_Report_DirectoryLocation) Then
        '                Process.Start(mtcMtrModelObj.export_Routing_Report_DirectoryLocation)
        '            End If

        '            WaitEndSave()
        '        Catch ex As Exception
        '            SolidEdgeCommunity.OleMessageFilter.Register()
        '            CustomLogUtil.Log("While Generating MTC MTR Report", ex.Message, ex.StackTrace)
        '        End Try


#End Region

    End Sub

    Private Function GetMTC_MTR_ReportData(ByVal mtcMTRModelObj As MTC_MTR_Model) As MTC_MTR_Model
        '9th Sep 2024 'Added try..catch
        Try
            'BEC

            mtcMTRModelObj = GetBECReportData(mtcMTRModelObj)

            'DGS
            mtcMTRModelObj = GetNonBECReportData(mtcMTRModelObj)


            Return mtcMTRModelObj
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Function

    '17th Sep 2024
    Private Function GetMTC_MTR_ReportDataPartDoc(ByVal mtcMTRModelObj As MTC_MTR_Model) As MTC_MTR_Model
        '9th Sep 2024 'Added try..catch
        Try
            'BEC

            mtcMTRModelObj = GetBECReportData(mtcMTRModelObj)

            'DGS
            mtcMTRModelObj = GetNonBECReportData(mtcMTRModelObj)



        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return mtcMTRModelObj
    End Function

    '17th Sep 2024
    Private Function GetMTC_MTR_ReportDataSheetMetalDoc(ByVal mtcMTRModelObj As MTC_MTR_Model) As MTC_MTR_Model
        '9th Sep 2024 'Added try..catch
        Try
            'BEC

            mtcMTRModelObj = GetBECReportData(mtcMTRModelObj)

            'DGS
            mtcMTRModelObj = GetNonBECReportData(mtcMTRModelObj)



        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return mtcMTRModelObj
    End Function

    Private Function IsValidStructurePart(ByVal mtcReviewObj As DocumentModel) As Boolean

        '
        Dim resValid As Boolean = False

        If mtcReviewObj.Title.ToUpper().Contains("TUBING") _
            Or mtcReviewObj.Title.ToUpper().Contains("FLAT") _
            Or mtcReviewObj.Title.ToUpper().Contains("BAR") _
            Or mtcReviewObj.Title.ToUpper().Contains("ROUND") _
            Or mtcReviewObj.Title.ToUpper().Contains("CHANNEL") _
            Or mtcReviewObj.Title.ToUpper().Contains("ANGLE") _
            Or mtcReviewObj.Title.ToUpper().Contains("KEY STOCK") _
            Or mtcReviewObj.Title.ToUpper().Contains("SQUARE") _
            Or mtcReviewObj.Title.ToUpper().Contains("STOCK") _
            Then

            resValid = True
        End If

        Return resValid

    End Function

    Private Function GetBECReportData(ByVal mtcMtrModelObj As MTC_MTR_Model) As MTC_MTR_Model
        Try


            Dim lstCompletedFileName As New List(Of String)()

            Dim rCnt As Integer = 1

            log.Info("BEC report data collection start")
            log.Info("=======================")
            CustomLogUtil.Heading("BEC report data collection start")
            For Each dr As DataRow In mtcMtrModelObj.dtBECAuthorAssemblyData.Rows

                Dim mtcReviewObj As New DocumentModel()

                mtcReviewObj = SetVariablesData(dr, mtcReviewObj, "BEC")


                'TEMP24AUG2023
                'Comment if you dont want to set BOM Count as partlist count
                'mtcReviewObj.partlistcount = mtcMtrModelObj.BOMCount

                waitFormObj.SetProgressInformationMessage($"BEC {mtcReviewObj.fileName}")

                waitFormObj.SetProgressCountMessage($"{rCnt}/{mtcMtrModelObj.dtBECAuthorAssemblyData.Rows.Count}")

                log.Info($"BEC {rCnt} {mtcReviewObj.fileName}")
                CustomLogUtil.Log($"BEC {rCnt} {mtcReviewObj.fileName}")

                log.Info($"BEC {rCnt} {mtcReviewObj.fullpath}")
                CustomLogUtil.Log($"BEC {rCnt} {mtcReviewObj.fullpath}")

                mtcReviewObj.isValidBaseLineDirectoryPath = IsValidBaseLineDirectoryPath(mtcMtrModelObj.baseLineDirectoryLocation, mtcReviewObj.fullpath)

                If lstCompletedFileName.Contains(mtcReviewObj.fileNameWithoutExt) Then
                    Continue For
                Else
                    lstCompletedFileName.Add(mtcReviewObj.fileNameWithoutExt)
                End If

                'Create new function

                Dim dv As New DataView(mtcMtrModelObj.dtM2M)
                Dim partno As String = ExcelUtil.ExcelMtcReview.fpartno.ToString()
                dv.RowFilter = $"{partno}='{mtcReviewObj.fileNameWithoutExt}'"



                Dim m2MDataObj As New M2MDataModel()
                m2MDataObj = SetM2mData(m2MDataObj, dv)

                If dv.Count = 0 Then
                    mtcReviewObj.ispartfound = False
                Else
                    mtcReviewObj.ispartfound = True
                End If

                '==

                If mtcReviewObj.fullpath.EndsWith(".psm") Then

                    mtcReviewObj = SetSheetMetalData(mtcReviewObj)

                    'MTC

                    Dim mtcSheetMetalObj As New MTC_SheetMetal()

                    mtcSheetMetalObj = GetMTCSheetMetal(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                    mtcMtrModelObj.mtcSheetMetalList_BEC.Add(mtcSheetMetalObj)

                    'MTR

                    Dim mtrSheetMetalObj As New MTR_SheetMetal()

                    mtrSheetMetalObj = GetMTRSheetMetal(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                    mtcMtrModelObj.mtrSheetMetalList_BEC.Add(mtrSheetMetalObj)

                    'Routing Sequence Report

                    Dim routingSeqSheetMetalObj As New RoutingSequence_SheetMetal()

                    routingSeqSheetMetalObj = GetRoutingSequenceSheetMetal(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                    mtcMtrModelObj.routingSequenceSheetMetalList_BEC.Add(routingSeqSheetMetalObj)

                    '=====

                ElseIf mtcReviewObj.fullpath.EndsWith(".par") Then

                    mtcReviewObj = SetPartData(mtcReviewObj)

                    If mtcReviewObj.isElectrical = True Then

                        'mtc
                        Dim mtcPartObj As New MTC_Electrical()

                        mtcPartObj = GetMTCElectricalPart(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                        mtcMtrModelObj.mtcElectricalPartList_BEC.Add(mtcPartObj)
                    Else

                        If mtcReviewObj.isBaseline = False Then

                            'mtc
                            Dim mtcPartObj As New MTC_Part()

                            mtcPartObj = GetMTCPart(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                            mtcMtrModelObj.mtcPartList_BEC.Add(mtcPartObj)

                            'mtr
                            Dim mtrPartObj As New MTR_Part()

                            mtrPartObj = GetMTRPart(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                            mtcMtrModelObj.mtrPartList_BEC.Add(mtrPartObj)
                        Else

                            Dim mtcPartBaseLineObj As New MTC_BaseLine()

                            mtcPartBaseLineObj = GetMTCPartBaseLine(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                            mtcMtrModelObj.mtcBaseLineList_BEC.Add(mtcPartBaseLineObj)

                        End If

                        If Not mtcReviewObj.category.ToUpper() = "HARDWARE" Then
                            'Contains something
                            'if valid for routing sequence
                            'Routing Sequence Report

                            If IsValidStructurePart(mtcReviewObj) Then

                                Dim routingSeqSheetMetalObj As New RoutingSequence_Structure()

                                routingSeqSheetMetalObj = GetRoutingSequenceStructure(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                                mtcMtrModelObj.routingSequenceStructureList_BEC.Add(routingSeqSheetMetalObj)

                            End If

                        End If

                    End If

                    If Not IsValidStructurePart(mtcReviewObj) Then

                        ' If part is not valida structure part then MISC
                        Dim routingSequenceMiscObj As New RoutingSequence_Structure()

                        routingSequenceMiscObj = GetRoutingSequenceStructure(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                        mtcMtrModelObj.routingSequenceMiscList_BEC.Add(routingSequenceMiscObj)
                    End If


                ElseIf mtcReviewObj.fullpath.EndsWith(".asm") Then

                    'mtc

                    mtcReviewObj = SetAssemblyData(mtcReviewObj)

                    Dim mtcAssemblyObj As MTC_Assembly = GetMTCAssembly(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                    mtcMtrModelObj.mtcAssemblyList_BEC.Add(mtcAssemblyObj)

                    'mtr

                    Dim mtrAssemblyObj As MTR_Assembly = GetMTRAssembly(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                    mtcMtrModelObj.mtrAssemblyList_BEC.Add(mtrAssemblyObj)

                    'routing Sequence Assembly

                    Dim routingSeqSheetMetalObj As New RoutingSequence_Assembly()

                    routingSeqSheetMetalObj = GetRoutingSequenceAssembly(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                    mtcMtrModelObj.routingSequenceAssemblyList_BEC.Add(routingSeqSheetMetalObj)

                End If

                rCnt += 1

            Next

            log.Info("BEC report data collection end")
            log.Info("=======================")
            CustomLogUtil.Heading("BEC report data collection end")

        Catch ex As Exception
            CustomLogUtil.Log("While fetching BEC Report Data", ex.Message, ex.StackTrace)
        End Try
        Return mtcMtrModelObj
    End Function

    Private Function GetNonBECReportData(ByVal mtcMtrModelObj As MTC_MTR_Model) As MTC_MTR_Model
        Try
            Dim lstCompletedFileName As New List(Of String)()

            Dim rCnt As Integer = 1

            log.Info("DGS report data collection start")
            log.Info("=======================")
            CustomLogUtil.Heading("DGS report data collection start")

            For Each dr As DataRow In mtcMtrModelObj.dtNonBECAuthorAssemblyData.Rows

                Dim mtcReviewObj As New DocumentModel()

                mtcReviewObj = SetVariablesData(dr, mtcReviewObj, "DGS")

                waitFormObj.SetProgressInformationMessage($"DGS {mtcReviewObj.fileName}")

                waitFormObj.SetProgressCountMessage($"{rCnt}/{mtcMtrModelObj.dtNonBECAuthorAssemblyData.Rows.Count}")

                log.Info($"DGS {rCnt} {mtcReviewObj.fileName}")
                CustomLogUtil.Log($"DGS {rCnt} {mtcReviewObj.fileName}")
                'mtcReviewObj.fullpath
                log.Info($"DGS {rCnt} {mtcReviewObj.fullpath}")
                CustomLogUtil.Log($"DGS {rCnt} {mtcReviewObj.fullpath}")

                mtcReviewObj.isValidBaseLineDirectoryPath = IsValidBaseLineDirectoryPath(mtcMtrModelObj.baseLineDirectoryLocation, mtcReviewObj.fullpath)

                If lstCompletedFileName.Contains(mtcReviewObj.fileNameWithoutExt) Then
                    Continue For
                Else
                    lstCompletedFileName.Add(mtcReviewObj.fileNameWithoutExt)
                End If

                'Create new function

                Dim dv As New DataView(mtcMtrModelObj.dtM2M)
                Dim partno As String = ExcelUtil.ExcelMtcReview.fpartno.ToString()
                dv.RowFilter = $"{partno}='{mtcReviewObj.fileNameWithoutExt}'"

                Dim m2MDataObj As New M2MDataModel()
                m2MDataObj = SetM2mData(m2MDataObj, dv)

                If dv.Count = 0 Then
                    mtcReviewObj.ispartfound = False
                Else
                    mtcReviewObj.ispartfound = True
                End If
                '==

                If mtcReviewObj.fullpath.EndsWith(".psm") Then

                    mtcReviewObj = SetSheetMetalData(mtcReviewObj)

                    'MTC

                    Dim mtcSheetMetalObj As New MTC_SheetMetal()

                    mtcSheetMetalObj = GetMTCSheetMetal(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                    mtcMtrModelObj.mtcSheetMetalList_DGS.Add(mtcSheetMetalObj)

                    'MTR

                    Dim mtrSheetMetalObj As New MTR_SheetMetal()

                    mtrSheetMetalObj = GetMTRSheetMetal(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                    mtcMtrModelObj.mtrSheetMetalList_DGS.Add(mtrSheetMetalObj)

                    'Routing Sequence Report

                    Dim routingSeqSheetMetalObj As New RoutingSequence_SheetMetal()

                    routingSeqSheetMetalObj = GetRoutingSequenceSheetMetal(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                    mtcMtrModelObj.routingSequenceSheetMetalList_DGS.Add(routingSeqSheetMetalObj)

                ElseIf mtcReviewObj.fullpath.EndsWith(".par") Then

                    mtcReviewObj = SetPartData(mtcReviewObj)

                    If mtcReviewObj.isElectrical = True Then

                        'mtc
                        Dim mtcPartObj As New MTC_Electrical()

                        mtcPartObj = GetMTCElectricalPart(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                        mtcMtrModelObj.mtcElectricalPartList_DGS.Add(mtcPartObj)
                    Else

                        If mtcReviewObj.isBaseline = False Then

                            'MTC

                            Dim mtcPartObj As New MTC_Part()

                            mtcPartObj = GetMTCPart(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                            mtcMtrModelObj.mtcPartList_DGS.Add(mtcPartObj)

                            'mtr
                            Dim mtrPartObj As New MTR_Part()

                            mtrPartObj = GetMTRPart(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                            mtcMtrModelObj.mtrPartList_DGS.Add(mtrPartObj)
                        Else

                            Dim mtcPartBaseLineObj As New MTC_BaseLine()

                            mtcPartBaseLineObj = GetMTCPartBaseLine(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                            mtcMtrModelObj.mtcBaseLineList_DGS.Add(mtcPartBaseLineObj)

                        End If

                        If Not mtcReviewObj.category.ToUpper() = "HARDWARE" Then
                            'Contains something
                            'if valid for routing sequence
                            'Routing Sequence Report

                            If IsValidStructurePart(mtcReviewObj) Then
                                Dim routingSeqSheetMetalObj As New RoutingSequence_Structure()

                                routingSeqSheetMetalObj = GetRoutingSequenceStructure(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                                mtcMtrModelObj.routingSequenceStructureList_DGS.Add(routingSeqSheetMetalObj)
                            End If

                        End If
                    End If


                    If Not IsValidStructurePart(mtcReviewObj) Then

                        ' If part is not valida structure part then MISC
                        Dim routingSequenceMiscObj As New RoutingSequence_Structure()

                        routingSequenceMiscObj = GetRoutingSequenceStructure(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                        mtcMtrModelObj.routingSequenceMiscList_DGS.Add(routingSequenceMiscObj)
                    End If

                ElseIf mtcReviewObj.fullpath.EndsWith(".asm") Then

                    mtcReviewObj = SetAssemblyData(mtcReviewObj)

                    'MTC

                    Dim mtcAssemblyObj As MTC_Assembly = GetMTCAssembly(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                    mtcMtrModelObj.mtcAssemblyList_DGS.Add(mtcAssemblyObj)

                    'mtr

                    Dim mtrAssemblyObj As MTR_Assembly = GetMTRAssembly(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                    mtcMtrModelObj.mtrAssemblyList_DGS.Add(mtrAssemblyObj)


                    'routing Sequence Assembly

                    Dim routingSeqSheetMetalObj As New RoutingSequence_Assembly()

                    routingSeqSheetMetalObj = GetRoutingSequenceAssembly(mtcReviewObj, mtcMtrModelObj, m2MDataObj)

                    mtcMtrModelObj.routingSequenceAssemblyList_DGS.Add(routingSeqSheetMetalObj)

                End If

                rCnt += 1

            Next

            log.Info("DGS report data collection end")
            log.Info("=======================")
            CustomLogUtil.Heading("DGS report data collection end")

        Catch ex As Exception
            CustomLogUtil.Log("While fetching DGS report data", ex.Message, ex.StackTrace)
        End Try
        Return mtcMtrModelObj
    End Function

    Private Function GetMTRAssembly(ByVal mtcReviewObj As DocumentModel,
                                    ByVal mtcMtrModelObj As MTC_MTR_Model, ByVal m2MDataObj As M2MDataModel) As MTR_Assembly
        Dim mtcAssemblyObj As New MTR_Assembly()
        Try

            mtcAssemblyObj.assemblyName = mtcReviewObj.fileNameWithoutExt

            '1.
            mtcAssemblyObj.isAdjustable = mtcReviewObj.isadjustable

            '2.
            mtcAssemblyObj.isInterPartCopiesDetected = mtcReviewObj.interPartCopiesDetected

            '3.
            mtcAssemblyObj.isPartCopiesDetected = mtcReviewObj.partCopiesDetected

            '4.
            If mtcReviewObj.author = Nothing Or mtcReviewObj.Title = Nothing Or mtcReviewObj.materialused = Nothing _
            Or mtcReviewObj.matlspec = Nothing Or mtcReviewObj.projectname = Nothing _
            Or mtcReviewObj.revisionNumber_Prop = Nothing Or mtcReviewObj.documentno = Nothing Or mtcReviewObj.UomProperty = Nothing Or mtcReviewObj.category = Nothing Or mtcReviewObj.comments = Nothing Or mtcReviewObj.keywords = Nothing Then
                mtcAssemblyObj.isValidAllCategories = "No"
            Else
                mtcAssemblyObj.isValidAllCategories = "Yes"
            End If

            '5.
            mtcAssemblyObj.isAssemblyFeatureExist = mtcReviewObj.checkAssemblyFeature

            '6.
            mtcAssemblyObj.isMatingPartInterferenceChecked = If(mtcReviewObj.statusfile = True, "Yes", "No")

            '7.
            mtcAssemblyObj.verifyInterference = If(mtcReviewObj.interference = True, "Yes", "No")

            '8.
            mtcAssemblyObj.verifyUpdateOnFileSave = String.Empty

            '9.
            mtcAssemblyObj.isGeometryBroken = mtcReviewObj.isGeometryBroken

            '10.
            mtcAssemblyObj.author = mtcReviewObj.author

            '11.
            mtcAssemblyObj.modifiedDate = mtcReviewObj.modifiedDate



        Catch ex As Exception
            CustomLogUtil.Log("While Getting MTR Assembly Data", ex.Message, ex.StackTrace)
        End Try
        Return mtcAssemblyObj
    End Function

    Private Function GetMTCAssembly(ByVal mtcReviewObj As DocumentModel,
                                    ByVal mtcMtrModelObj As MTC_MTR_Model, ByVal m2MDataObj As M2MDataModel) As MTC_Assembly
        Dim mtcAssemblyObj As New MTC_Assembly()
        Try

            mtcAssemblyObj.assemblyName = mtcReviewObj.fileNameWithoutExt

            '1.
            mtcAssemblyObj.ecoNumber = mtcReviewObj.ECO

            mtcAssemblyObj.isPartFound = mtcReviewObj.ispartfound

            '2.
            mtcAssemblyObj.partNumber = If(mtcReviewObj.ispartfound, "Yes", "No")

            '3.
            mtcAssemblyObj.revisionLevel = mtcReviewObj.revisionNumber_Prop

            '4.
            mtcAssemblyObj.author = mtcReviewObj.author

            mtcAssemblyObj.projectName = mtcReviewObj.projectname

            '5.
            mtcAssemblyObj.projectNameExist = If((mtcMtrModelObj.projectNameList.Contains(mtcReviewObj.projectname)), "Yes", "No")

            '6.
            mtcAssemblyObj.revisionNumberCorrect = If((mtcReviewObj.revisionNumber_Prop = mtcReviewObj.revisionNumber_FileName), "Yes", "No")

            mtcAssemblyObj.documentNumber = mtcReviewObj.documentno

            '7.mTCReviewObj.isUOMMatch_M2M
            mtcAssemblyObj.documentNumberCorrect = If((mtcReviewObj.fileNameWithoutExt.Contains(mtcReviewObj.documentno)), "Yes", "No")

            '8.
            mtcAssemblyObj.authorExists = If(mtcMtrModelObj.authorList.Contains(mtcReviewObj.author), "Yes", "No")

            '9.
            If mtcReviewObj.author = Nothing Or mtcReviewObj.Title = Nothing Or mtcReviewObj.materialused = Nothing _
            Or mtcReviewObj.matlspec = Nothing Or mtcReviewObj.projectname = Nothing _
            Or mtcReviewObj.revisionNumber_Prop = Nothing Or mtcReviewObj.documentno = Nothing Or mtcReviewObj.UomProperty = Nothing Or mtcReviewObj.category = Nothing Or mtcReviewObj.comments = Nothing Or mtcReviewObj.keywords = Nothing Then
                'If mtcReviewObj.author = Nothing Or mtcReviewObj.Title = Nothing Or mtcReviewObj.materialused = Nothing _
                'Or mtcReviewObj.matlspec = Nothing Or mtcReviewObj.projectname = Nothing _
                'Or mtcReviewObj.revisionNumber_Prop = Nothing Or mtcReviewObj.documentno = Nothing Or mtcReviewObj.UomProperty = Nothing Or mtcReviewObj.category = Nothing Or mtcReviewObj.keywords = Nothing Then
                mtcAssemblyObj.isDashPopulated = "No"
            Else
                mtcAssemblyObj.isDashPopulated = "Yes"
            End If

            mtcAssemblyObj.title = mtcReviewObj.Title
            '10.

            mtcAssemblyObj.isTitleMatch_ItemMaster = If(m2MDataObj.M2Mdescript.Trim() = mtcReviewObj.Title.Trim(), "Yes", "No")

            mtcAssemblyObj.UomProperty = mtcReviewObj.UomProperty

            '11.
            mtcAssemblyObj.isUOMMatch_M2M = If(mtcReviewObj.UomProperty.Contains(m2MDataObj.M2Mmeasure), "Yes", "No")
            'If Not mtcReviewObj.UomProperty = "" Then
            '    mtcAssemblyObj.isUOMMatch_M2M = If(mtcReviewObj.UomProperty.Contains(m2MDataObj.M2Mmeasure), "Yes", "No")
            'Else
            '    mtcAssemblyObj.isUOMMatch_M2M = "No"
            'End If


            '12.
            'TEMP24AUG2023
            'Comment if you dont want to set BOM Count as partlist count
            'mtcAssemblyObj.partListCount = mtcMtrModelObj.BOMCount
            mtcAssemblyObj.partListCount = mtcReviewObj.partlistcount

            '13.
            mtcAssemblyObj.isInterferenceFound = If(mtcReviewObj.statusfile = True And mtcReviewObj.interference = True, "Yes", "No")

            '14.
            mtcAssemblyObj.isInterPartCopiesDetected = mtcReviewObj.interPartCopiesDetected

            '15.
            mtcAssemblyObj.isPartCopiesDetected = mtcReviewObj.partCopiesDetected

            '16.
            mtcAssemblyObj.isBrokenFilePathDetected = mtcReviewObj.documentLinkBroken

            '17.
            mtcAssemblyObj.modifiedDate = mtcReviewObj.modifiedDate

            '18.
            mtcAssemblyObj.isAdjustable = mtcReviewObj.isadjustable

        Catch ex As Exception
            CustomLogUtil.Log("While Fetching MTRAssembly", ex.Message, ex.StackTrace)
        End Try
        Return mtcAssemblyObj
    End Function

    Private Function SetAssemblyData(ByVal mtcReviewObj As DocumentModel) As DocumentModel
        Try

            Dim documents As SolidEdgeFramework.Documents = objApp.Documents

            Dim asmdoc As SolidEdgeAssembly.AssemblyDocument = Nothing

            If Not IO.File.Exists(mtcReviewObj.fullpath) Then
            Else

                asmdoc = DirectCast(documents.Open(mtcReviewObj.fullpath), SolidEdgeFramework.SolidEdgeDocument)

                Threading.Thread.Sleep(2000)

                Try
                    Dim occurrences As SolidEdgeAssembly.Occurrences = asmdoc.Occurrences
                    Dim partlst As New List(Of String)

                    For Each occur As SolidEdgeAssembly.Occurrence In occurrences
                        If Not partlst.Contains(occur.Name) Then
                            partlst.Add(occur.Name)
                            mtcReviewObj.partlistcount += 1
                        End If

                    Next
                    mtcReviewObj.checkAssemblyFeature = CheckAssemblyFeatureExistence(asmdoc)

                    mtcReviewObj.interPartLink = CheckInterPartLinksPSM()

                    mtcReviewObj.isGeometryBroken = IsGeomtryBroken_assembly(asmdoc)
                Catch ex As Exception
                    MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
                End Try

                asmdoc.Close(False)
            End If

        Catch ex As Exception
            CustomLogUtil.Log("While Setting the Assembly Data", ex.Message, ex.StackTrace)
        End Try

        Return mtcReviewObj
    End Function

    Private Function IsGeomtryBroken_assembly(ByRef objSheetMetalDocument As SolidEdgeAssembly.AssemblyDocument) As String

        Dim geoMetryBroken As String = "Yes"
        Try
            Dim ipl As SolidEdgeFramework.InterpartLinks = objSheetMetalDocument.InterpartLinks
            If ipl.Count > 1 Then
                geoMetryBroken = "No"
            End If
        Catch ex As Exception
            MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
            CustomLogUtil.Log("While fetching the Broken Geometry Assembly", ex.Message, ex.StackTrace)
        End Try
        Return geoMetryBroken

    End Function

    Private Function CheckAssemblyFeatureExistence(ByRef asssemblyDocObj As AssemblyDocument) As String

        Dim result As String = String.Empty
        Try

            Dim asmModelObj As SolidEdgePart.Model = asssemblyDocObj.AssemblyModel
            Dim asmModelFeatures As Features = asmModelObj.Features
            Dim cnt As Integer = asmModelFeatures.Count
            If cnt > 0 Then
                result = "Yes"
            Else
                result = "No"
            End If
        Catch ex As Exception
            result = ex.Message
            MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")

        End Try
        Return result
    End Function

    Private Function GetMTRSheetMetal(ByVal mtcReviewObj As DocumentModel,
                                    ByVal mtcMtrModelObj As MTC_MTR_Model, ByVal m2MDataObj As M2MDataModel) As MTR_SheetMetal

        '0.
        '1.
        '2.
        '3.
        '4.
        '5.
        Dim mtrSheetMetalObj As New MTR_SheetMetal With {
            .assemblyName = mtcReviewObj.fileNameWithoutExt,
            .isFeatureFullyConstrained = mtcReviewObj.sketchisfullydefined,
            .verifySuppressFeature = mtcReviewObj.issupress,
            .isAdjustable = mtcReviewObj.isadjustable,
            .isInterPartCopiesDetected = mtcReviewObj.interPartCopiesDetected,
            .isPartCopiesDetected = mtcReviewObj.partCopiesDetected
        }

        '6.
        If mtcReviewObj.author = Nothing Or mtcReviewObj.Title = Nothing Or mtcReviewObj.materialused = Nothing _
            Or mtcReviewObj.matlspec = Nothing Or mtcReviewObj.projectname = Nothing _
            Or mtcReviewObj.revisionNumber_Prop = Nothing Or mtcReviewObj.documentno = Nothing Or mtcReviewObj.UomProperty = Nothing Or mtcReviewObj.category = Nothing Or mtcReviewObj.comments = Nothing Or mtcReviewObj.keywords = Nothing Then
            mtrSheetMetalObj.isValidAllCategories = "No"
        Else
            mtrSheetMetalObj.isValidAllCategories = "Yes"
        End If

        '7.
        mtrSheetMetalObj.verifyWeightMass = mtcReviewObj.density

        '8.
        mtrSheetMetalObj.verifyUpdateOnFileSave = ""

        '9.
        mtrSheetMetalObj.isGeometryBroken = mtcReviewObj.isGeometryBroken

        '10.
        mtrSheetMetalObj.author = mtcReviewObj.author

        '11.
        mtrSheetMetalObj.modifiedDate = mtcReviewObj.modifiedDate

        mtrSheetMetalObj.isValidPart = mtcReviewObj.isValidPart

        Return mtrSheetMetalObj

    End Function

    Private Function GetRoutingSequenceAssembly(ByVal mtcReviewObj As DocumentModel,
                                    ByVal mtcMtrModelObj As MTC_MTR_Model, ByVal m2MDataObj As M2MDataModel) As RoutingSequence_Assembly
        Dim mtcSheetMetalObj As New RoutingSequence_Assembly With {
            .partNumber = mtcReviewObj.fileNameWithoutExt,
            .massItem = mtcReviewObj.massItem,
            .m2mFSource = m2MDataObj.M2Msource,
            .projectName = mtcReviewObj.projectname,
            .title = mtcReviewObj.Title,
            .filePath = mtcReviewObj.fullpath,
            .lastAuthor = mtcReviewObj.lastAuthor,
            .floc = m2MDataObj.M2Mflocation,
            .fbin = m2MDataObj.M2MFbin,
            .qAQC = mtcReviewObj.qAQC,
            .quantity = mtcReviewObj.quantity
        }

        Return mtcSheetMetalObj

    End Function

    Private Function GetRoutingSequenceStructure(ByVal mtcReviewObj As DocumentModel,
                                    ByVal mtcMtrModelObj As MTC_MTR_Model, ByVal m2MDataObj As M2MDataModel) As RoutingSequence_Structure

        Dim mtcSheetMetalObj As New RoutingSequence_Structure()

        mtcSheetMetalObj.assemblyName = mtcReviewObj.fileNameWithoutExt
        mtcSheetMetalObj.partNumber = mtcReviewObj.fileNameWithoutExt
        mtcSheetMetalObj.material = mtcReviewObj.material
        mtcSheetMetalObj.materialSpec = mtcReviewObj.matlspec
        mtcSheetMetalObj.materialUsed = mtcReviewObj.materialused
        mtcSheetMetalObj.isValidPart = mtcReviewObj.isValidPart
        mtcSheetMetalObj.m2mfSource = m2MDataObj.M2Msource
        mtcSheetMetalObj.material = mtcReviewObj.material
        mtcSheetMetalObj.holeQty = mtcReviewObj.holeQty
        mtcSheetMetalObj.holeFeature = mtcReviewObj.isHoleFeatureExists
        mtcSheetMetalObj.holeFit = mtcReviewObj.holeFit
        mtcSheetMetalObj.materialDescription = mtcReviewObj.Title
        mtcSheetMetalObj.filePath = mtcReviewObj.fullpath
        mtcSheetMetalObj.projectName = mtcReviewObj.projectname
        mtcSheetMetalObj.massItem = mtcReviewObj.massItem
        mtcSheetMetalObj.m2mflocation = m2MDataObj.M2Mflocation
        mtcSheetMetalObj.m2mFbin = m2MDataObj.M2MFbin
        mtcSheetMetalObj.quantity = mtcReviewObj.quantity
        mtcSheetMetalObj.category = mtcReviewObj.category

        Return mtcSheetMetalObj

    End Function

    Private Function GetRoutingSequenceSheetMetal(ByVal mtcReviewObj As DocumentModel,
                                    ByVal mtcMtrModelObj As MTC_MTR_Model, ByVal m2MDataObj As M2MDataModel) As RoutingSequence_SheetMetal

        Dim mtcSheetMetalObj As New RoutingSequence_SheetMetal()

        mtcSheetMetalObj.assemblyName = mtcReviewObj.fileNameWithoutExt

        mtcSheetMetalObj.partNumber = mtcReviewObj.fileNameWithoutExt

        mtcSheetMetalObj.material = mtcReviewObj.material

        mtcSheetMetalObj.materialThickness = mtcReviewObj.materialThickness

        mtcSheetMetalObj.materialSpec = mtcReviewObj.matlspec

        mtcSheetMetalObj.materialUsed = mtcReviewObj.materialused

        mtcSheetMetalObj.density = mtcReviewObj.density

        mtcSheetMetalObj.bendRadius = mtcReviewObj.bendRadius

        mtcSheetMetalObj.flat_Pattern_Model_CutSizeX = mtcReviewObj.flat_Pattern_Model_CutSizeX

        mtcSheetMetalObj.flat_Pattern_Model_CutSizeY = mtcReviewObj.flat_Pattern_Model_CutSizeY

        mtcSheetMetalObj.isValidPart = mtcReviewObj.isValidPart

        mtcSheetMetalObj.m2mfSource = m2MDataObj.M2Msource

        mtcSheetMetalObj.material = mtcReviewObj.material
        mtcSheetMetalObj.materialThickness = mtcReviewObj.materialThickness
        mtcSheetMetalObj.bendRadius = mtcReviewObj.bendRadius
        mtcSheetMetalObj.flat_Pattern_Model_CutSizeX = mtcReviewObj.flat_Pattern_Model_CutSizeX
        mtcSheetMetalObj.flat_Pattern_Model_CutSizeY = mtcReviewObj.flat_Pattern_Model_CutSizeY

        mtcSheetMetalObj.holeQty = mtcReviewObj.holeQty
        mtcSheetMetalObj.holeFeature = mtcReviewObj.isHoleFeatureExists
        mtcSheetMetalObj.bendQty = mtcReviewObj.bendQty

        mtcSheetMetalObj.louverExists = mtcReviewObj.louverExists
        mtcSheetMetalObj.hemExists = mtcReviewObj.hemExists
        mtcSheetMetalObj.beadExists = mtcReviewObj.beadExists
        mtcSheetMetalObj.gussetExists = mtcReviewObj.gussetExists

        mtcSheetMetalObj.hem_Bead_GussetExists = mtcReviewObj.hem_Bead_GussetExists
        mtcSheetMetalObj.holeFit = mtcReviewObj.holeFit

        mtcSheetMetalObj.materialDescription = mtcReviewObj.Title
        mtcSheetMetalObj.filePath = mtcReviewObj.fullpath
        mtcSheetMetalObj.projectName = mtcReviewObj.projectname

        mtcSheetMetalObj.massItem = mtcReviewObj.massItem

        mtcSheetMetalObj.m2mflocation = m2MDataObj.M2Mflocation
        mtcSheetMetalObj.m2mFbin = m2MDataObj.M2MFbin

        mtcSheetMetalObj.quantity = mtcReviewObj.quantity
        Return mtcSheetMetalObj

    End Function

    Private Function GetMTCSheetMetal(ByVal mtcReviewObj As DocumentModel,
                                    ByVal mtcMtrModelObj As MTC_MTR_Model, ByVal m2MDataObj As M2MDataModel) As MTC_SheetMetal

        Dim mtcSheetMetalObj As New MTC_SheetMetal()

        '0.
        mtcSheetMetalObj.assemblyName = mtcReviewObj.fileNameWithoutExt

        '1.
        mtcSheetMetalObj.ecoNumber = mtcReviewObj.ECO

        '2.
        mtcSheetMetalObj.partNumber = If(mtcReviewObj.ispartfound, "Yes", "No")

        '3.
        mtcSheetMetalObj.revisionLevel = mtcReviewObj.revisionNumber_Prop

        '4.
        mtcSheetMetalObj.author = mtcReviewObj.author

        mtcSheetMetalObj.projectName = mtcReviewObj.projectname

        '5.
        mtcSheetMetalObj.projectNameExist = If(mtcMtrModelObj.projectNameList.Contains(mtcReviewObj.projectname), "Yes", "No")

        '6.
        mtcSheetMetalObj.revisionNumberCorrect = If(mtcSheetMetalObj.revisionLevel = mtcReviewObj.revisionNumber_FileName, "Yes", "No")

        mtcSheetMetalObj.documentNumber = mtcReviewObj.documentno
        '7.
        mtcSheetMetalObj.documentNumberCorrect = If(mtcReviewObj.fileNameWithoutExt.Contains(mtcReviewObj.documentno), "Yes", "No")

        '8.
        mtcSheetMetalObj.authorExists = If(mtcMtrModelObj.authorList.Contains(mtcReviewObj.author), "Yes", "No")

        '9.
        If mtcReviewObj.author = Nothing Or mtcReviewObj.Title = Nothing Or mtcReviewObj.materialused = Nothing _
            Or mtcReviewObj.matlspec = Nothing Or mtcReviewObj.projectname = Nothing _
            Or mtcReviewObj.revisionNumber_Prop = Nothing Or mtcReviewObj.documentno = Nothing Or mtcReviewObj.UomProperty = Nothing Or mtcReviewObj.category = Nothing Or mtcReviewObj.comments = Nothing Or mtcReviewObj.keywords = Nothing Then
            mtcSheetMetalObj.isDashPopulated = "No"
        Else
            mtcSheetMetalObj.isDashPopulated = "Yes"
        End If

        mtcSheetMetalObj.title = mtcReviewObj.Title

        '10.
        mtcSheetMetalObj.isTitleMatch_ItemMaster = If(m2MDataObj.M2Mdescript.Trim() = mtcReviewObj.Title.Trim(), "Yes", "No")

        mtcSheetMetalObj.UomProperty = mtcReviewObj.UomProperty

        '11.
        mtcSheetMetalObj.isUOMMatch_M2M = If(mtcReviewObj.UomProperty.Contains($"({m2MDataObj.M2Mmeasure})"), "Yes", "No")

        mtcSheetMetalObj.materialSpec = mtcReviewObj.matlspec

        '12.
        mtcSheetMetalObj.isMaterialSpecExists = If(mtcReviewObj.matlspec = "", "No", "Yes")

        mtcSheetMetalObj.materialUsed = mtcReviewObj.materialused

        '13.
        mtcSheetMetalObj.isMaterialUsedExists = If(mtcReviewObj.materialused = "", "No", "Yes")

        '14.
        mtcSheetMetalObj.gageExcelFile = mtcReviewObj.gageeexcelfile

        '15.
        mtcSheetMetalObj.isFlatPatternActive = mtcReviewObj.isflatpattern

        '16.
        mtcSheetMetalObj.holeToolsUsed = mtcReviewObj.iscutout

        '17.
        mtcSheetMetalObj.isAdjustatble = mtcReviewObj.isadjustable

        '18.
        mtcSheetMetalObj.m2mSource = GetM2MSource(m2MDataObj.M2Msource)

        '19.
        mtcSheetMetalObj.modifiedDate = mtcReviewObj.modifiedDate

        mtcSheetMetalObj.isValidPart = mtcReviewObj.isValidPart

        mtcSheetMetalObj.material = mtcReviewObj.material
        mtcSheetMetalObj.materialThickness = mtcReviewObj.materialThickness
        mtcSheetMetalObj.bendRadius = mtcReviewObj.bendRadius
        mtcSheetMetalObj.flat_Pattern_Model_CutSizeX = mtcReviewObj.flat_Pattern_Model_CutSizeX
        mtcSheetMetalObj.flat_Pattern_Model_CutSizeY = mtcReviewObj.flat_Pattern_Model_CutSizeY

        mtcSheetMetalObj.holeQty = mtcReviewObj.holeQty
        mtcSheetMetalObj.holeFeatureExists = mtcReviewObj.isHoleFeatureExists
        mtcSheetMetalObj.bendQty = mtcReviewObj.bendQty

        mtcSheetMetalObj.louverExists = mtcReviewObj.louverExists
        mtcSheetMetalObj.hemExists = mtcReviewObj.hemExists
        mtcSheetMetalObj.beadExists = mtcReviewObj.beadExists
        mtcSheetMetalObj.gussetExists = mtcReviewObj.gussetExists
        mtcSheetMetalObj.hem_Bead_GussetExists = mtcReviewObj.hem_Bead_GussetExists
        mtcSheetMetalObj.holeFit = mtcReviewObj.holeFit

        mtcSheetMetalObj.filePath = mtcReviewObj.fullpath
        mtcSheetMetalObj.materialDesc = mtcReviewObj.Title

        Return mtcSheetMetalObj

    End Function

    Private Function GetMTCPartBaseLine(ByVal mtcReviewObj As DocumentModel,
                                    ByVal mtcMtrModelObj As MTC_MTR_Model, ByVal m2MDataObj As M2MDataModel) As MTC_BaseLine

        Dim mtcBaseLineObj As New MTC_BaseLine()

        mtcBaseLineObj.assemblyName = mtcReviewObj.fileNameWithoutExt

        mtcBaseLineObj.isValidBaseLineDirectoryPath = mtcReviewObj.isValidBaseLineDirectoryPath

        '1.
        mtcBaseLineObj.isPartFound = If(mtcReviewObj.ispartfound = True, "Yes", "No")

        '2.
        mtcBaseLineObj.revisionLevel = mtcReviewObj.revisionNumber_Prop

        '3.
        mtcBaseLineObj.author = mtcReviewObj.author

        '4.
        mtcBaseLineObj.category = mtcReviewObj.category

        '5.
        mtcBaseLineObj.isVirtualThreadExists = mtcReviewObj.isThreadExists

        '6.
        mtcBaseLineObj.isSketchFullyDefined = mtcReviewObj.sketchisfullydefined

        '7.
        mtcBaseLineObj.isSuppressFeatureFound = mtcReviewObj.issupress

        mtcBaseLineObj.materialSpec = mtcReviewObj.matlspec

        '8.
        mtcBaseLineObj.isMaterialSpecExists = If(mtcReviewObj.matlspec = "", "No", "Yes")

        mtcBaseLineObj.materialUsed = mtcReviewObj.materialused

        '9.
        mtcBaseLineObj.isMaterialUsedExists = If(mtcReviewObj.materialused = "", "No", "Yes")

        mtcBaseLineObj.title = mtcReviewObj.Title

        '10.
        mtcBaseLineObj.isMaterialDesc_Title = If(m2MDataObj.M2Mdescript.Trim() = mtcReviewObj.Title.Trim(), "Yes", "No")

        '11.
        mtcBaseLineObj.isAuthorExists = If(mtcMtrModelObj.authorList.Contains(mtcReviewObj.author), "Yes", "No")

        '12.

        mtcBaseLineObj.title = mtcReviewObj.Title

        '13.
        'mtcBaseLineObj.isCommentExist = If(m2MDataObj.commentformate.Replace(" ", "").Trim().ToUpper() = mtcReviewObj.comments.Replace(" ", "").Trim().ToUpper(), "Yes", "No")
        mtcBaseLineObj.isCommentExist = ""

        mtcBaseLineObj.documentNumber = mtcReviewObj.documentno

        '14.
        mtcBaseLineObj.isCorrectDocumentNumber = If(mtcReviewObj.fileNameWithoutExt.Contains(mtcReviewObj.documentno), "Yes", "No")

        '15.
        mtcBaseLineObj.isCorrectRevisionNumber = If(mtcReviewObj.revisionNumber_Prop = mtcReviewObj.revisionNumber_FileName, "Yes", "No")

        '16.
        mtcBaseLineObj.isCorrectProjectName = mtcReviewObj.isBrookVilleProject_Baseline

        mtcBaseLineObj.documentType = mtcReviewObj.Documenttype

        mtcBaseLineObj.hardwarePart = mtcReviewObj.hardwarepart

        '17.
        '26th september 2024
        'mtcBaseLineObj.isHardwarePartBoxChecked = If(mtcReviewObj.Documenttype = "BaseLined" And mtcReviewObj.hardwarepart = True, "Yes", "No")
        mtcBaseLineObj.isHardwarePartBoxChecked = If(mtcReviewObj.hardwarepart = True, "Yes", "No")

        '18.
        If mtcReviewObj.author = Nothing Or mtcReviewObj.Title = Nothing Or mtcReviewObj.materialused = Nothing _
            Or mtcReviewObj.matlspec = Nothing Or mtcReviewObj.projectname = Nothing _
            Or mtcReviewObj.revisionNumber_Prop = Nothing Or mtcReviewObj.documentno = Nothing Or mtcReviewObj.UomProperty = Nothing Or mtcReviewObj.category = Nothing Or mtcReviewObj.comments = Nothing Or mtcReviewObj.keywords = Nothing Then
            mtcBaseLineObj.isDashPopulated = "No"
        Else
            mtcBaseLineObj.isDashPopulated = "Yes"
        End If

        mtcBaseLineObj.UomProperty = mtcReviewObj.UomProperty

        '19.
        mtcBaseLineObj.isUOMMatch_M2M = If(mtcReviewObj.UomProperty.Contains(m2MDataObj.M2Mmeasure), "Yes", "No")

        '20.
        mtcBaseLineObj.isM2MSourceStocked = If(m2MDataObj.M2Msource = "S", "Yes", "No")

        '21.
        mtcBaseLineObj.isHoleToolUsed = mtcReviewObj.iscutout

        '22.

        '23.
        mtcBaseLineObj.isInterPartCopiesDetected = mtcReviewObj.interPartCopiesDetected

        '24.
        mtcBaseLineObj.isPartCopiesDetected = mtcReviewObj.partCopiesDetected

        '25.
        mtcBaseLineObj.isBrokenFilePathDetected = mtcReviewObj.documentLinkBroken

        '26.
        mtcBaseLineObj.isAdjustable = mtcReviewObj.isadjustable

        '27.
        mtcBaseLineObj.hasSESimplifiedFeature = mtcReviewObj.SEfeatures

        '28.
        mtcBaseLineObj.hasSEStatusBaseLined = If(mtcReviewObj.Documenttype.ToUpper() = "BASELINED", "Yes", "No")

        '29.
        mtcBaseLineObj.modifiedDate = mtcReviewObj.modifiedDate

        Return mtcBaseLineObj

    End Function

    Private Function GetMTRPart(ByVal mtcReviewObj As DocumentModel,
                                    ByVal mtcMtrModelObj As MTC_MTR_Model, ByVal m2MDataObj As M2MDataModel) As MTR_Part

        Dim mtcPartObj As New MTR_Part()

        mtcPartObj.assemblyName = mtcReviewObj.fileNameWithoutExt

        '1.
        mtcPartObj.isFeatureFullyConstrained = mtcReviewObj.sketchisfullydefined

        '2.
        mtcPartObj.verifySuppressFeature = mtcReviewObj.issupress

        '3.
        mtcPartObj.isAdjustable = mtcReviewObj.isadjustable

        '4.
        If mtcReviewObj.author = Nothing Or mtcReviewObj.Title = Nothing Or mtcReviewObj.materialused = Nothing _
            Or mtcReviewObj.matlspec = Nothing Or mtcReviewObj.projectname = Nothing _
            Or mtcReviewObj.revisionNumber_Prop = Nothing Or mtcReviewObj.documentno = Nothing Or mtcReviewObj.UomProperty = Nothing Or mtcReviewObj.category = Nothing Or mtcReviewObj.comments = Nothing Or mtcReviewObj.keywords = Nothing Then
            mtcPartObj.isValidAllCategories = "No"
        Else
            mtcPartObj.isValidAllCategories = "Yes"
        End If

        '5.
        If mtcReviewObj.hardwarepart Then
            mtcPartObj.verifyFastenerHardwarePart = "Yes"
        Else
            mtcPartObj.verifyFastenerHardwarePart = "No"
        End If

        '6.
        mtcPartObj.verifyWeightMass = mtcReviewObj.density

        '7.
        mtcPartObj.verifyUpdateOnFileSave = String.Empty

        '8.
        mtcPartObj.author = mtcReviewObj.author

        '9.
        mtcPartObj.modifiedDate = mtcReviewObj.modifiedDate

        Return mtcPartObj

    End Function

    '

    Private Function GetMTCElectricalPart(ByVal mtcReviewObj As DocumentModel,
                                    ByVal mtcMtrModelObj As MTC_MTR_Model, ByVal m2MDataObj As M2MDataModel) As MTC_Electrical

        Dim mtcPartObj As New MTC_Electrical()

        mtcPartObj.assemblyName = mtcReviewObj.fileNameWithoutExt

        '1.
        mtcPartObj.ecoNumber = mtcReviewObj.ECO

        mtcPartObj.isPartFound = mtcReviewObj.ispartfound

        '2.
        mtcPartObj.partNumber = If(mtcReviewObj.ispartfound = True, "Yes", "No")

        '3.
        mtcPartObj.revisionLevel = mtcReviewObj.revisionNumber_Prop

        '4.
        mtcPartObj.author = mtcReviewObj.author

        mtcPartObj.projectName = mtcReviewObj.projectname

        '5.
        mtcPartObj.projectNameExist = If(mtcMtrModelObj.projectNameList.Contains(mtcReviewObj.projectname), "Yes", "No")

        '6.
        mtcPartObj.revisionNumberCorrect = If(mtcReviewObj.revisionNumber_Prop = mtcReviewObj.revisionNumber_FileName, "Yes", "No")

        mtcPartObj.documentNumber = mtcReviewObj.documentno

        '7.
        mtcPartObj.documentNumberCorrect = If(mtcReviewObj.documentno = mtcReviewObj.fileNameWithoutExt, "Yes", "No")

        '8.
        mtcPartObj.authorExists = If(mtcMtrModelObj.authorList.Contains(mtcReviewObj.author), "Yes", "No")

        '9.
        If mtcReviewObj.author = Nothing Or mtcReviewObj.Title = Nothing Or mtcReviewObj.materialused = Nothing _
            Or mtcReviewObj.matlspec = Nothing Or mtcReviewObj.projectname = Nothing _
            Or mtcReviewObj.revisionNumber_Prop = Nothing Or mtcReviewObj.documentno = Nothing Or mtcReviewObj.UomProperty = Nothing Or mtcReviewObj.category = Nothing Or mtcReviewObj.comments = Nothing Or mtcReviewObj.keywords = Nothing Then
            mtcPartObj.isDashPopulated = "No"
        Else
            mtcPartObj.isDashPopulated = "Yes"
        End If

        mtcPartObj.title = mtcReviewObj.Title

        '10.
        mtcPartObj.isTitleMatch_ItemMaster = If(m2MDataObj.M2Mdescript.Trim() = mtcReviewObj.Title.Trim(), "Yes", "No")

        '11.
        mtcPartObj.modifiedDate = mtcReviewObj.modifiedDate

        Return mtcPartObj

    End Function

    Private Function GetMTCPart(ByVal mtcReviewObj As DocumentModel,
                                    ByVal mtcMtrModelObj As MTC_MTR_Model, ByVal m2MDataObj As M2MDataModel) As MTC_Part

        Dim mtcPartObj As New MTC_Part()

        mtcPartObj.assemblyName = mtcReviewObj.fileNameWithoutExt

        '1.
        mtcPartObj.ecoNumber = mtcReviewObj.ECO

        mtcPartObj.isPartFound = mtcReviewObj.ispartfound

        '2.
        mtcPartObj.partNumber = If(mtcReviewObj.ispartfound = True, "Yes", "No")

        '3.
        mtcPartObj.revisionLevel = mtcReviewObj.revisionNumber_Prop

        '4.
        mtcPartObj.author = mtcReviewObj.author

        mtcPartObj.projectName = mtcReviewObj.projectname

        '5.
        mtcPartObj.projectNameExist = If(mtcMtrModelObj.projectNameList.Contains(mtcReviewObj.projectname), "Yes", "No")

        '6.
        mtcPartObj.revisionNumberCorrect = If(mtcReviewObj.revisionNumber_Prop = mtcReviewObj.revisionNumber_FileName, "Yes", "No")

        mtcPartObj.documentNumber = mtcReviewObj.documentno

        '7.
        mtcPartObj.documentNumberCorrect = If(mtcReviewObj.documentno = mtcReviewObj.fileNameWithoutExt, "Yes", "No")

        '8.
        mtcPartObj.authorExists = If(mtcMtrModelObj.authorList.Contains(mtcReviewObj.author), "Yes", "No")

        '9.
        If mtcReviewObj.author = Nothing Or mtcReviewObj.Title = Nothing Or mtcReviewObj.materialused = Nothing _
            Or mtcReviewObj.matlspec = Nothing Or mtcReviewObj.projectname = Nothing _
            Or mtcReviewObj.revisionNumber_Prop = Nothing Or mtcReviewObj.documentno = Nothing Or mtcReviewObj.UomProperty = Nothing Or mtcReviewObj.category = Nothing Or mtcReviewObj.comments = Nothing Or mtcReviewObj.keywords = Nothing Then
            mtcPartObj.isDashPopulated = "No"
        Else
            mtcPartObj.isDashPopulated = "Yes"
        End If

        mtcPartObj.title = mtcReviewObj.Title

        '10.
        mtcPartObj.isTitleMatch_ItemMaster = If(m2MDataObj.M2Mdescript.Trim() = mtcReviewObj.Title.Trim(), "Yes", "No")

        mtcPartObj.UomProperty = mtcReviewObj.UomProperty

        '11.
        mtcPartObj.isUOMMatch_M2M = If(mtcReviewObj.UomProperty.Contains(m2MDataObj.M2Mmeasure), "Yes", "No")

        mtcPartObj.materialSpec = mtcReviewObj.matlspec

        '12.
        mtcPartObj.isMaterialSpecExists = If(mtcReviewObj.matlspec = "", "No", "Yes")

        mtcPartObj.materialUsed = mtcReviewObj.materialused

        '13.
        mtcPartObj.isMaterialUsedExists = If(mtcReviewObj.materialused = "", "No", "Yes")

        '14.
        mtcPartObj.isHoleToolUsed = mtcReviewObj.iscutout

        '15.
        mtcPartObj.isSketchFullyConstraint = mtcReviewObj.sketchisfullydefined

        '16.
        mtcPartObj.haveSuppressedFeatureRemoved = mtcReviewObj.issupress 'If(mtcReviewObj.issupress, "Yes", "No")

        '17.
        mtcPartObj.isAdjustable = mtcReviewObj.isadjustable

        '18.
        mtcPartObj.m2mSource = GetM2MSource(m2MDataObj.M2Msource)

        '19.
        mtcPartObj.modifiedDate = mtcReviewObj.modifiedDate

        mtcPartObj.isValidPart = mtcReviewObj.isValidPart

        Return mtcPartObj

    End Function

    Private Function GetM2MSource(ByVal m2mSource As String) As String
        Dim m2mSource2 As String = String.Empty
        'B-BUY, M-MAKE, P-PHANTOM,S-STOCK

        If m2mSource = "S" Then
            m2mSource2 = "STOCK"
        ElseIf m2mSource = "B" Then
            m2mSource2 = "BUY"
        ElseIf m2mSource = "M" Then
            m2mSource2 = "MAKE"
        ElseIf m2mSource = "P" Then
            m2mSource2 = "PHANTOM"
        End If

        Return m2mSource2

    End Function

    Private Function SetPartData(ByVal mtcReviewObj As DocumentModel) As DocumentModel
        Try


            Dim documents As SolidEdgeFramework.Documents = objApp.Documents

            If Not IO.File.Exists(mtcReviewObj.fullpath) Then

                mtcReviewObj.documentLinkBroken = "Yes"
            Else

                Dim doc As SolidEdgeFramework.SolidEdgeDocument = documents.Open(mtcReviewObj.fullpath)
                Dim objPartDocument As SolidEdgePart.PartDocument = Nothing
                Dim isPartDoc As Boolean = False
                Try
                    objPartDocument = DirectCast(documents.Open(mtcReviewObj.fullpath), SolidEdgeFramework.SolidEdgeDocument)
                    isPartDoc = True
                Catch ex As Exception
                    MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
                    CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
                End Try

                If isPartDoc = False Then
                    doc.Close(SaveChanges:=False)
                    mtcReviewObj.isValidPart = False
                    Return mtcReviewObj
                End If

                Threading.Thread.Sleep(2000)

                'Adjustable
                'mtcReviewObj.adjustable = objPartDocument.IsAdjustablePart
                Try
                    mtcReviewObj.adjustable = objPartDocument.IsAdjustablePart
                Catch ex As Exception
                    MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
                    CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
                End Try

                If mtcReviewObj.adjustable Then
                    mtcReviewObj.isadjustable = "Yes"
                Else
                    mtcReviewObj.isadjustable = "No"
                End If

                'Hardware Part
                mtcReviewObj.hardwarepart = objPartDocument.HardwareFile


                'Flatpattern
                Try
                    mtcReviewObj.isflatpattern = "No"
                    Dim flatpatternmodel As SolidEdgePart.FlatPatternModels = objPartDocument.FlatPatternModels
                    If flatpatternmodel.Count = 0 Then
                        mtcReviewObj.isflatpattern = "No"
                    Else
                        mtcReviewObj.isflatpattern = "Yes"
                    End If
                Catch ex As Exception
                    MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
                    CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
                End Try

                Dim models As SolidEdgePart.Models = Nothing
                Dim model As SolidEdgePart.Model = Nothing
                models = objPartDocument.Models

                'check skect is fully defined or not
                mtcReviewObj.sketchisfullydefined = Sketchdefined(objPartDocument)

                'Check Supress

                Try
                    mtcReviewObj.issupress = "No"

                    If objPartDocument.Models.Count > 0 Then

                        Dim features As SolidEdgePart.Features = objPartDocument.Models.Item(1).Features
                        For i = 1 To features.Count
                            Dim obj As Object = features.Item(i)
                            Dim supressvariable As Boolean = obj.suppress
                            If supressvariable = True Then
                                mtcReviewObj.issupress = "Yes"
                                Exit For
                            End If
                        Next

                    End If
                Catch ex As Exception
                    MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
                    CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
                End Try

                'Cutout
                mtcReviewObj.iscutout = "No"
                'mtcReviewObj.iscutout = ""
                Try
                    If models.Count > 0 Then

                        model = models.Item(1)
                        Dim cutouts As SolidEdgePart.ExtrudedCutouts = model.ExtrudedCutouts
                        For Each cutout As SolidEdgePart.ExtrudedCutout In cutouts

                            Try
                                Dim profile As Profile = cutout.Profile
                                If profile.Circles2d.Count = "0" Then
                                    mtcReviewObj.iscutout = "Yes"
                                Else
                                    mtcReviewObj.iscutout = "No"
                                End If
                            Catch ex As Exception
                                MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
                                CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
                            End Try
                        Next

                    End If
                Catch ex As Exception
                    MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
                    CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
                End Try

                'Update on File Save
                objPartDocument.UpdateOnFileSave = False

                Try
                    If objPartDocument.Models.Count > 0 Then

                        'Recompute
                        objPartDocument.Models.Item(1).Recompute()

                        objPartDocument.UpdateOnFileSave = True

                    End If
                Catch ex As Exception
                    MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
                    CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
                End Try

                'check all-interpartlink,copypart,geomtry broken
                mtcReviewObj.allinterpartcopycheck = Interpartcopycheck(objPartDocument, mtcReviewObj)

                'Part Features
                mtcReviewObj.checkPartFeature = CheckPartConstructionFeature(objPartDocument)

                'Interpart link MTR
                mtcReviewObj.interPartLink = "No"

                If mtcReviewObj.isBaseline = True Then

                    mtcReviewObj.isThreadExists = CheckThread(objPartDocument)

                End If

                objPartDocument.Close(SaveChanges:=False)

            End If


        Catch ex As Exception
            CustomLogUtil.Log("While Setting Part Data", ex.Message, ex.StackTrace)
        End Try
        Return mtcReviewObj
    End Function

    Private Function CheckThread(ByVal ObjPartDoc As SolidEdgeFramework.SolidEdgeDocument) As String

        Dim isThreadExist As String = "No"
        Try
            Dim objApp As SolidEdgeFramework.Application = Nothing
            objApp = Marshal.GetActiveObject("SolidEdge.Application")
            ObjPartDoc = objApp.ActiveDocument

            Dim models As SolidEdgePart.Models = Nothing
            Dim model As SolidEdgePart.Model = Nothing
            models = ObjPartDoc.Models
            model = models.Item(1)
            If model IsNot Nothing Then

                If model.Threads.Count > 0 Then
                    isThreadExist = "Yes"
                End If

            End If
        Catch ex As Exception
            MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
            CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
        End Try
        Return isThreadExist
    End Function

    Private Function CheckPartConstructionFeature(ByRef partDoc As PartDocument) As String
        Dim resultCheckPart As String = String.Empty
        Try

            Dim copyConstructionCount As Integer = partDoc.Constructions.CopyConstructions.Count
            If copyConstructionCount > 0 Then
                resultCheckPart = "No"
            Else
                resultCheckPart = "Yes"
            End If
        Catch ex As Exception
            resultCheckPart = ex.Message
            MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
            CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
        End Try
        Return resultCheckPart
    End Function

    Private Function SetSheetMetalData(ByVal mtcReviewObj As DocumentModel) As DocumentModel

        Dim documents As SolidEdgeFramework.Documents = objApp.Documents

        If Not IO.File.Exists(mtcReviewObj.fullpath) Then
            mtcReviewObj.documentLinkBroken = "Yes"
        Else

            Dim doc As SolidEdgeFramework.SolidEdgeDocument = documents.Open(mtcReviewObj.fullpath)
            Dim objSheetMetalDocument As SolidEdgePart.SheetMetalDocument = Nothing
            Dim isSheetMetalDoc As Boolean = False
            Try
                objSheetMetalDocument = DirectCast(documents.Open(mtcReviewObj.fullpath), SolidEdgeFramework.SolidEdgeDocument)
                isSheetMetalDoc = True
            Catch ex As Exception
                MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
                CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
            End Try

            If isSheetMetalDoc = False Then
                doc.Close(SaveChanges:=False)
                mtcReviewObj.isValidPart = False
                Return mtcReviewObj
            End If

            Threading.Thread.Sleep(2000)

            mtcReviewObj.sketchisfullydefined = Sketchdefined(objSheetMetalDocument)

            'Gage excel attached?
            Dim excelfilecount As Integer = Getgagename(objSheetMetalDocument)
            If excelfilecount = 0 Then
                mtcReviewObj.gageeexcelfile = "No"
            ElseIf excelfilecount = 1 Then
                mtcReviewObj.gageeexcelfile = "Yes"
            End If

            'Flat pattern
            Dim flatpatternmodel As SolidEdgePart.FlatPatternModels = objSheetMetalDocument.FlatPatternModels
            If flatpatternmodel.Count = 0 Then
                mtcReviewObj.isflatpattern = "No"
            Else
                mtcReviewObj.isflatpattern = "Yes"
            End If

            Dim models As SolidEdgePart.Models = objSheetMetalDocument.Models
            Dim model As SolidEdgePart.Model = Nothing
            objSheetMetalDocument.UpdateOnFileSave = False

            Try
                objSheetMetalDocument.Models.Item(1).Recompute()
                objSheetMetalDocument.UpdateOnFileSave = True
            Catch ex As Exception
                MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
                CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
            End Try

            'Check supress
            Dim features As SolidEdgePart.Features = objSheetMetalDocument.Models.Item(1).Features
            For i = 1 To features.Count
                Try
                    Dim obj As Object = features.Item(i)
                    Dim supressvariable As Boolean = obj.suppress
                    If supressvariable = True Then
                        mtcReviewObj.issupress = "Yes"
                        Exit For
                    End If
                Catch ex As Exception

                End Try

            Next

            'Cutout
            mtcReviewObj.iscutout = "Yes"

            Try
                model = models.Item(1)
                Try
                    If model.ExtrudedCutouts.Count > 0 Then
                        Dim cutouts As SolidEdgePart.ExtrudedCutouts = model.ExtrudedCutouts
                        For Each cutout As SolidEdgePart.ExtrudedCutout In cutouts
                            Dim profile As Profile = cutout.Profile
                            If profile.Circles2d.Count = "0" Then
                                mtcReviewObj.iscutout = "Yes"
                            Else
                                mtcReviewObj.iscutout = "No"
                            End If
                        Next
                    ElseIf model.NormalCutouts.Count > 0 Then
                        Dim cutouts As SolidEdgePart.NormalCutouts = model.NormalCutouts
                        For Each cutout As SolidEdgePart.NormalCutout In cutouts

                            Dim profile As Profile = cutout.Profile
                            If profile.Circles2d.Count = "0" Then
                                mtcReviewObj.iscutout = "Yes"
                            Else
                                mtcReviewObj.iscutout = "No"
                            End If
                        Next
                    End If
                Catch ex As Exception
                    MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
                    CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
                End Try
            Catch ex As Exception
                MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
                CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
            End Try

            'Adjustable
            Try

                mtcReviewObj.adjustable = objSheetMetalDocument.IsAdjustablePart
            Catch ex As Exception
                Debug.Print("aaaa")
                MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
                CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
            End Try
            If mtcReviewObj.adjustable Then
                mtcReviewObj.isadjustable = "Yes"
            Else
                mtcReviewObj.isadjustable = "No"
            End If

            'Check SEFeatures
            Try
                Dim SEModels As SimplifiedModels = objSheetMetalDocument.SimplifiedModels
                If SEModels.Count > 0 Then
                    mtcReviewObj.SEfeatures = "Yes"
                Else
                    mtcReviewObj.SEfeatures = "No"
                End If
            Catch ex As Exception
                mtcReviewObj.SEfeatures = $"{ex.Message}"
                MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
                CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
            End Try

            'check all-interpartlink,copypart,geomtry broken

            mtcReviewObj.allinterpartcopycheck = Interpartcopycheck(objSheetMetalDocument, mtcReviewObj)

            mtcReviewObj.checkPartFeature = CheckPartConstructionFeature2(objSheetMetalDocument)

            If mtcReviewObj.isBaseline = False Then

                mtcReviewObj.isGeometryBroken = IsGeomtryBroken(objSheetMetalDocument, mtcReviewObj)

                mtcReviewObj.interPartLink = CheckInterPartLinksPSM()

            End If

            'Hole Feature exists
            Try
                If objSheetMetalDocument.Models.Item(1).HoleGeometries.Count > 0 Then ' objSheetMetalDocument.HoleDataCollection.Count > 0 Then
                    mtcReviewObj.holeQty = objSheetMetalDocument.Models.Item(1).HoleGeometries.Count.ToString()
                    mtcReviewObj.isHoleFeatureExists = "True"
                Else
                    mtcReviewObj.holeQty = "0"
                    mtcReviewObj.isHoleFeatureExists = "False"
                End If
            Catch ex As Exception
            End Try

            'Bend Qty
            Try
                mtcReviewObj.bendQty = objSheetMetalDocument.BendTable.BendCount.ToString()
            Catch ex As Exception
            End Try

            'Design Edge bar
            Try
                Dim designEdgeBarFeatureObj As EdgebarFeatures = objSheetMetalDocument.DesignEdgebarFeatures
                Dim flatPatternEdgebarFeaturesObj As EdgebarFeatures = objSheetMetalDocument.FlatPatternEdgebarFeatures
                Dim simplifyEdgebarFeaturesObj As EdgebarFeatures = objSheetMetalDocument.SimplifyEdgebarFeatures

                If designEdgeBarFeatureObj.Count > 0 Then

                    'Louver Feature, Hem feature exists

                    Dim louverCnt As Integer = 0
                    Dim hemCnt As Integer = 0
                    Dim beadCnt As Integer = 0
                    Dim gussetCnt As Integer = 0

                    Dim louverNames As New StringBuilder()
                    Dim hemNames As New StringBuilder()
                    Dim beadNames As New StringBuilder()
                    Dim gussetNames As New StringBuilder()

                    For Each item As Object In designEdgeBarFeatureObj

                        Try
                            Dim louverObj As Louver = DirectCast(item, Louver)
                            Dim louverName As String = louverObj.Name
                            louverNames.AppendLine(louverName)
                            louverCnt += 1
                        Catch ex As Exception
                        End Try

                        Try
                            Dim hemObj As Hem = DirectCast(item, Hem)
                            Dim hemName As String = hemObj.Name
                            hemNames.AppendLine(hemName)
                            hemCnt += 1
                        Catch ex As Exception
                        End Try

                        Try
                            Dim beadObj As Bead = DirectCast(item, Bead)
                            Dim beadName As String = beadObj.Name
                            beadNames.AppendLine(beadName)
                            beadCnt += 1
                        Catch ex As Exception
                        End Try

                        Try
                            Dim gussetObj As Gusset = DirectCast(item, Gusset)
                            Dim gussetName As String = gussetObj.Name
                            gussetNames.AppendLine(gussetName)
                            gussetCnt += 1
                        Catch ex As Exception
                        End Try

                    Next

                    If louverCnt > 0 Then
                        mtcReviewObj.louverExists = "True"
                    End If

                    If hemCnt > 0 Then
                        mtcReviewObj.hemExists = "True"
                    End If

                    If beadCnt > 0 Then
                        mtcReviewObj.beadExists = "True"
                    End If

                    If gussetCnt > 0 Then
                        mtcReviewObj.gussetExists = "True"
                    End If

                    If hemCnt > 0 And beadCnt > 0 And gussetCnt > 0 Then
                        mtcReviewObj.hem_Bead_GussetExists = "True"
                    End If

                End If
            Catch ex As Exception
            End Try

            'Hole Fit

            Try
                Dim holeFitNames As New StringBuilder()
                Dim holeFitCnt As Integer = 0
                Dim objHoleDataCollection As SolidEdgePart.HoleDataCollection = objSheetMetalDocument.HoleDataCollection
                For Each holeDataObj As HoleData In objHoleDataCollection
                    Try
                        Dim holeFit As String = holeDataObj.Fit
                        If holeFit IsNot String.Empty Then
                            holeFitNames.AppendLine($"{holeFit} ,")
                            holeFitCnt += 1
                        End If
                    Catch ex As Exception
                    End Try
                Next
                If holeFitCnt > 0 Then
                    'mtcReviewObj.holeFit = holeFitNames.ToString().Trim()
                    'mtcReviewObj.holeFit = mtcReviewObj.holeFit.Substring(0, mtcReviewObj.holeFit.Length - 1)

                    mtcReviewObj.holeFit = GetHoleFit(holeFitNames)

                End If
            Catch ex As Exception

            End Try

            objSheetMetalDocument.Close(SaveChanges:=False)

        End If

        Return mtcReviewObj

    End Function

    Private Function GetHoleFit(ByVal holeFitNames As StringBuilder) As String
        Dim holeFit As String = "Loose"

        If holeFitNames.ToString().ToUpper().Contains("EXACT") Then
            holeFit = "Exact"
        ElseIf holeFitNames.ToString().ToUpper().Contains("CLOSE") Then
            holeFit = "Close"
        ElseIf holeFitNames.ToString().ToUpper().Contains("NORMAL") Then
            holeFit = "Normal"
        ElseIf holeFitNames.ToString().ToUpper().Contains("NOMINAL") Then
            holeFit = "Nominal"
        ElseIf holeFitNames.ToString().ToUpper().Contains("CLEARANCE") Then
            holeFit = "Clearance"
        ElseIf holeFitNames.ToString().ToUpper().Contains("TRANSITIONAL") Then
            holeFit = "Transitional"
        ElseIf holeFitNames.ToString().ToUpper().Contains("PRESS") Then
            holeFit = "Press"
        Else
            holeFit = "Loose"
        End If
        Return holeFit
    End Function
    Private Function CheckInterPartLinksPSM() As String
        Dim resCheckInterPartLinkBroken As String = "No"
        Try
            Dim sheetMetalDoc As SheetMetalDocument = objApp.ActiveDocument
            Dim resHasInterPartLinks As Boolean
            sheetMetalDoc.HasInterpartLinks(resHasInterPartLinks)
            If resHasInterPartLinks Then
                resCheckInterPartLinkBroken = "No"
            Else
                resCheckInterPartLinkBroken = "Yes"
            End If
        Catch ex As Exception
            MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
        End Try

        Return resCheckInterPartLinkBroken
    End Function

    Private Function IsGeomtryBroken(ByRef objSheetMetalDocument As SolidEdgePart.SheetMetalDocument, ByRef mTCReviewObj As DocumentModel) As String

        Dim geoMetryBroken As String = "Yes"
        Try
            Dim ipl As SolidEdgeFramework.InterpartLinks = objSheetMetalDocument.InterpartLinks
            If ipl.Count > 1 Then
                geoMetryBroken = "No"
            End If
        Catch ex As Exception
            MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
            CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
        End Try
        Return geoMetryBroken

    End Function

    Private Function Sketchdefined(ByVal ObjPartDoc As SolidEdgeFramework.SolidEdgeDocument) As String
        Dim objApp As SolidEdgeFramework.Application = Nothing
        'Dim ObjPartDoc As SolidEdgeFramework.SolidEdgeDocument = Nothing
        Dim objProfiles As SolidEdgePart.Profiles = Nothing
        Dim objProfilesets As SolidEdgePart.ProfileSets = Nothing

        objApp = Marshal.GetActiveObject("SolidEdge.Application")
        ObjPartDoc = objApp.ActiveDocument

        Dim models As SolidEdgePart.Models = Nothing
        Dim model As SolidEdgePart.Model = Nothing
        models = ObjPartDoc.Models
        model = models.Item(1)
        objProfilesets = ObjPartDoc.ProfileSets
        Dim sketchfullydefined As String = "Yes"
        For Each profileset As ProfileSet In objProfilesets

            Dim isunderdefined As Boolean = profileset.IsUnderDefined

            If isunderdefined = True Then
                sketchfullydefined = "No"
                Return sketchfullydefined
            End If

        Next

        Return sketchfullydefined
    End Function

    Private Function Getgagename(ByVal objSMDoc As SolidEdgePart.SheetMetalDocument) As Integer
        '   Dim objApplication As SolidEdgeFramework.Application = Nothing

        Dim myMatTable As SolidEdgeFramework.MatTable = Nothing
        Dim strCurrGageName As String = ""
        Dim strGageFilePath As String = ""
        Dim nMTLUsingExcel As Integer
        Dim strMTLGageTableName As String = ""
        Dim nDocUsingExcel As Integer = 0
        Dim strDocGageTableName As String = ""
        Dim nCountBendRadiusVals As Integer
        Dim nCountBendAngleVals As Integer
        Dim nCountNFVals As Integer

        Try

            myMatTable = objApp.GetMaterialTable()

            If (objSMDoc Is Nothing) Then
                MessageBox.Show("Failed to get Sheet Metal Document object.", "Message")
            End If
            Call myMatTable.GetPSMGaugeInfoForDoc(objSMDoc,
                                      strCurrGageName,
                                      strGageFilePath,
                                      nMTLUsingExcel,
                                      strMTLGageTableName,
                                      nDocUsingExcel,
                                      strDocGageTableName,
                                      nCountBendRadiusVals,
                                      nCountBendAngleVals,
                                      nCountNFVals)
        Catch ex As Exception
            MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
            CustomLogUtil.Log("While Fetching the Gage Name", ex.Message, ex.StackTrace)
        End Try
        Return nDocUsingExcel

    End Function

    Private Function Interpartcopycheck(ByVal partDoc As SolidEdgeFramework.SolidEdgeDocument, ByRef mTCReviewObj As DocumentModel) As String
        Dim CheckPart As String = String.Empty
        Dim interpartcheck As String = String.Empty

        Dim result As String = String.Empty
        Try

            'part copies detected
            Dim copyConstructionCount As Integer = partDoc.Constructions.CopyConstructions.Count
            If copyConstructionCount > 0 Then
                CheckPart = "Yes"
                mTCReviewObj.partCopiesDetected = "Yes"
            Else
                CheckPart = "No"
                mTCReviewObj.partCopiesDetected = "No"
            End If

            'Inter part copies detected
            Dim interpartlinkscount As Integer = partDoc.InterpartLinks.count
            If interpartlinkscount > 0 Then
                interpartcheck = "Yes"
                mTCReviewObj.interPartCopiesDetected = "Yes"
            Else
                interpartcheck = "No"
                mTCReviewObj.interPartCopiesDetected = "No"
            End If

            If CheckPart = "Yes" And interpartcheck = "Yes" Then
                result = "Yes"
            Else
                result = "N0"
            End If
        Catch ex As Exception
            MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
            CustomLogUtil.Log("While Fetching the Interpart Copy Check")
        End Try
        Return result
    End Function

    Private Function CheckPartConstructionFeature2(ByRef partDoc As SheetMetalDocument) As String
        Dim resultCheckPart As String = String.Empty

        Try

            Dim copyConstructionCount As Integer = partDoc.Constructions.CopyConstructions.Count
            If copyConstructionCount > 0 Then
                resultCheckPart = "No"
            Else
                resultCheckPart = "Yes"
            End If
        Catch ex As Exception
            resultCheckPart = ex.Message
            MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
            CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
        End Try
        Return resultCheckPart
    End Function

    Public Function SetM2mData(ByVal m2MDataObj As M2MDataModel, ByRef dv As DataView) As M2MDataModel

        For Each drv As DataRowView In dv

            m2MDataObj.M2Mpartno = If(drv(ExcelUtil.ExcelMtcReview.fpartno.ToString) = Nothing, "", drv(ExcelUtil.ExcelMtcReview.fpartno.ToString))



            m2MDataObj.M2Mdescript = If(drv(ExcelUtil.ExcelMtcReview.fdescript.ToString) = Nothing, "", drv(ExcelUtil.ExcelMtcReview.fdescript.ToString))

            m2MDataObj.M2Mmeasure = If(drv(ExcelUtil.ExcelMtcReview.fmeasure.ToString) = Nothing, "", drv(ExcelUtil.ExcelMtcReview.fmeasure.ToString))

            m2MDataObj.M2Msource = If(drv(ExcelUtil.ExcelMtcReview.fsource.ToString) = Nothing, "", drv(ExcelUtil.ExcelMtcReview.fsource.ToString))

            'm2MDataObj.M2Mvendorname = If(drv(ExcelUtil.ExcelMtcReview.VendorName.ToString) = Nothing, "", drv(ExcelUtil.ExcelMtcReview.VendorName.ToString))
            m2MDataObj.M2Mvendorname = ""

            m2MDataObj.M2Mflocation = If(drv(ExcelUtil.ExcelMtcReview.flocation.ToString) = Nothing, "", drv(ExcelUtil.ExcelMtcReview.flocation.ToString))

            'm2MDataObj.M2MFbin = If(drv(ExcelUtil.ExcelMtcReview.Fbin.ToString) = Nothing, "", drv(ExcelUtil.ExcelMtcReview.Fbin.ToString))
            m2MDataObj.M2MFbin = If(Convert.IsDBNull(drv(ExcelUtil.ExcelMtcReview.Fbin.ToString)), "", drv(ExcelUtil.ExcelMtcReview.Fbin.ToString))
        Next

        m2MDataObj.commentformate = m2MDataObj.M2Mvendorname + " = " + m2MDataObj.M2Mpartno

        Return m2MDataObj
    End Function

    'Private Function GetRevisionLevel(ByVal fileName As String) As String

    '    Dim revisionLevel As String = "Invalid"

    '    If fileName.Length > 3 Then
    '        revisionLevel = fileName.Substring(fileName.Length - 3)
    '        If revisionLevel.Contains("-") Then
    '            Dim myDelims As String() = New String() {"-"}
    '            Dim revisionLevelSplit = revisionLevel.Split(myDelims, StringSplitOptions.None)
    '            If revisionLevelSplit.Length > 1 Then
    '                revisionLevel = revisionLevelSplit(revisionLevelSplit.Length - 1)
    '            End If
    '        End If
    '    End If

    '    Return revisionLevel

    'End Function

    '13th Sep 2024
    Private Function GetRevisionLevel(ByVal fileName As String, ByVal revisionNumber_Prop As String) As String

        Dim revisionLevel As String = "Invalid"

        If fileName.Length > 3 Then
            revisionLevel = fileName.Substring(fileName.Length - 3)
            If revisionLevel.Contains("-") Then
                Dim myDelims As String() = New String() {"-"}
                Dim revisionLevelSplit = revisionLevel.Split(myDelims, StringSplitOptions.None)
                If revisionLevelSplit.Length > 1 Then
                    revisionLevel = revisionLevelSplit(revisionLevelSplit.Length - 1)
                End If
            ElseIf revisionNumber_Prop = "0" Then
                revisionLevel = "0"
            End If
        End If

        Return revisionLevel

    End Function

    ' If(dicProperties.ContainsKey("Revision"), dicProperties("Revision"), "")

    Private Function SetVariablesData(ByRef dr As DataRow, ByVal mtcReviewObj As DocumentModel, ByVal excelName As String) As DocumentModel

#Region "Set Variables Data"

        If dr("Item Number") IsNot Nothing Then
            mtcReviewObj.itemnumber = dr("Item Number").ToString()
        End If


        If dr("Revision Number") IsNot Nothing Then
            mtcReviewObj.revisionNumber_Prop = dr("Revision Number").ToString()
        End If

        If dr("File Name (no extension)") IsNot Nothing Then
            mtcReviewObj.fileNameWithoutExt = dr("File Name (no extension)").ToString()
            '12th Sep 2024
            If mtcReviewObj.fileNameWithoutExt.Contains("_") Or mtcReviewObj.fileNameWithoutExt.Contains("-") Then
                mtcReviewObj.revisionNumber_FileName = GetRevisionLevel(mtcReviewObj.fileNameWithoutExt, mtcReviewObj.revisionNumber_Prop)
            Else
                mtcReviewObj.fileName = mtcReviewObj.fileNameWithoutExt
            End If

        End If


        If dr("Title") IsNot Nothing Then

            'mtcReviewObj.Title = dr("Title").ToString()
            Dim TitleValue As String = Strings.Left(dr("Title").ToString(), 35)
            mtcReviewObj.Title = TitleValue
            'mtcReviewObj.materialDescription = dr("Title").ToString()
        End If

        If dr("Author") IsNot Nothing Then
#Region "DGSAuthor"
            'If excelName = "DGS" Then
            '    mtcReviewObj.author = "DGS"
            'Else
            '    mtcReviewObj.author = dr("Author").ToString()
            'End If
#End Region
            mtcReviewObj.author = dr("Author").ToString()

        End If

        If dr("Document Number") IsNot Nothing Then
            mtcReviewObj.documentno = dr("Document Number").ToString()
        End If

        If dr("Comments") IsNot Nothing Then
            mtcReviewObj.comments = dr("Comments").ToString()
        End If

        If dr("Category") IsNot Nothing Then
            mtcReviewObj.category = dr("Category").ToString()
        End If

        If mtcReviewObj.category.ToUpper().Trim() = "ELECTRICAL" Then
            mtcReviewObj.isElectrical = True
        End If

        If dr("Material Used") IsNot Nothing Then
            mtcReviewObj.materialused = dr("Material Used").ToString()
        End If

        If dr("MATL SPEC") IsNot Nothing Then
            mtcReviewObj.matlspec = dr("MATL SPEC").ToString()
        End If

        If mtcReviewObj.matlspec.ToUpper() = "PURCHASED" Or mtcReviewObj.materialused.ToUpper() = "PURCHASED" Then
            mtcReviewObj.isBaseline = True
        End If

        If dr("File Name (full path)") IsNot Nothing Then
            mtcReviewObj.fullpath = dr("File Name (full path)").ToString()
            'mtcReviewObj.filePath = dr("File Name (full path)").ToString()
        End If

        If dr("Last Author") IsNot Nothing Then
            mtcReviewObj.lastsaved = dr("Last Author").ToString()
        End If

        If dr("Density") IsNot Nothing Then
            mtcReviewObj.density = dr("Density").ToString()
        End If

        If dr("Project") IsNot Nothing Then
            mtcReviewObj.projectname = dr("Project").ToString()
        End If

        If mtcReviewObj.projectname = "BROOKVILLE EQUIPMENT CORP" Then
            mtcReviewObj.isBrookVilleProject_Baseline = "Yes"
        End If

        If dr("Status Text") IsNot Nothing Then
            mtcReviewObj.Documenttype = dr("Status Text").ToString()
        End If

        '26th september 2024
        'If dr("UOM") IsNot Nothing Then
        '    mtcReviewObj.UomProperty = dr("UOM").ToString()
        'End If
        If dr("UOM") IsNot Nothing Then
            If Not dr("UOM").ToString() = "" Then
                mtcReviewObj.UomProperty = dr("UOM").ToString()
            End If
        End If

        If dr("Keywords") IsNot Nothing Then
            mtcReviewObj.keywords = dr("Keywords").ToString()
        End If

        If dr("ECO/SOW") IsNot Nothing Then
            mtcReviewObj.ECO = dr("ECO/SOW").ToString()
        End If

        If dr("Modified") IsNot Nothing Then
            mtcReviewObj.modifiedDate = dr("Modified").ToString()
        End If

        '====

        If dr("Material") IsNot Nothing Then
            mtcReviewObj.material = dr("Material").ToString()
        End If

        If dr("Material Thickness") IsNot Nothing Then
            mtcReviewObj.materialThickness = dr("Material Thickness").ToString()
        End If

        If dr("Bend Radius") IsNot Nothing Then
            mtcReviewObj.bendRadius = dr("Bend Radius").ToString()
        End If

        If dr("Flat_Pattern_Model_CutSizeX") IsNot Nothing Then
            mtcReviewObj.flat_Pattern_Model_CutSizeX = dr("Flat_Pattern_Model_CutSizeX").ToString()
        End If

        If dr("Flat_Pattern_Model_CutSizeY") IsNot Nothing Then
            mtcReviewObj.flat_Pattern_Model_CutSizeY = dr("Flat_Pattern_Model_CutSizeY").ToString()
        End If

        If dr("Mass (Item)") IsNot Nothing Then
            mtcReviewObj.massItem = dr("Mass (Item)").ToString()
        End If

        If dr("QAQC") IsNot Nothing Then
            mtcReviewObj.qAQC = dr("QAQC").ToString()
        End If

        If dr("Quantity") IsNot Nothing Then
            mtcReviewObj.quantity = dr("Quantity").ToString()
        End If

#End Region

        Return mtcReviewObj
    End Function

    Private Function IsValidBaseLineDirectoryPath(ByVal baselineDirPath As String, ByVal docFullPath As String)
        Dim resValidDir As Boolean = False
        Try
            Dim docDirPath As String = IO.Path.GetDirectoryName(docFullPath)

            If docDirPath.Contains(baselineDirPath) Then
                resValidDir = True
            End If
        Catch ex As Exception
            MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
            CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
        End Try
        Return resValidDir
    End Function

    Private Function FilterCurrentAssemblyData(ByVal mtcMtrModelObj As MTC_MTR_Model) As System.Data.DataTable

        Dim dtFinalData As New Data.DataTable("FinalData")

        Dim dt As System.Data.DataTable = mtcMtrModelObj.dtCurrentAssemblyData.Copy()
        dt.Rows.Clear()
        Dim dvFinal As New DataView(dt)

        Dim dvAssembly As DataView = FilterCurrentAssemblyData_ExtensionWise(mtcMtrModelObj, ".asm")
        Dim dvPart As DataView = FilterCurrentAssemblyData_ExtensionWise(mtcMtrModelObj, ".par")
        Dim dvSheetMetal As DataView = FilterCurrentAssemblyData_ExtensionWise(mtcMtrModelObj, ".psm")

        If chkAssembly.Checked Then
            dvFinal.Table.Merge(dvAssembly.ToTable)
        End If
        If chkPart.Checked Then
            dvFinal.Table.Merge(dvPart.ToTable)
        End If
        If chkSheetMetal.Checked Then
            dvFinal.Table.Merge(dvSheetMetal.ToTable)
        End If

        dtFinalData = dvFinal.ToTable()

        Return dtFinalData
    End Function

    '17th Sep 2024
    Private Function FilterCurrentPartData(ByVal mtcMtrModelObj As MTC_MTR_Model) As System.Data.DataTable

        Dim dtFinalData As New Data.DataTable("FinalData")

        Dim dt As System.Data.DataTable = mtcMtrModelObj.dtCurrentPartData.Copy()
        dt.Rows.Clear()
        Dim dvFinal As New DataView(dt)

        Dim dvAssembly As DataView = FilterCurrentAssemblyData_ExtensionWise(mtcMtrModelObj, ".asm")
        Dim dvPart As DataView = FilterCurrentAssemblyData_ExtensionWise(mtcMtrModelObj, ".par")
        Dim dvSheetMetal As DataView = FilterCurrentAssemblyData_ExtensionWise(mtcMtrModelObj, ".psm")

        If chkAssembly.Checked Then
            dvFinal.Table.Merge(dvAssembly.ToTable)
        End If
        If chkPart.Checked Then
            dvFinal.Table.Merge(dvPart.ToTable)
        End If
        If chkSheetMetal.Checked Then
            dvFinal.Table.Merge(dvSheetMetal.ToTable)
        End If

        dtFinalData = dvFinal.ToTable()

        Return dtFinalData
    End Function


    '17th Sep 2024
    Private Function FilterCurrentSheetMetalData(ByVal mtcMtrModelObj As MTC_MTR_Model) As System.Data.DataTable

        Dim dtFinalData As New Data.DataTable("FinalData")

        Dim dt As System.Data.DataTable = mtcMtrModelObj.dtCurrentSheetMetalData.Copy()
        dt.Rows.Clear()
        Dim dvFinal As New DataView(dt)

        Dim dvAssembly As DataView = FilterCurrentAssemblyData_ExtensionWise(mtcMtrModelObj, ".asm")
        Dim dvPart As DataView = FilterCurrentAssemblyData_ExtensionWise(mtcMtrModelObj, ".par")
        Dim dvSheetMetal As DataView = FilterCurrentAssemblyData_ExtensionWise(mtcMtrModelObj, ".psm")

        If chkAssembly.Checked Then
            dvFinal.Table.Merge(dvAssembly.ToTable)
        End If
        If chkPart.Checked Then
            dvFinal.Table.Merge(dvPart.ToTable)
        End If
        If chkSheetMetal.Checked Then
            dvFinal.Table.Merge(dvSheetMetal.ToTable)
        End If

        dtFinalData = dvFinal.ToTable()

        Return dtFinalData
    End Function

    Private Function FilterCurrentAssemblyData_ExtensionWise(ByVal mtcMtrModelObj As MTC_MTR_Model, ByVal ext As String) As DataView
        Dim dt As System.Data.DataTable = mtcMtrModelObj.dtCurrentAssemblyData.Copy() ' dtAssemblyData.Copy()
        Dim DV As New DataView(dt)
        Try

            DV.RowFilter = "[File Name (full path)] LIKE '%" + ext + "%' or [File Name (full path)] LIKE '%" + ext.ToUpper() + "%'"
        Catch ex As Exception
            MTC_MTR_ReviewForm2.log.Error($"****{ex.Message}{vbNewLine}{ex.StackTrace}")
            CustomLogUtil.Log("****", ex.Message, ex.StackTrace)
        End Try
        'dgvDocumentDetails.DataSource = DV.ToTable()
        'dtfilter = DV.ToTable
        Return DV
    End Function

    Private Sub BtnBrowseExportDir_Click(sender As Object, e As EventArgs) Handles BtnBrowseExportDirMTR.Click
        Try
            Dim folderpath As String = ""
            folderpath = BrowseFolderAdvanced()
            If Not folderpath = String.Empty Then
                txtExportDirLocationMTR.Text = folderpath
            End If
        Catch ex1 As Exception
            Try
                Dim MyFolderBrowser As New Ookii.Dialogs.VistaFolderBrowserDialog
                Dim dlgResult As DialogResult = MyFolderBrowser.ShowDialog
                If dlgResult = DialogResult.OK Then
                    txtExportDirLocationMTR.Text = MyFolderBrowser.SelectedPath
                End If
            Catch ex As Exception
                If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
                    Dim path As String = FolderBrowserDialog1.SelectedPath
                    txtExportDirLocationMTR.Text = path
                End If
            End Try

        End Try
    End Sub



    Private Sub MTC_MTR_ReviewForm2_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'KillSolidEdgeProcess.killSilent()
        If objApp.Visible = False Then
            objApp.Quit()
        End If
    End Sub
End Class