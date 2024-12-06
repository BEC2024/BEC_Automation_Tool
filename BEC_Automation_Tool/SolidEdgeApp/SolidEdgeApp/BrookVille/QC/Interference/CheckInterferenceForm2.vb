Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports SolidEdgeAssembly
'Imports SolidEdge.Framework.Interop
Imports SolidEdgeFramework
Imports WK.Libraries.BetterFolderBrowserNS

Public Class CheckInterferenceForm2
    Dim dictMaterials As New Dictionary(Of String, List(Of String))()
    Dim objApp As SolidEdgeFramework.Application = Nothing
    Dim objDocument As SolidEdgeFramework.SolidEdgeDocument = Nothing
    Dim objSheetMetalDocument As SolidEdgePart.SheetMetalDocument = Nothing
    Dim objAssemblyDocument As SolidEdgeAssembly.AssemblyDocument
    Dim dtInteferenceMaterials As New DataTable()
    Dim myMatTable As SolidEdgeFramework.MatTable = Nothing
    Dim cnt As Integer = 1
    Dim skip As Boolean = True

    Private Sub BtnBrowseExcel_Click(sender As Object, e As EventArgs) Handles btnBrowseExcel.Click
        Try
            Using dialog As New OpenFileDialog
                dialog.Filter = "Excel files (*.xlsx)|*.xlsx"
                If dialog.ShowDialog() <> DialogResult.OK Then Return
                txtExcelPath.Text = dialog.FileName
            End Using

            If IO.File.Exists(txtExcelPath.Text) Then

                Dim dictIntereference As New Dictionary(Of String, DataSet)()
                Dim filepath As String = txtExcelPath.Text ' $"{System.IO.Directory.GetCurrentDirectory}\Bom_20220503-1.xlsx"
                dictIntereference = ExcelUtil.ReadRawMaterials2Interference(filepath)
                Dim ds As DataSet = dictIntereference.Values(0)
                dtInteferenceMaterials = ds.Tables(0)
                dgvDocumentDetails.DataSource = dtInteferenceMaterials
                'FillMaterials(dt)
                FillSearchCombo()
            Else
                MessageBox.Show("Please select material excel details.", "Message")
            End If
        Catch ex As Exception
        End Try

    End Sub

    Private Sub FillSearchCombo()
        ComboBoxFields.Items.Clear()
        ComboBoxFields.Items.Add("Select")

        'For Each columns As System.Windows.Forms.DataGridViewColumn In dgvDocumentDetails.Columns
        '    ComboBoxFields.Items.Add(columns.Name.ToString())
        'Next
        ComboBoxFields.Items.Add(ExcelUtil.ExcelSheetColumnsIntereference.Material.ToString())
        ComboBoxFields.Items.Add(ExcelUtil.ExcelSheetColumnsIntereference.Type.ToString())
        If ComboBoxFields.Items.Count > 0 Then
            ComboBoxFields.SelectedItem = ComboBoxFields.Items(0) 'ComboBoxFields.Items.Count - 1)
        End If
    End Sub

    Private Function IsMaterialExists(ByVal materialName As String) As Boolean

        'Dim isMaterialAvai As Boolean = True
        'For Each item In CheckedListBox1.CheckedItems

        '    If materialName = item Then
        '        isMaterialAvai = False
        '        Exit For
        '    End If

        'Next
        'Return isMaterialAvai
        Return False
    End Function

    Public Function GetAlloccur(ByVal occs As SolidEdgeAssembly.Occurrences, ByVal highset As SolidEdgeFramework.HighlightSet) As SolidEdgeFramework.HighlightSet

        Dim oCustom As SolidEdgeFramework.Properties
        Dim oProp As SolidEdgeFramework.Property

        For Each occ As SolidEdgeAssembly.Occurrence In occs

            Dim obj As ObjectType = occ.Type

            If obj = ObjectType.igPart Then
                Try

                    oCustom = occ.PartDocument.Properties("Custom")

                    oProp = oCustom.Item("Material Used")
                    Dim a As String = oProp.Value
                    If Not IsMaterialExists(a) Then ' a = "PL14GAA606" Then

                        highset.AddItem(occ)

                    End If
                Catch ex As Exception
                End Try
            End If
            If obj = ObjectType.igSubAssembly Then
                Try
                    Dim assembldocument As AssemblyDocument = occ.OccurrenceDocument
                    GetAlloccur(assembldocument.Occurrences, highset)
                Catch ex As Exception
                End Try
            End If

        Next

        Return highset
    End Function

    Private Sub FillMaterials(ByVal dt As DataTable)

        'For Each dr As DataRow In dt.Rows

        '    Dim materialsUsed As String = dr("Material")
        '    If Not CheckedListBox1.Items.Contains(materialsUsed) Then
        '        CheckedListBox1.Items.Add(materialsUsed)
        '    End If
        'Next

    End Sub

    Dim occurancelist As New List(Of String)
    Private Function IsValid() As Boolean

        SetSolidEdgeInstance()
        DisableBtn
        If objApp Is Nothing Then
            Return False
        Else
            Return True
        End If

    End Function
    Public Sub DisableBtn()
        If objApp Is Nothing Then
            btnInterferenceExcludeMaterial.Enabled = False
            btnBrowseExcel.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = False
            Button1.Enabled = False
            btnGenerateInterferenceReport.Enabled = False
            btnSearchFile.Enabled = False
        Else
            btnInterferenceExcludeMaterial.Enabled = True
            btnBrowseExcel.Enabled = True
            Button2.Enabled = True
            Button3.Enabled = True
            Button1.Enabled = True
            btnGenerateInterferenceReport.Enabled = True
            btnSearchFile.Enabled = True
        End If

    End Sub
    Private Sub CheckIntereferenceForm2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CustomLogUtil.Heading("Interference Form Open.....")
        If IsValid() Then

            SetControls(dgvDocumentDetails)

            myMatTable = objApp.GetMaterialTable()

            txtExcelPath.Text = Config.configObj.interferenceExcludeMaterialExcelPath

            Me.Text = $"{Me.Text} ({GlobalEntity.Version})"
        Else
            MessageBox.Show("Please Open Solid-Edge Assembly and Restart the Application", "Message")
            CustomLogUtil.Log("Please Open Solid-Edge Assembly and Restart the Application", "", "")
        End If
    End Sub

    Private Sub SetSolidEdgeInstance()
        Try
            objApp = Marshal.GetActiveObject("SolidEdge.Application")
        Catch ex As Exception
            'MessageBox.Show($"Error in fetching the Solid-Edge instance {ex.Message} {vbNewLine} {ex.StackTrace}", "Message")
            'CustomLogUtil.Log("While fetching the Solid-Edge instance..", ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub SetControls(ByRef DataGridViewComments As DataGridView)
        Try
            For Each col As DataGridViewColumn In DataGridViewComments.Columns
                col.SortMode = DataGridViewColumnSortMode.NotSortable
            Next

            DataGridViewComments.AlternatingRowsDefaultCellStyle.BackColor = Drawing.Color.GhostWhite
            DataGridViewComments.AlternatingRowsDefaultCellStyle.ForeColor = Drawing.Color.DarkBlue

            DataGridViewComments.DefaultCellStyle.SelectionBackColor = Drawing.Color.FromArgb(6, 150, 215) ' Drawing.Color.WhiteSmoke
            DataGridViewComments.DefaultCellStyle.SelectionForeColor = Drawing.Color.WhiteSmoke 'Drawing.Color.OrangeRed 'FromArgb(6, 150, 215) '

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

        Catch ex As Exception
            MessageBox.Show($"Error Seting the Controls ", "Message")
            CustomLogUtil.Log("While Seting the Controls..", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Function GetSelectedMaterials() As List(Of String)
        Dim selectedMaterialList As New List(Of String)()
        Try
            Dim dt As DataTable = dtInteferenceMaterials
            Dim DVSelected As New DataView(dt, "Select=True", "", DataViewRowState.CurrentRows)
            Dim dtSelected As DataTable = DVSelected.ToTable()

            For Each dr As DataRow In dtSelected.Rows
                If Not selectedMaterialList.Contains(dr(ExcelUtil.ExcelSheetColumnsIntereference.Type.ToString())) Then
                    Dim materialName As String = dr(ExcelUtil.ExcelSheetColumnsIntereference.Type.ToString()).ToString().Trim()
                    selectedMaterialList.Add(materialName) '(dr(ExcelUtil.ExcelSheetColumnsIntereference.Type.ToString()))
                End If
            Next


        Catch ex As Exception
            MessageBox.Show($"Error While Getting Selected Materials ", "Error")
            CustomLogUtil.Log("While Getting Selected Materials", ex.Message, ex.StackTrace)
        End Try
        Return selectedMaterialList
    End Function

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

    Private Sub BtnSelectParts_Click(sender As Object, e As EventArgs) Handles btnGenerateInterferenceReport.Click
        CustomLogUtil.Heading("Generate Interference Report Starts")
        Try
            WaitStartSave()
            CustomLogUtil.Log("GetSelectedMaterials")
            Dim selectedMaterialList As List(Of String) = GetSelectedMaterials()
            CustomLogUtil.Log("UpdateOccurencesReferenceAnalysisProperty")
            Dim str As String = UpdateOccurencesReferenceAnalysisProperty(selectedMaterialList)
            CustomLogUtil.Log("GenerateInterferenceReport")
            GenerateInterferenceReport()
            CustomLogUtil.Log("UpdateOccurrencePropertyReport")
            UpdateOccurrencePropertyReport(str)

            WaitEndSave()
            CustomLogUtil.Heading("Generate Interference Report Process Completed")
        Catch ex As Exception
            MessageBox.Show($"Error While Clicking On SelectParts Button", "Error")
            CustomLogUtil.Log("While Clicking On SelectParts Button", ex.Message, ex.StackTrace)
        End Try
    End Sub
    Public Shared Function browseFolderAdvanced() As String

        Dim folderpath As String = ""
        Try
            Dim BetterFolderBrowser As New BetterFolderBrowser()

            BetterFolderBrowser.Title = "Select folders"

            BetterFolderBrowser.RootFolder = "C:\\"

            BetterFolderBrowser.Multiselect = False
            If BetterFolderBrowser.ShowDialog() = DialogResult.OK Then
                folderpath = BetterFolderBrowser.SelectedFolder
            End If
        Catch ex As Exception
            MessageBox.Show($"Advanced Browse folder", "Error")
            CustomLogUtil.Log("Advanvce Browse Folder", ex.Message, ex.StackTrace)
        End Try


        Return folderpath
    End Function
    Private Function OutputDirPath() As String
        Dim folderpath As String = ""
        Try

            folderpath = browseFolderAdvanced()
            If Not folderpath = String.Empty Then

            End If
        Catch ex1 As Exception
            Try
                Dim MyFolderBrowser As Ookii.Dialogs.VistaFolderBrowserDialog = New Ookii.Dialogs.VistaFolderBrowserDialog
                Dim dlgResult As DialogResult = MyFolderBrowser.ShowDialog
                If dlgResult = DialogResult.OK Then
                    folderpath = MyFolderBrowser.SelectedPath
                End If
            Catch ex As Exception
                If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
                    folderpath = FolderBrowserDialog1.SelectedPath

                End If
            End Try

        End Try
        Return folderpath
    End Function
    Private Sub UpdateOccurrencePropertyReport(ByVal interferenceDetails As String)
        Try
            Dim dirPath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)

            Dim assemblyDocument1 As SolidEdgeAssembly.AssemblyDocument = objApp.ActiveDocument

            Dim docPath As String = assemblyDocument1.FullName

            Dim docName As String = IO.Path.GetFileNameWithoutExtension(docPath)

            Dim reportName As String = $"UpdateOccurencePropertyDetails_{IO.Path.GetFileNameWithoutExtension(docName)}_{ Date.Now.Day.ToString + "-" + Date.Now.Month.ToString + "-" + Date.Now.Year.ToString}.txt"
            Dim folderPath As String = OutputDirPath()

            'Dim reportPath As String = IO.Path.Combine(dirPath, reportName)
            Dim reportPath As String = IO.Path.Combine(folderPath, reportName)

            reportPath = New Uri(reportPath).LocalPath

            Try

                My.Computer.FileSystem.WriteAllText(reportPath, interferenceDetails, True)
            Catch ex As Exception
            End Try

            Try
                Process.Start(reportPath)
            Catch ex As Exception

            End Try

            Exit Sub
        Catch ex As Exception
            MessageBox.Show($"Error on Update Occurrence Property Report", "Error")
            CustomLogUtil.Log("on Update Occurrence Property Report ", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub GenerateInterferenceReport()
        Try
            Dim nComparisonMethod As Integer
            Dim nSet1 As Integer
            Dim nSet2 As Integer
            Dim objOccurrences As SolidEdgeAssembly.Occurrences = Nothing
            Dim objOccurrence As SolidEdgeAssembly.Occurrence = Nothing
            Dim objInterfOcc As SolidEdgeAssembly.Occurrence = Nothing
            Dim a_objSet1() As Object
            Dim objTemp As Object
            Dim nNumInterferences As Long
            Dim nStatus As SolidEdgeAssembly.InterferenceStatusConstants

            objApp.DisplayAlerts = False
            objAssemblyDocument = objApp.ActiveDocument
            nComparisonMethod = SolidEdgeConstants.InterferenceComparisonConstants.seInterferenceComparisonSet1vsItself
            nSet1 = 0
            nSet2 = 0
            objOccurrences = objAssemblyDocument.Occurrences
            ReDim a_objSet1(objOccurrences.Count)
            For nIndex = 1 To objOccurrences.Count
                objTemp = objOccurrences.Item(nIndex)
                a_objSet1(nSet1) = objTemp
                nSet1 += 1
            Next nIndex

            'Add intereference as part in assembly
            Call objAssemblyDocument.CheckInterference2(
                                        NumElementsSet1:=nSet1,
                                        Set1:=a_objSet1,
                                        Status:=nStatus,
                                        ComparisonMethod:=nComparisonMethod,
                                        AddInterferenceAsOccurrence:=True,
                                        NumInterferences:=nNumInterferences,
                                        InterferenceOccurrence:=objInterfOcc,
                                        IgnoreNonThreadVsThreadConstant:=SolidEdgeConstants.InterferenceOptionsConstants.seIntfOptIgnoreThreadVsNonThreaded,
                                        ReportFilename:="report.txt") ', ReportType:="TEXT") ',


        Catch ex As Exception
            MessageBox.Show($"Error In generate interference", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            CustomLogUtil.Log("Error In generate interference", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Function GetOccurencePropertyMaterialName(ByVal occurrence As Occurrence) As String
        Dim oCustom As Properties
        Dim oProp As [Property]
        oCustom = occurrence.PartDocument.Properties("Custom")

        oProp = oCustom.Item("Material Used")
        Dim material As String = oProp.Value
        Return material
    End Function

    Private Function GetOccurenceDocumentMaterialName(ByVal occurrence As Occurrence) As String

        Dim materialName As String = String.Empty
        Try
            myMatTable.GetCurrentMaterialName(occurrence.PartDocument, materialName)
        Catch ex As Exception
            materialName = $"{ex.Message}{vbTab}{ex.StackTrace} "
        End Try

        Return materialName
    End Function

    Public Function UpdateOccurencesReferenceAnalysisProperty(ByVal selectedMaterialList As List(Of String)) As String

        Dim application As SolidEdgeFramework.Application = objApp
        Dim assemblyDocument As SolidEdgeAssembly.AssemblyDocument = Nothing
        Dim occurrences As SolidEdgeAssembly.Occurrences = Nothing
        Dim interferenceStatus As SolidEdgeAssembly.InterferenceStatusConstants = Nothing
        Dim compare As SolidEdgeConstants.InterferenceComparisonConstants = SolidEdgeConstants.InterferenceComparisonConstants.seInterferenceComparisonSet1vsAllOther
        Dim reportType As SolidEdgeConstants.InterferenceReportConstants = SolidEdgeConstants.InterferenceReportConstants.seInterferenceReportPartNames

        Dim sbUpdateOccReferenceAnalysisProperty As New StringBuilder()
        sbUpdateOccReferenceAnalysisProperty.AppendLine($"********************{System.DateTime.Now}********************")
        sbUpdateOccReferenceAnalysisProperty.AppendLine("Update the Occurence Properties")
        sbUpdateOccReferenceAnalysisProperty.AppendLine($"*************************************************************")
        Try
            ' Register with OLE to handle concurrency issues on the current thread.
            SolidEdgeCommunity.OleMessageFilter.Register()

            ' Get a reference to the active assembly document.
            assemblyDocument = objApp.ActiveDocument 'application.GetActiveDocument(Of SolidEdgeAssembly.AssemblyDocument)(False)

            'Dim oCustom As Properties
            'Dim oProp As [Property]

            If assemblyDocument IsNot Nothing Then

                Dim interferenceCount As Integer = 1

                ' Get a reference to the Occurrences collection.
                occurrences = assemblyDocument.Occurrences
                Dim material As String = String.Empty
                For Each occurrence In occurrences.OfType(Of SolidEdgeAssembly.Occurrence)()

                    Try
                        material = GetOccurenceDocumentMaterialName(occurrence)

                        If selectedMaterialList.Contains(material) Then

                            occurrence.IncludeInInterference = False
                            sbUpdateOccReferenceAnalysisProperty.AppendLine($"{occurrence.Name}{vbTab}{vbTab}Material:{material}{vbTab}{vbTab}IncludeInInterference:{occurrence.IncludeInInterference}{vbTab}{vbTab}Path:{(occurrence.PartFileName)}")

                        End If
                    Catch ex As Exception
                        sbUpdateOccReferenceAnalysisProperty.AppendLine($"Error in Occurence {occurrence.Name}{vbTab}{vbTab} Materail {material} {vbTab}{vbTab} {ex.Message}{vbTab}{vbTab}{ex.StackTrace}")
                    End Try

                Next occurrence
            Else
                Throw New System.Exception("No active document.")
            End If
        Catch ex As System.Exception
            Console.WriteLine(ex.Message)

        Finally
            SolidEdgeCommunity.OleMessageFilter.Unregister()
        End Try
        sbUpdateOccReferenceAnalysisProperty.AppendLine("")
        sbUpdateOccReferenceAnalysisProperty.AppendLine("")
        CustomLogUtil.Log(sbUpdateOccReferenceAnalysisProperty.ToString())
        Debug.Print(sbUpdateOccReferenceAnalysisProperty.ToString())

        Return sbUpdateOccReferenceAnalysisProperty.ToString()
    End Function

    Dim sbUpdateOccReferenceAnalysisProperty As New StringBuilder()

    Public Function OccurencesReferenceAnalysisProperty(ByVal selectedMaterialList As List(Of String), ByVal occurrences As SolidEdgeAssembly.Occurrences) As String

        'Dim application As SolidEdgeFramework.Application = objApp
        'Dim assemblyDocument As SolidEdgeAssembly.AssemblyDocument = Nothing
        'Dim occurrences As SolidEdgeAssembly.Occurrences = Nothing
        Dim interferenceStatus As SolidEdgeAssembly.InterferenceStatusConstants = Nothing
        Dim compare As SolidEdgeConstants.InterferenceComparisonConstants = SolidEdgeConstants.InterferenceComparisonConstants.seInterferenceComparisonSet1vsAllOther
        Dim reportType As SolidEdgeConstants.InterferenceReportConstants = SolidEdgeConstants.InterferenceReportConstants.seInterferenceReportPartNames

        sbUpdateOccReferenceAnalysisProperty.AppendLine($"********************{System.DateTime.Now}********************")
        sbUpdateOccReferenceAnalysisProperty.AppendLine("Update the Occurence Properties")
        sbUpdateOccReferenceAnalysisProperty.AppendLine($"*************************************************************")
        Try
            ' Register with OLE to handle concurrency issues on the current thread.
            SolidEdgeCommunity.OleMessageFilter.Register()

            ' Get a reference to the active assembly document.
            ' assemblyDocument = objApp.ActiveDocument 'application.GetActiveDocument(Of SolidEdgeAssembly.AssemblyDocument)(False)

            'Dim oCustom As Properties
            'Dim oProp As [Property]

            Dim interferenceCount As Integer = 1

            ' Get a reference to the Occurrences collection.
            'Occurrences = AssemblyDocument.Occurrences
            Dim material As String = String.Empty
            For Each occurrence As SolidEdgeAssembly.Occurrence In occurrences 'Occurrences.OfType(Of SolidEdgeAssembly.Occurrence)()

                Dim type As ObjectType = occurrence.Type
                If type = ObjectType.igSubAssembly Then
                    Dim assemblydoc As AssemblyDocument = occurrence.OccurrenceDocument
                    GetQty(assemblydoc.Occurrences, selectedMaterialList)
                End If

                Try
                    material = GetOccurenceDocumentMaterialName(occurrence)
                    If selectedMaterialList.Contains(material) Then
                        occurrence.IncludeInInterference = False
                        sbUpdateOccReferenceAnalysisProperty.AppendLine($"{occurrence.Name}{vbTab}{vbTab}Material:{material}{vbTab}{vbTab}IncludeInInterference:{occurrence.IncludeInInterference}{vbTab}{vbTab}Path:{(occurrence.PartFileName)}")

                    End If
                Catch ex As Exception
                    sbUpdateOccReferenceAnalysisProperty.AppendLine($"Error in Occurence {occurrence.Name}{vbTab}{vbTab} Materail {material} {vbTab}{vbTab} {ex.Message}{vbTab}{vbTab}{ex.StackTrace}")
                End Try

            Next occurrence

            ' Throw New System.Exception("No active document.")
        Catch ex As System.Exception
            Console.WriteLine(ex.Message)
            CustomLogUtil.Log("Error While Generating Report", ex.Message, ex.StackTrace)
        Finally
            SolidEdgeCommunity.OleMessageFilter.Unregister()
        End Try

        sbUpdateOccReferenceAnalysisProperty.AppendLine("")

        sbUpdateOccReferenceAnalysisProperty.AppendLine("")

        Debug.Print(sbUpdateOccReferenceAnalysisProperty.ToString())

        Return sbUpdateOccReferenceAnalysisProperty.ToString()
    End Function

    Public Sub GetQty(ByVal occs As SolidEdgeAssembly.Occurrences, ByVal selectedMaterialList As List(Of String))

        'Dim oCustom As SolidEdgeFramework.Properties
        'Dim oProp As SolidEdgeFramework.Property

        Dim material As String = String.Empty

        For Each occ As SolidEdgeAssembly.Occurrence In occs

            Dim obj As ObjectType = occ.Type

            If obj = ObjectType.igPart Then
                Try

                    material = GetOccurenceDocumentMaterialName(occ)

                    If selectedMaterialList.Contains(material) Then

                        occ.IncludeInInterference = False
                        sbUpdateOccReferenceAnalysisProperty.AppendLine($"{occ.Name}{vbTab}{vbTab}Material:{material}{vbTab}{vbTab}IncludeInInterference:{occ.IncludeInInterference}{vbTab}{vbTab}Path:{(occ.PartFileName)}")

                    End If
                Catch ex As Exception
                End Try
            End If
            If obj = ObjectType.igSubAssembly Then
                Try
                    Dim assembldocument As AssemblyDocument = occ.OccurrenceDocument
                    GetQty(assembldocument.Occurrences, selectedMaterialList)
                Catch ex As Exception
                End Try
            End If

        Next

    End Sub

    Public Function GetOccurencesInterferenceDetails(ByVal selectedMaterialList As List(Of String)) As String

        Dim application As SolidEdgeFramework.Application = objApp
        Dim assemblyDocument As SolidEdgeAssembly.AssemblyDocument = Nothing
        Dim occurrences As SolidEdgeAssembly.Occurrences = Nothing
        Dim interferenceStatus As SolidEdgeAssembly.InterferenceStatusConstants = Nothing
        Dim compare As SolidEdgeConstants.InterferenceComparisonConstants = SolidEdgeConstants.InterferenceComparisonConstants.seInterferenceComparisonSet1vsAllOther
        Dim reportType As SolidEdgeConstants.InterferenceReportConstants = SolidEdgeConstants.InterferenceReportConstants.seInterferenceReportPartNames
        Dim sbInterfearance As New StringBuilder()
        Dim dtIntereferenceReport As New DataTable("Intereference Report")

        sbInterfearance.AppendLine($"********************{System.DateTime.Now}********************")
        sbInterfearance.AppendLine("Interferences")
        sbInterfearance.AppendLine($"*************************************************************")
        Try
            ' Register with OLE to handle concurrency issues on the current thread.
            SolidEdgeCommunity.OleMessageFilter.Register()

            '' Connect to or start Solid Edge.
            'application = SolidEdgeCommunity.SolidEdgeUtils.Connect(True, True)

            ' Get a reference to the active assembly document.
            assemblyDocument = objApp.ActiveDocument 'application.GetActiveDocument(Of SolidEdgeAssembly.AssemblyDocument)(False)

            'Dim oCustom As Properties
            'Dim oProp As [Property]

            If assemblyDocument IsNot Nothing Then

                Dim interferenceCount As Integer = 1

                ' Get a reference to the Occurrences collection.
                occurrences = assemblyDocument.Occurrences

                For Each occurrence In occurrences.OfType(Of SolidEdgeAssembly.Occurrence)()

                    'Dim obj As ObjectType = occurrence.Type

                    'oCustom = occurrence.PartDocument.Properties("Custom")

                    'oProp = oCustom.Item("Material Used")

                    Dim material As String = "" 'oProp.Value

                    Try
                        material = GetOccurenceDocumentMaterialName(occurrence)
                    Catch ex As Exception

                    End Try
                    If selectedMaterialList.Contains(material) Then
                        Continue For
                    End If

                    'Dim objDocument As SolidEdgePart.PartDocument = occurrence.PartDocument

                    'occurrence.OccurrenceDocument.Models.Model.GetMultiBodyPublishFileName()

                    'Dim numOfBodies As Integer
                    'Dim bodyArra As Array
                    'occurrence.GetSimplifiedBodies(numOfBodies, bodyArra)

                    Dim set1 As Array = Array.CreateInstance(occurrence.GetType(), 1)
                    Dim numInterferences As Object = 0
                    Dim retSet1 As Object = Array.CreateInstance(GetType(SolidEdgeAssembly.Occurrence), 0)
                    Dim retSet2 As Object = Array.CreateInstance(GetType(SolidEdgeAssembly.Occurrence), 0)
                    Dim confirmedInterference As Object = Nothing
                    Dim interferenceOccurrence As Object = Nothing

                    set1.SetValue(occurrence, 0)

                    ' Check interference.
                    'assemblyDocument.CheckInterference(NumElementsSet1:=set1.Length, Set1:=set1, Status:=interferenceStatus, ComparisonMethod:=compare, NumElementsSet2:=0, Set2:=Missing.Value, AddInterferenceAsOccurrence:=True, ReportFilename:=Missing.Value, ReportType:=reportType, NumInterferences:=numInterferences, InterferingPartsSet1:=retSet1, InterferingPartsOtherSet:=retSet2, ConfirmedInterference:=confirmedInterference, InterferenceOccurrence:=interferenceOccurrence, IgnoreThreadInterferences:=Missing.Value)
                    assemblyDocument.CheckInterference(NumElementsSet1:=set1.Length, Set1:=set1, Status:=interferenceStatus, ComparisonMethod:=compare, NumElementsSet2:=0, Set2:=Missing.Value, AddInterferenceAsOccurrence:=False, "IntereferenceReport.txt", ReportType:=reportType, NumInterferences:=numInterferences, InterferingPartsSet1:=retSet1, InterferingPartsOtherSet:=retSet2, ConfirmedInterference:=confirmedInterference, InterferenceOccurrence:=interferenceOccurrence, IgnoreThreadInterferences:=Missing.Value)

                    ' Process status.
                    Select Case interferenceStatus
                        Case SolidEdgeAssembly.InterferenceStatusConstants.seInterferenceStatusNoInterference
                        Case SolidEdgeAssembly.InterferenceStatusConstants.seInterferenceStatusConfirmedAndProbableInterference, SolidEdgeAssembly.InterferenceStatusConstants.seInterferenceStatusConfirmedInterference, SolidEdgeAssembly.InterferenceStatusConstants.seInterferenceStatusIncompleteAnalysis, SolidEdgeAssembly.InterferenceStatusConstants.seInterferenceStatusProbableInterference
                            If retSet2 IsNot Nothing Then
                                For j As Integer = 0 To (DirectCast(numInterferences, Integer)) - 1
                                    Dim obj1 As Object = DirectCast(retSet1, Array).GetValue(j)
                                    Dim obj2 As Object = DirectCast(retSet2, Array).GetValue(j)

                                    ' Use helper class to get the object type.
                                    Dim objectType1 = SolidEdgeCommunity.Runtime.InteropServices.ComObject.GetPropertyValue(Of SolidEdgeFramework.ObjectType)(obj1, "Type", CType(0, SolidEdgeFramework.ObjectType))
                                    Dim objectType2 = SolidEdgeCommunity.Runtime.InteropServices.ComObject.GetPropertyValue(Of SolidEdgeFramework.ObjectType)(obj2, "Type", CType(0, SolidEdgeFramework.ObjectType))

                                    Dim reference1 As SolidEdgeFramework.Reference = Nothing
                                    Dim reference2 As SolidEdgeFramework.Reference = Nothing
                                    Dim occurrence1 As SolidEdgeAssembly.Occurrence = Nothing
                                    Dim occurrence2 As SolidEdgeAssembly.Occurrence = Nothing

                                    Select Case objectType1
                                        Case SolidEdgeFramework.ObjectType.igReference
                                            reference1 = DirectCast(obj1, SolidEdgeFramework.Reference)
                                        Case SolidEdgeFramework.ObjectType.igPart, SolidEdgeFramework.ObjectType.igOccurrence
                                            occurrence1 = DirectCast(obj1, SolidEdgeAssembly.Occurrence)
                                    End Select

                                    Select Case objectType2
                                        Case SolidEdgeFramework.ObjectType.igReference
                                            reference2 = DirectCast(obj2, SolidEdgeFramework.Reference)
                                        Case SolidEdgeFramework.ObjectType.igPart, SolidEdgeFramework.ObjectType.igOccurrence
                                            occurrence2 = DirectCast(obj2, SolidEdgeAssembly.Occurrence)
                                    End Select

                                    Try

                                        Dim material1 As String = GetOccurenceDocumentMaterialName(occurrence) ' GetOccurencePropertyMaterialName(occurrence1)
                                        Dim material2 As String = GetOccurenceDocumentMaterialName(occurrence) ' GetOccurencePropertyMaterialName(occurrence2)

                                        If selectedMaterialList.Contains(material1) Then ' = "RUBBER" Then
                                            Continue For
                                        End If

                                        If selectedMaterialList.Contains(material2) Then 'If material2 = "RUBBER" Then
                                            Continue For
                                        End If

                                        occurancelist.Add(occurrence1.Name)

                                        sbInterfearance.AppendLine("")
                                        sbInterfearance.AppendLine("")
                                        sbInterfearance.AppendLine($"Interference {interferenceCount}")

                                        sbInterfearance.AppendLine($"{occurrence1.Name}{vbTab}{vbTab}Material:{material1}{vbTab}{vbTab}Path:{(occurrence1.PartFileName)}")
                                        sbInterfearance.AppendLine($"{occurrence2.Name}{vbTab}{vbTab}Material:{material2}{vbTab}{vbTab}Path:{(occurrence2.PartFileName)}")
                                        interferenceCount += 1
                                    Catch ex As Exception
                                        sbInterfearance.AppendLine(ex.Message)
                                    End Try

                                Next j
                            End If
                    End Select
                Next occurrence
            Else
                Throw New System.Exception("No active document.")
            End If
        Catch ex As System.Exception
            Console.WriteLine(ex.Message)
        Finally
            SolidEdgeCommunity.OleMessageFilter.Unregister()
        End Try
        sbInterfearance.AppendLine("")
        sbInterfearance.AppendLine("")
        Debug.Print(sbInterfearance.ToString())

        Return sbInterfearance.ToString()
    End Function

    Private Sub ComboBoxFields_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxFields.SelectedIndexChanged
        Try
            dgvDocumentDetails.DataSource = dtInteferenceMaterials
            SetAutoComplete_CSV()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("Auto complete." + vbNewLine + ex.Message, "Search", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub SetAutoComplete_CSV()
        Try
            txtSearch.Text = String.Empty
            Dim lst As New List(Of String)
            Dim MySource As New AutoCompleteStringCollection()

            If ComboBoxFields.SelectedIndex > 0 Then
                For Each outerList As DataGridViewRow In dgvDocumentDetails.Rows
                    lst.Add(outerList.Cells((ComboBoxFields.SelectedIndex) - 1).Value.ToString()) '.Cells(ComboBoxFields.SelectedIndex))
                Next
            End If

            lst = lst.Distinct().ToList()
            MySource.AddRange(lst.ToArray)
            txtSearch.AutoCompleteCustomSource = MySource
            txtSearch.AutoCompleteMode = AutoCompleteMode.SuggestAppend
            txtSearch.AutoCompleteSource = AutoCompleteSource.CustomSource
            'lst.Clear()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TxtSearch_KeyUp(sender As Object, e As KeyEventArgs) Handles txtSearch.KeyUp
        Try
            If (e.KeyValue = Keys.Enter) Then
                BtnSearchFile_Click(sender, e)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub BtnSearchFile_Click(sender As Object, e As EventArgs) Handles btnSearchFile.Click
        TestRemoveCode()
        'TestRemoveCode()
    End Sub

    Private Sub TestRemoveCode()
        Try
            ControlsEnability(False)
            If txtSearch.Text.Trim = String.Empty Then
                ' dgvDocumentDetails.DataSource = Nothing
                dgvDocumentDetails.DataSource = dtInteferenceMaterials
                'Exit Sub
            End If

            If txtSearch.Text = String.Empty Then
                'Dim dt As DataTable = DataGridViewComments.DataSource
                'Dim DV As DataView = New DataView(dt)
                'Try
                '    DV.RowFilter = String.Format("" + ComboBoxFields.Text + " = ''")
                'Catch ex As Exception
                'End Try
                'DataGridViewComments.DataSource = DV
            Else
                Dim dt As DataTable = dtInteferenceMaterials
                Dim DV As New DataView(dt)
                Try
                    DV.RowFilter = String.Format("" + ComboBoxFields.Text + " LIKE '%{0}%'", txtSearch.Text)
                Catch ex As Exception
                    DV.RowFilter = String.Format("" + ComboBoxFields.Text + "={0}", txtSearch.Text)
                End Try
                dgvDocumentDetails.DataSource = DV
            End If
        Catch ex As Exception
            MessageBox.Show("Unable to search.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CustomLogUtil.Log($"Unable To Search..", ex.Message, ex.StackTrace)
        Finally

            ControlsEnability(True)
            'SetGridValidationColor()

        End Try
    End Sub

    Public Sub ControlsEnability(ByVal flag As Boolean)

        'btnOpenDocument.Enabled = flag
        'cmbCategory.Enabled = flag
        'btnRefresh.Enabled = flag
        'btnApply.Enabled = flag
        'cmbCategory.Enabled = flag
        'cmbMaterialUsed2_Mw.Enabled = flag

    End Sub

    Private Sub DgvDocumentDetails_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvDocumentDetails.CellValueChanged

        'Dim rowIndex As Integer = e.RowIndex
        'Dim selected As Boolean = dgvDocumentDetails.CurrentRow.Cells("Select").Value.ToString 'dgvDocumentDetails("Select")(rowIndex - 1).value
        'Dim srVal As String = dgvDocumentDetails.CurrentRow.Cells("Sr").Value.ToString 'dgvDocumentDetails("Select")(rowIndex - 1).value

    End Sub

    Private Sub DgvDocumentDetails_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvDocumentDetails.CellContentClick
        'Dim rowIndex As Integer = e.RowIndex
        'Dim selected As Boolean = dgvDocumentDetails.CurrentRow.Cells("Select").Value.ToString 'dgvDocumentDetails("Select")(rowIndex - 1).value
        'Dim srVal As String = dgvDocumentDetails.CurrentRow.Cells("Sr").Value.ToString 'dgvDocumentDetails("Select")(rowIndex - 1).value
        'Dim dt As DataTable = dgvDocumentDetails.DataSource

        'Debug.Print("aaaa")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        CustomLogUtil.Heading("To False Interferance Occurance Properties Starts")
        WaitStartSave()
        CustomLogUtil.Log("GetSelectedMaterials")
        Dim selectedMaterialList As List(Of String) = GetSelectedMaterials()

        Dim application As SolidEdgeFramework.Application = objApp

        Dim assemblyDocument As SolidEdgeAssembly.AssemblyDocument = Nothing

        Dim occurrences As SolidEdgeAssembly.Occurrences = Nothing

        assemblyDocument = objApp.ActiveDocument
        CustomLogUtil.Log("OccurencesReferenceAnalysisProperty")
        Dim str As String = OccurencesReferenceAnalysisProperty(selectedMaterialList, assemblyDocument.Occurrences)
        CustomLogUtil.Log("UpdateOccurrencePropertyReport")
        UpdateOccurrencePropertyReport(str)
        CustomLogUtil.Heading("To False Interferance Occurance Properties Report Completed")
        WaitEndSave()

        MessageBox.Show("Completed")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        CustomLogUtil.Heading("Check Interference TopLevel Report Starts")
        cnt = 1
        WaitStartSave()
        CustomLogUtil.Heading("Interfearnce")
        Interfearnce()
        WaitEndSave()
        MessageBox.Show("Report Generated successfully", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        CustomLogUtil.Heading("Report Generated successfully")
        objAssemblyDocument = objApp.ActiveDocument
        Dim fullpah As String = objAssemblyDocument.FullName
        Dim Location As String = IO.Path.GetDirectoryName(fullpah)
        Process.Start(Location)
    End Sub

    Public Sub Interfearnce()
        Try
            Dim nComparisonMethod As Integer
            Dim nSet1 As Integer
            Dim nSet2 As Integer
            Dim objOccurrences As SolidEdgeAssembly.Occurrences = Nothing
            Dim objOccurrence As SolidEdgeAssembly.Occurrence = Nothing
            Dim objInterfOcc As SolidEdgeAssembly.Occurrence = Nothing
            Dim a_objSet1() As Object
            Dim objTemp As Object
            Dim nNumInterferences As Long
            Dim nStatus As SolidEdgeAssembly.InterferenceStatusConstants

            Dim application As SolidEdgeFramework.Application = objApp
            Dim objAsm As SolidEdgeAssembly.AssemblyDocument = Nothing
            Dim occurrences As SolidEdgeAssembly.Occurrences = Nothing
            objAsm = objApp.ActiveDocument

            ' Dim reportPath As String = $"C:\Users\vimalb\Desktop\New Text Document{cnt.ToString()}.txt"
            Dim fullpath As String = objAsm.FullName
            Dim reportname As Object = IO.Path.GetDirectoryName(fullpath) + "\" + "report.txt"
            Dim filename = Nothing

            Dim datetime As String = System.DateTime.Now.ToString("dd/MM/yy_HH.mm.ss")
            datetime = datetime.Replace("/", "-")
            objApp.DisplayAlerts = False
            objAsm = objApp.ActiveDocument
            nComparisonMethod = SolidEdgeConstants.InterferenceComparisonConstants.seInterferenceComparisonSet1vsAllOther
            nSet1 = 0
            nSet2 = 0
            objOccurrences = objAsm.Occurrences
            Dim sb As New StringBuilder()
            ReDim a_objSet1(0)
            For nIndex = 1 To objOccurrences.Count
                objTemp = objOccurrences.Item(nIndex)
                Dim type As ObjectType = objTemp.type
                a_objSet1(0) = objTemp
                nSet1 = 1
                filename = IO.Path.GetFileNameWithoutExtension(fullpath) + "_" + "TopLevel_InterferenceReport" + "_" + datetime + ".txt"
                'Dim filename As String = IO.Path.GetFileNameWithoutExtension(objTemp.Occurrencefilename) + "_" + "InterferenceReport1" + ".txt"
                Try
                    'Add intereference as part in assembly
                    Call objAsm.CheckInterference2(
                    NumElementsSet1:=nSet1,
                    Set1:=a_objSet1,
                    Status:=nStatus,
                    ComparisonMethod:=nComparisonMethod,
                    AddInterferenceAsOccurrence:=False,
                    NumInterferences:=nNumInterferences,
                    InterferenceOccurrence:=objInterfOcc,
                    IgnoreNonThreadVsThreadConstant:=SolidEdgeConstants.InterferenceOptionsConstants.seIntfOptIgnoreThreadVsNonThreaded,
                    ReportFilename:="Report.txt") ', ReportType:="TEXT") ',
                Catch ex As Exception

                End Try




                Dim status As String = Nothing
                If nStatus = SolidEdgeAssembly.InterferenceStatusConstants.seInterferenceStatusConfirmedInterference Then
                    status = "ConfirmedInterference"
                ElseIf nStatus = SolidEdgeAssembly.InterferenceStatusConstants.seInterferenceStatusNoInterference Then
                    status = "NoInterference"
                ElseIf nStatus = SolidEdgeAssembly.InterferenceStatusConstants.seInterferenceStatusConfirmedAndProbableInterference Then
                    status = "ConfirmedAndProbableInterference "
                ElseIf nStatus = SolidEdgeAssembly.InterferenceStatusConstants.seInterferenceStatusIncompleteAnalysis Then
                    status = "IncompleteAnalysis"
                ElseIf nStatus = SolidEdgeAssembly.InterferenceStatusConstants.seInterferenceStatusProbableInterference Then
                    status = "ProbableInterference"
                End If




                sb.AppendLine(IO.Path.GetFileNameWithoutExtension(objTemp.Occurrencefilename) + "_Status:" + status)

                ' RenameFile(cnt, reportname)

                ' cnt = cnt + 1
                ' a_objSet1 = Nothing
            Next nIndex
            Dim statusfilenamerepo As String = IO.Path.GetFileNameWithoutExtension(fullpath) + "_" + "TopLevelInterferenceReportStatus" + "_" + datetime + ".txt"
            'Dim statusfilenamerepo As String = IO.Path.GetFileNameWithoutExtension(fullpath) + "_" + "TopLevelInterferenceReportStatus" + ".txt"
            Dim statusreponame As Object = IO.Path.GetDirectoryName(fullpath) + "\" + statusfilenamerepo
            Dim location As String = IO.Path.GetDirectoryName(fullpath)


            Dim objWriter As New System.IO.StreamWriter(statusreponame, True)
            objWriter.WriteLine(sb)
            objWriter.Close()

            Dim ReportName1 = IO.Path.GetDirectoryName(fullpath) + "\Report.txt"
            If IO.File.Exists(ReportName1) Then
                My.Computer.FileSystem.RenameFile(ReportName1, filename)
            End If
            Process.Start(location)
        Catch ex As Exception

            If skip = True Then
                MessageBox.Show($"Error While Generating Report", "Error")
            End If
            CustomLogUtil.Log("Error While Generating Report", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub RenameFile(ByVal cnt As Integer, ByVal reportname As String)

        IO.File.Move(reportname, $"{reportname.Replace(".txt", $"{cnt}.txt")}")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        CustomLogUtil.Heading("Check child interferences Report Starts")
        WaitStartSave()
        CustomLogUtil.Log("Childinterference")
        Childinterference()
        WaitEndSave()

        MessageBox.Show("Report Generated successfully", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        CustomLogUtil.Log("Report Generated successfully")
        objAssemblyDocument = objApp.ActiveDocument
        Dim fullpah As String = objAssemblyDocument.FullName
        Dim Location As String = IO.Path.GetDirectoryName(fullpah)
        Process.Start(Location)
    End Sub

    Public Sub Childinterference()
        Try
            Dim nComparisonMethod As Integer
            Dim nSet1 As Integer
            Dim nSet2 As Integer
            Dim objOccurrences As SolidEdgeAssembly.Occurrences = Nothing
            Dim objOccurrence As SolidEdgeAssembly.Occurrence = Nothing
            Dim objInterfOcc As SolidEdgeAssembly.Occurrence = Nothing
            Dim a_objSet1() As Object
            Dim objTemp As Object
            Dim nNumInterferences As Long
            Dim nStatus As SolidEdgeAssembly.InterferenceStatusConstants


            objApp.DisplayAlerts = False
            objAssemblyDocument = objApp.ActiveDocument

            Dim fullpath As String = objAssemblyDocument.FullName

            Dim datetime As String = System.DateTime.Now.ToString("dd/MM/yy_HH.mm.ss")
            datetime = datetime.Replace("/", "-")

            Dim filename As String = IO.Path.GetFileNameWithoutExtension(fullpath) + "_" + "Child_InterferenceReport" + "_" + datetime + ".txt"
            'Dim filename As String = "InterferenceReport.txt"
            Dim reportname As Object = "ChildInterferenceReport.txt"
            nComparisonMethod = SolidEdgeConstants.InterferenceComparisonConstants.seInterferenceComparisonSet1vsItself
            nSet1 = 0
            nSet2 = 0
            objOccurrences = objAssemblyDocument.Occurrences
            ReDim a_objSet1(objOccurrences.Count)
            For nIndex = 1 To objOccurrences.Count
                objTemp = objOccurrences.Item(nIndex)
                a_objSet1(nSet1) = objTemp
                nSet1 += 1
            Next nIndex
            Try
                'Add intereference as part in assembly
                Call objAssemblyDocument.CheckInterference2(
                                        NumElementsSet1:=nSet1,
                                        Set1:=a_objSet1,
                                        Status:=nStatus,
                                        ComparisonMethod:=nComparisonMethod,
                                        AddInterferenceAsOccurrence:=False,
                                        NumInterferences:=nNumInterferences,
                                        InterferenceOccurrence:=objInterfOcc,
                                        IgnoreNonThreadVsThreadConstant:=SolidEdgeConstants.InterferenceOptionsConstants.seIntfOptIgnoreThreadVsNonThreaded,
            ReportFilename:="Report.txt") ', ReportType:="TEXT") ',

                'ReportFilename:=filename

            Catch ex As Exception

            End Try


            'Call objAssemblyDocument.CheckInterference2(
            '    NumElementsSet1:=nSet1,
            '    Set1:=a_objSet1,
            '    Status:=nStatus,
            '    ComparisonMethod:=nComparisonMethod,
            '    AddInterferenceAsOccurrence:=True,
            '    NumInterferences:=nNumInterferences,
            '    InterferenceOccurrence:=objInterfOcc,
            '    IgnoreNonThreadVsThreadConstant:=SolidEdgeConstants.InterferenceOptionsConstants.seIntfOptIgnoreThreadVsNonThreaded,
            '    ReportFilename:=reportname)

            Dim status As String = String.Empty
            If nStatus = SolidEdgeAssembly.InterferenceStatusConstants.seInterferenceStatusConfirmedInterference Then
                status = "ConfirmedInterference"
            ElseIf nStatus = SolidEdgeAssembly.InterferenceStatusConstants.seInterferenceStatusNoInterference Then
                status = "NoInterference"
            ElseIf nStatus = SolidEdgeAssembly.InterferenceStatusConstants.seInterferenceStatusConfirmedAndProbableInterference Then
                status = "ConfirmedAndProbableInterference "
            ElseIf nStatus = SolidEdgeAssembly.InterferenceStatusConstants.seInterferenceStatusIncompleteAnalysis Then
                status = "IncompleteAnalysis"
            ElseIf nStatus = SolidEdgeAssembly.InterferenceStatusConstants.seInterferenceStatusProbableInterference Then
                status = "ProbableInterference"
            End If
            Dim statusfilenamerepo As String = IO.Path.GetFileNameWithoutExtension(fullpath) + "_" + "Child_InterferenceReportStatus" + "_" + datetime + ".txt"
            'Dim statusfilenamerepo As String = IO.Path.GetFileNameWithoutExtension(fullpath) + "_" + "ChildInterferenceReportStatus.txt"
            Dim statusreponame As Object = IO.Path.GetDirectoryName(fullpath) + "\" + statusfilenamerepo
            'Dim statusreponame1 As Object = statusfilenamerepo
            Dim sb As New StringBuilder()
            sb.Append("Status:")
            sb.Append(status)

            Dim objWriter As New System.IO.StreamWriter(statusreponame, True)
            objWriter.WriteLine(sb)
            objWriter.Close()
            Dim ReportName1 = IO.Path.GetDirectoryName(fullpath) + "\Report.txt"
            If IO.File.Exists(ReportName1) Then
                My.Computer.FileSystem.RenameFile(ReportName1, filename)
            End If
        Catch ex As Exception
            If skip = True Then
                MessageBox.Show($"Error While Fetching Data in Child interface", "Error")
            End If
            CustomLogUtil.Log($"Error While Fetching Data in Child interface", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub BtnInterferenceExcludeMaterial_Click(sender As Object, e As EventArgs) Handles btnInterferenceExcludeMaterial.Click
        Try
            CustomLogUtil.Heading("Get Interference Exclude Material Process Starts")
            If IO.File.Exists(txtExcelPath.Text) Then

                Dim dictIntereference As New Dictionary(Of String, DataSet)()
                Dim filepath As String = txtExcelPath.Text ' $"{System.IO.Directory.GetCurrentDirectory}\Bom_20220503-1.xlsx"
                CustomLogUtil.Heading("ReadRawMaterials2Interference")
                dictIntereference = ExcelUtil.ReadRawMaterials2Interference(filepath)
                Dim ds As DataSet = dictIntereference.Values(0)
                dtInteferenceMaterials = ds.Tables(0)
                dgvDocumentDetails.DataSource = dtInteferenceMaterials
                'FillMaterials(dt)
                FillSearchCombo()
            Else
                MsgBox("Please select material excel details.")
            End If
            CustomLogUtil.Heading("Get Interference Exclude Material Process Completed")
        Catch ex As Exception
            MessageBox.Show($"Error While Fetching Data From Interference ExcludeMaterial", "Error")
            CustomLogUtil.Log($"Error While Fetching Data From Interference ExcludeMaterial", ex.Message, ex.StackTrace)
        End Try
    End Sub
End Class