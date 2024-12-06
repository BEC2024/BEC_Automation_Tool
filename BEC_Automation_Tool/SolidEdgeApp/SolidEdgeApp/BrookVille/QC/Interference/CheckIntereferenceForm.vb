Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports SolidEdgeAssembly
Imports SolidEdgeFramework

Public Class CheckIntereferenceForm

    Dim dictMaterials As Dictionary(Of String, List(Of String)) = New Dictionary(Of String, List(Of String))()
    Dim objApp As SolidEdgeFramework.Application = Nothing
    Dim objDocument As SolidEdgeFramework.SolidEdgeDocument = Nothing
    Dim objSheetMetalDocument As SolidEdgePart.SheetMetalDocument = Nothing
    Dim objAssemblyDocument As SolidEdgeAssembly.AssemblyDocument
    Dim dtInteferenceMaterials As DataTable = New DataTable()
    Dim myMatTable As SolidEdgeFramework.MatTable = Nothing

    Private Sub btnBrowseExcel_Click(sender As Object, e As EventArgs) Handles btnBrowseExcel.Click
        Try
            Using dialog As New OpenFileDialog
                dialog.Filter = "Excel files (*.xlsx)|*.xlsx"
                If dialog.ShowDialog() <> DialogResult.OK Then Return
                txtExcelPath.Text = dialog.FileName
            End Using

            If IO.File.Exists(txtExcelPath.Text) Then

                Dim dictIntereference As Dictionary(Of String, DataSet) = New Dictionary(Of String, DataSet)()
                Dim filepath As String = txtExcelPath.Text ' $"{System.IO.Directory.GetCurrentDirectory}\Bom_20220503-1.xlsx"
                dictIntereference = ExcelUtil.ReadRawMaterials2Interference(filepath)
                Dim ds As DataSet = dictIntereference.Values(0)
                dtInteferenceMaterials = ds.Tables(0)
                dgvDocumentDetails.DataSource = dtInteferenceMaterials
                'FillMaterials(dt)
                FillSearchCombo()
            Else
                MsgBox("Please select material excel details.")
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

    Dim occurancelist As List(Of String) = New List(Of String)

    Private Sub CheckIntereferenceForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        SetSolidEdgeInstance()

        SetControls(dgvDocumentDetails)

        myMatTable = objApp.GetMaterialTable()

        Me.Text = $"{Me.Text} ({GlobalEntity.Version})"

    End Sub

    Private Sub SetSolidEdgeInstance()
        Try
            objApp = Marshal.GetActiveObject("SolidEdge.Application")
        Catch ex As Exception
            MessageBox.Show($"Error in fetching the Solid-Edge instance", "Error")
            CustomLogUtil.Log($"Error in fetching the Solid-Edge instance", ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub SetControls(ByRef DataGridViewComments As DataGridView)
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
    End Sub

    Private Function GetSelectedMaterials() As List(Of String)
        Dim dt As DataTable = dtInteferenceMaterials
        Dim DVSelected As DataView = New DataView(dt, "Select=True", "", DataViewRowState.CurrentRows)
        Dim dtSelected As DataTable = DVSelected.ToTable()
        Dim selectedMaterialList As List(Of String) = New List(Of String)()
        For Each dr As DataRow In dtSelected.Rows
            If Not selectedMaterialList.Contains(dr(ExcelUtil.ExcelSheetColumnsIntereference.Material.ToString())) Then
                selectedMaterialList.Add(dr(ExcelUtil.ExcelSheetColumnsIntereference.Material.ToString()))
            End If
        Next
        Return selectedMaterialList
    End Function

    Private Sub btnSelectParts_Click(sender As Object, e As EventArgs) Handles btnGenerateInterferenceReport.Click

        Dim selectedMaterialList As List(Of String) = GetSelectedMaterials()

        Dim interferenceDetails As String = GetOccurencesInterferenceDetails(selectedMaterialList)

        Dim dirPath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)

        Dim assemblyDocument1 As SolidEdgeAssembly.AssemblyDocument = objApp.ActiveDocument

        Dim docPath As String = assemblyDocument1.FullName

        Dim docName As String = IO.Path.GetFileNameWithoutExtension(docPath)

        Dim reportName As String = $"InterferenceDetails_{IO.Path.GetFileNameWithoutExtension(docName)}_{ System.DateTime.Now.ToString("yyyyMMdd")}.txt"

        Dim reportPath As String = IO.Path.Combine(dirPath, reportName)

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

        Dim assemblyDocument As SolidEdgeAssembly.AssemblyDocument = Nothing
        objAssemblyDocument = objApp.ActiveDocument

        Dim selectSet As SolidEdgeFramework.SelectSet = Nothing
        Dim highlightSets As SolidEdgeFramework.HighlightSets = Nothing
        Dim highlightSet As SolidEdgeFramework.HighlightSet = Nothing

        selectSet = objApp.ActiveSelectSet
        assemblyDocument = CType(objApp.ActiveDocument, SolidEdgeAssembly.AssemblyDocument)
        highlightSets = assemblyDocument.HighlightSets
        highlightSet = highlightSets.Add()

        highlightSet = GetAlloccur(objAssemblyDocument.Occurrences, highlightSet)
        selectSet.Add(highlightSet)

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
        myMatTable.GetCurrentMaterialName(occurrence.PartDocument, materialName)
        Return materialName
    End Function

    Public Function GetOccurencesInterferenceDetails(ByVal selectedMaterialList As List(Of String)) As String

        Dim application As SolidEdgeFramework.Application = objApp
        Dim assemblyDocument As SolidEdgeAssembly.AssemblyDocument = Nothing
        Dim occurrences As SolidEdgeAssembly.Occurrences = Nothing
        Dim interferenceStatus As SolidEdgeAssembly.InterferenceStatusConstants = Nothing
        Dim compare As SolidEdgeConstants.InterferenceComparisonConstants = SolidEdgeConstants.InterferenceComparisonConstants.seInterferenceComparisonSet1vsAllOther
        Dim reportType As SolidEdgeConstants.InterferenceReportConstants = SolidEdgeConstants.InterferenceReportConstants.seInterferenceReportPartNames
        Dim sbInterfearance As StringBuilder = New StringBuilder()
        Dim dtIntereferenceReport As DataTable = New DataTable("Intereference Report")

        sbInterfearance.AppendLine($"********************{System.DateTime.Now.ToString()}********************")
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
                                        sbInterfearance.AppendLine($"Interference {interferenceCount.ToString()}")

                                        sbInterfearance.AppendLine($"{occurrence1.Name}{vbTab}{vbTab}Material:{material1}{vbTab}{vbTab}Path:{(occurrence1.PartFileName)}")
                                        sbInterfearance.AppendLine($"{occurrence2.Name}{vbTab}{vbTab}Material:{material2}{vbTab}{vbTab}Path:{(occurrence2.PartFileName)}")
                                        interferenceCount = interferenceCount + 1
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

    Private Sub txtSearch_KeyUp(sender As Object, e As KeyEventArgs) Handles txtSearch.KeyUp
        Try
            If (e.KeyValue = Keys.Enter) Then
                btnSearchFile_Click(sender, e)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSearchFile_Click(sender As Object, e As EventArgs) Handles btnSearchFile.Click
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
                Dim DV As DataView = New DataView(dt)
                Try
                    DV.RowFilter = String.Format("" + ComboBoxFields.Text + " LIKE '%{0}%'", txtSearch.Text)
                Catch ex As Exception
                    DV.RowFilter = String.Format("" + ComboBoxFields.Text + "={0}", txtSearch.Text)
                End Try
                dgvDocumentDetails.DataSource = DV
            End If
        Catch ex As Exception
            MessageBox.Show("Unable to search." + vbNewLine + ex.Message, "Search", MessageBoxButtons.OK, MessageBoxIcon.Error)
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

    Private Sub dgvDocumentDetails_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvDocumentDetails.CellValueChanged

        'Dim rowIndex As Integer = e.RowIndex
        'Dim selected As Boolean = dgvDocumentDetails.CurrentRow.Cells("Select").Value.ToString 'dgvDocumentDetails("Select")(rowIndex - 1).value
        'Dim srVal As String = dgvDocumentDetails.CurrentRow.Cells("Sr").Value.ToString 'dgvDocumentDetails("Select")(rowIndex - 1).value

    End Sub

    Private Sub dgvDocumentDetails_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvDocumentDetails.CellContentClick
        'Dim rowIndex As Integer = e.RowIndex
        'Dim selected As Boolean = dgvDocumentDetails.CurrentRow.Cells("Select").Value.ToString 'dgvDocumentDetails("Select")(rowIndex - 1).value
        'Dim srVal As String = dgvDocumentDetails.CurrentRow.Cells("Sr").Value.ToString 'dgvDocumentDetails("Select")(rowIndex - 1).value
        'Dim dt As DataTable = dgvDocumentDetails.DataSource

        'Debug.Print("aaaa")
    End Sub

End Class