Imports System.Runtime.InteropServices
Imports System.Text
Imports SolidEdgeAssembly
Imports SolidEdgeFramework
Imports SolidEdgePart

Public Class AssemblyAutomationForm

    Public Enum AssemblyColumns

        ParentDocumentName
        DocumentName
        PartType
        Size
        Grade
        GageName
        MaterialThickness
        BendRadius
        MaterialUsed
        MaterialSpec
        BECMaterial
        DocumentPath
    End Enum

    Dim objApp As SolidEdgeFramework.Application = Nothing
    Dim objDocument As SolidEdgeFramework.SolidEdgeDocument = Nothing
    Dim objSheetMetalDocument As SolidEdgePart.SheetMetalDocument = Nothing
    Dim objAssemblyDocument As SolidEdgeAssembly.AssemblyDocument
    Dim dtAssemblyData As DataTable = Nothing
    Dim objMatTable As SolidEdgeFramework.MatTable = Nothing
    Dim dt As New DataTable("")
    Dim dictMaterials As New Dictionary(Of String, List(Of String))()
    Dim listOfLibraries As Object = Nothing
    Dim numMaterials As Long

    ReadOnly materialLib As String = "BEC MATERIAL LIBRARY"

    Dim mainObj As New MainClass

    Dim dicData As New Dictionary(Of String, DataSet)()

    Dim dtstructure As New DataTable("Structure")
    Dim dtsheetmetal As New DataTable("SheetMetal")
    Dim isSheetMetalPart As Boolean = False

    Private Function IsValid() As Boolean

        SetSolidEdgeInstance()
        DisableBtn()
        If objApp Is Nothing Then
            Return False
        Else
            Return True

        End If

    End Function
    Private Sub AssemblyAutomationForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load, MyBase.Resize
        'btnRefresh.Visible = False
        'btnApply.Visible = False
        'btnClose.Visible = False
        'btnOpenDocument.Visible = False
        CustomLogUtil.Heading("AssemblyAutomationForm Open.....")

        If IsValid() Then
            txtExcelPath.Text = Config.configObj.becMaterialExcelPath

            Me.BackColor = Color.WhiteSmoke

            SetControlFirst()

            dtAssemblyData = AddColumns(dtAssemblyData)

            SetSolidEdgeInstance()

            If Not mainObj.SolidEdgeinstance = "Close" Then

                SetMaterialTable()

                GetMaterialLibraryList()

                dictMaterials = GetMaterialCollection(listOfLibraries)

                SetControls(dgvDocumentDetails)

                'SplitContainer1.SplitterDistance = 50%

            End If
        Else
            MessageBox.Show("Please open Solid-Edge Assembly and restart The Application", "Message")
            CustomLogUtil.Log("Please open Solid-Edge Assembly and restart The Application", "", "")
        End If

    End Sub
    Private Sub DisableBtn()
        If objApp Is Nothing Then
            btnBrowseExcel.Enabled = False
            btnGetData.Enabled = False
            btnApply.Enabled = False
            btnClose.Enabled = False
            btnOpenDocument.Enabled = False
            btnRefresh.Enabled = False
        Else
            btnBrowseExcel.Enabled = True
            btnGetData.Enabled = True
            btnApply.Enabled = True
            btnClose.Enabled = True
            btnOpenDocument.Enabled = True
            btnRefresh.Enabled = True
        End If
    End Sub

    Private Sub SetControlFirst()
        ' btnBrowseExcel.Enabled = False
        btnApply.Enabled = False
        btnOpenDocument.Enabled = False
        btnRefresh.Enabled = False

    End Sub

    Private Sub SetControlGetAssemblyData()
        btnBrowseExcel.Enabled = True
    End Sub

    Private Sub SetControlBrowseExcel()
        btnApply.Enabled = True
        btnOpenDocument.Enabled = True
        btnRefresh.Enabled = True
    End Sub

    Private Function GetMaterialCollection(ByVal listOfLibraries As Object) As Dictionary(Of String, List(Of String))
        Dim dictMaterials As New Dictionary(Of String, List(Of String))()
        For Each libr As String In listOfLibraries

            Try
                Dim listOfMaterials1 As Object = Nothing
                Dim numMaterials1 As Long
                objMatTable.GetMaterialListFromLibrary(libr, numMaterials1, listOfMaterials1)

                Dim lstMaterials As New List(Of String)()
                For Each m1 As String In listOfMaterials1
                    lstMaterials.Add(m1)
                Next
                dictMaterials.Add(libr, lstMaterials)
            Catch ex As Exception

            End Try

        Next

        Return dictMaterials
    End Function

    Private Sub GetMaterialLibraryList()
        'TEMP_6SEPT203
        Try
            objMatTable.GetMaterialLibraryList(listOfLibraries, numMaterials)
        Catch ex As Exception

        End Try


    End Sub

    Private Sub SetGridValidationColor()

        Dim rowCnt As Integer = 0
        Dim dt As DataTable = dgvDocumentDetails.DataSource
        For Each dr As DataRow In dt.Rows

            If dr("IsValidData") = True Then
                dgvDocumentDetails.Rows(rowCnt).DefaultCellStyle.BackColor = Color.White
                dgvDocumentDetails.Rows(rowCnt).DefaultCellStyle.ForeColor = Color.DarkBlue
            Else
                dgvDocumentDetails.Rows(rowCnt).DefaultCellStyle.BackColor = Color.FromArgb(255, 199, 206)
                dgvDocumentDetails.Rows(rowCnt).DefaultCellStyle.ForeColor = Color.FromArgb(156, 0, 6)
            End If

            rowCnt += 1
        Next

    End Sub

    Private Sub ValidateAssemblyData(ByVal dtAssemblydata As DataTable, ByVal dtExcelData As DataTable)

        Dim rowCnt As Integer = 0

        For Each dr As DataRow In dtAssemblydata.Rows

            Try
                Dim materialUsed_C As String = dr(AssemblyColumns.MaterialUsed.ToString()).ToString()
                Dim becMaterial_C As String = dr(AssemblyColumns.BECMaterial.ToString()).ToString()
                Dim materialThickness_C As String = dr(AssemblyColumns.MaterialThickness.ToString()).ToString()

                Dim materialUsed_Excel As String = ExcelUtil.ExcelSheetColumns.Material_Used.ToString()
                Dim becMaterial_Excel As String = ExcelUtil.ExcelSheetColumns.BEC_Material.ToString()
                Dim materialThickness_Excel As String = ExcelUtil.ExcelSheetColumns.Thickness.ToString()

                Dim dv As New DataView(dtExcelData) With {
                    .RowFilter = $"{materialUsed_Excel}='{materialUsed_C}' And {becMaterial_Excel}='{becMaterial_C}' And {materialThickness_Excel}='{materialThickness_C}'"
                }

                If Not dv.Count > 0 Then
                    dgvDocumentDetails.Rows(rowCnt).DefaultCellStyle.BackColor = Color.FromArgb(255, 199, 206)
                    dgvDocumentDetails.Rows(rowCnt).DefaultCellStyle.ForeColor = Color.FromArgb(156, 0, 6)
                    ' dgvDocumentDetails.Rows(rowCnt)("IsValidData").va
                    dgvDocumentDetails.Rows(rowCnt).Cells("IsValidData").Value = 0
                Else
                    dgvDocumentDetails.Rows(rowCnt).DefaultCellStyle.BackColor = Color.White
                    dgvDocumentDetails.Rows(rowCnt).DefaultCellStyle.ForeColor = Color.DarkBlue
                    dgvDocumentDetails.Rows(rowCnt).Cells("IsValidData").Value = 1
                End If

                'txtMaterialUsed_C.BackColor = Color.FromArgb(255, 199, 206)
                'txtMaterialUsed_C.ForeColor = Color.FromArgb(156, 0, 6)

                'txtSize_C.Text = dr(AssemblyColumns.Size.ToString()).ToString()
                'txtGrade_C.Text = dr(AssemblyColumns.Grade.ToString()).ToString()
                'txtGageName_C.Text = dr(AssemblyColumns.GageName.ToString()).ToString()
                'txtGageTable.Text = String.Empty 'dr(AssemblyColumns.DocumentPath.ToString()).ToString()

                'txtBendRadius_C.Text = dr(AssemblyColumns.BendRadius.ToString()).ToString()
                'txtPartType_C.Text = dr(AssemblyColumns.PartType.ToString()).ToString()

                'txtMaterialSpec_C.Text = dr(AssemblyColumns.MaterialSpec.ToString()).ToString()
            Catch ex As Exception
                CustomLogUtil.Log("While Validate Assembly Data", ex.Message, ex.StackTrace)
            End Try

            rowCnt += 1
        Next

    End Sub

    Public Sub Rnd()

        Dim dt As New DataTable()
        dt.Columns.Add(AssemblyColumns.DocumentPath.ToString())

        Dim dr As DataRow = dt.NewRow()
        dr(AssemblyColumns.DocumentPath.ToString) = "aaa"

        dt.Rows.Add(dr)

        dgvDocumentDetails.DataSource = dt

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

        DataGridViewComments.ReadOnly = True
    End Sub

    Private Sub SetMaterialTable()
        objMatTable = objApp.GetMaterialTable()
    End Sub

    Private Sub SetDocumentNames()

        lblAssemblyName.Text = IO.Path.GetFileName(objAssemblyDocument.FullName)
        lblAssemblyPath.Text = objAssemblyDocument.FullName

    End Sub

    Private Sub SetSolidEdgeInstance()
        Try

            objApp = Marshal.GetActiveObject("SolidEdge.Application")
        Catch ex As Exception
            'MsgBox($"Please Open Solid-Edge")
            'CustomLogUtil.Log("Please Open Solid-Edge", ex.Message, ex.StackTrace)

        End Try

    End Sub

    Public Sub Closefn(mainObj As MainClass)
        mainObj.SolidEdgeinstance = "Close"
    End Sub

    Private Function IsAssemblyDocument() As Boolean

        Dim res As Boolean = False
        Try
            objDocument = objApp.ActiveDocument

            If objDocument.Type = SolidEdgeFramework.DocumentTypeConstants.igAssemblyDocument Then
                res = True
            End If
        Catch ex As Exception
            MessageBox.Show($"Error in checking the assembly document ", "Error")

            CustomLogUtil.Log("in checking the assembly document", ex.Message, ex.StackTrace)
        End Try

        Return res

    End Function

    Private Sub BtnGetData_Click(sender As Object, e As EventArgs) Handles btnGetData.Click

        If Not IsAssemblyDocument() Then
            MsgBox("Please open assembly document")
            Exit Sub
        Else
            objAssemblyDocument = objApp.ActiveDocument
        End If

        If IO.File.Exists(txtExcelPath.Text) Then
            dicData = ExcelUtil.ReadMaterials(txtExcelPath.Text)
            Dim ds As DataSet = dicData("Structure")
            dtstructure = ds.Tables(0)
            Dim ds2 As DataSet = dicData("SheetMetal")
            dtsheetmetal = ds2.Tables(0)

            WaitStartSave()

            FillCategory()
        Else
            MessageBox.Show("Please select valid excelpath for bec material")
        End If

        SetControlBrowseExcel()

        dtAssemblyData = Nothing

        dtAssemblyData = SetAssemblyDetails(dtAssemblyData, objAssemblyDocument)

        dgvDocumentDetails.DataSource = dtAssemblyData

        SetControlGetAssemblyData()

        FillAssemblyDataExcelBased()

        WaitEndSave()
    End Sub

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

    Public Function SetAssemblyDetails(ByVal dtAssemblyData As DataTable, ByVal objAssemblyDocument As AssemblyDocument)

        Try

            SetDocumentNames()

            dtAssemblyData = AddColumns(dtAssemblyData)

            Dim parentDocPath As String = objAssemblyDocument.FullName

            Dim objOccurrences As SolidEdgeAssembly.Occurrences = objAssemblyDocument.Occurrences

            Dim dicDocumentDetails As New Dictionary(Of String, CustomProperties)()

            dicDocumentDetails = GetAllDocumentsDetails(parentDocPath, objOccurrences, dicDocumentDetails)

            dtAssemblyData = GetDocumentDetailsDatTable(dicDocumentDetails, dtAssemblyData)
        Catch ex As Exception

            MessageBox.Show($"error in set assembly details", "Error")
            CustomLogUtil.Log("in set assembly details", ex.Message, ex.StackTrace)
        End Try

        Return dtAssemblyData

    End Function

    Private Function GetDocumentDetailsDatTable(ByVal dicDocumentDetails As Dictionary(Of String, CustomProperties), ByVal dtAssemblyData As DataTable) As DataTable

        Try

            For Each kvp As KeyValuePair(Of String, CustomProperties) In dicDocumentDetails

                Dim custProp As CustomProperties = kvp.Value

                Dim dr As DataRow = dtAssemblyData.NewRow()

                If custProp.documentPath.ToUpper().EndsWith(".ASM") Then
                    Continue For
                End If

                If custProp.parentDocumentPath = String.Empty Then
                    Continue For
                End If

                dr(AssemblyColumns.ParentDocumentName.ToString()) = IO.Path.GetFileName(custProp.parentDocumentPath)
                dr(AssemblyColumns.DocumentName.ToString()) = IO.Path.GetFileName(custProp.documentPath)
                dr(AssemblyColumns.DocumentPath.ToString()) = custProp.documentPath
                dr(AssemblyColumns.PartType.ToString()) = custProp.partType
                dr(AssemblyColumns.Size.ToString()) = custProp.size
                dr(AssemblyColumns.Grade.ToString()) = custProp.grade
                dr(AssemblyColumns.GageName.ToString()) = custProp.gageName
                dr(AssemblyColumns.MaterialThickness.ToString()) = custProp.materialThickness
                dr(AssemblyColumns.BendRadius.ToString()) = custProp.bendRadius
                dr(AssemblyColumns.MaterialUsed.ToString()) = custProp.materialUsed
                dr(AssemblyColumns.MaterialSpec.ToString()) = custProp.materialSpec
                dr(AssemblyColumns.BECMaterial.ToString()) = custProp.materialName

                dtAssemblyData.Rows.Add(dr)

            Next
        Catch ex As Exception
            CustomLogUtil.Log("while Creating document Details Dat Table", ex.Message, ex.StackTrace)
        End Try

        Return dtAssemblyData
    End Function

    Private Function GetMaterialName(ByVal isSheetMetalPart As Boolean, ByRef objOccu As Occurrence) As String
        Dim materialName As String = String.Empty

        If isSheetMetalPart Then

            Dim sheetMetalDoc As SolidEdgePart.SheetMetalDocument = objOccu.OccurrenceDocument

            objMatTable.GetCurrentMaterialName(sheetMetalDoc, materialName)
        Else

            Dim partDoc As SolidEdgePart.PartDocument = objOccu.OccurrenceDocument

            objMatTable.GetCurrentMaterialName(partDoc, materialName)

        End If

        Return materialName
    End Function

    'Dim currentGaugeName As String = String.Empty
    '        objMatTable.GetCurrentGageName(objSheetMetalDocument, currentGaugeName)

    Private Function GetGageName(ByVal isSheetMetalPart As Boolean, ByRef objOccu As Occurrence) As String
        Dim currentGaugeName As String = String.Empty

        If isSheetMetalPart Then

            Dim sheetMetalDoc As SolidEdgePart.SheetMetalDocument = objOccu.OccurrenceDocument

            objMatTable.GetCurrentGageName(sheetMetalDoc, currentGaugeName)
        Else

            Dim partDoc As SolidEdgePart.PartDocument = objOccu.OccurrenceDocument

            objMatTable.GetCurrentGageName(partDoc, currentGaugeName)

        End If

        Return currentGaugeName
    End Function

    Private Function IsSheetMetalOcc(ByVal documentName As String) As Boolean

        Dim isSheetMetalPart As Boolean = False

        If documentName.ToUpper.EndsWith(".PSM") Then
            isSheetMetalPart = True
        End If

        Return isSheetMetalPart
    End Function

    Private Function IsAssemblyOcc(ByVal documentName As String) As Boolean

        Dim isAssemblyDoc1 As Boolean = False

        If documentName.ToUpper.EndsWith(".ASM") Then
            isAssemblyDoc1 = True
        End If

        Return isAssemblyDoc1
    End Function

    Private Function GetAllDocumentsDetails(ByVal parentDocPath As String, ByRef objOccurrences As SolidEdgeAssembly.Occurrences, ByRef dicDocumentDetails As Dictionary(Of String, CustomProperties))

        Dim errSb As New StringBuilder()
        For Each objOccu As Occurrence In objOccurrences

            Dim objDocument As Object = objOccu.OccurrenceDocument

            Try

                Dim docFullPath As String = objDocument.FullName

                Dim documentName As String = IO.Path.GetFileName(docFullPath)

                If dicDocumentDetails.ContainsKey(docFullPath) Then
                    Continue For
                End If

                Dim custPropertiesObj As New CustomProperties()

                If Not IsAssemblyOcc(documentName) Then

                    Dim isSheetMetalPart As Boolean = IsSheetMetalOcc(documentName)

                    Dim materialName As String = GetMaterialName(isSheetMetalPart, objOccu)

                    custPropertiesObj = ReadCustomProperties(isSheetMetalPart, objDocument)

                    custPropertiesObj.documentPath = docFullPath
                    custPropertiesObj.parentDocumentPath = parentDocPath
                    custPropertiesObj.materialName = materialName

                    Dim gageName As String = GetGageName(isSheetMetalPart, objOccu)
                    custPropertiesObj.gageName = gageName

                    dicDocumentDetails.Add(docFullPath, custPropertiesObj)
                Else

                    dicDocumentDetails.Add(docFullPath, custPropertiesObj)

                    'Dim asmDoc2 As AssemblyDocument = objOccu.OccurrenceDocument

                    'GetAllDocumentsDetails(asmDoc2.Path, asmDoc2.Occurrences, dicDocumentDetails)

                End If
            Catch ex As Exception
                errSb.AppendLine($"{parentDocPath},{ex.Message},{ex.StackTrace}")
                CustomLogUtil.Log($"while fetching document path: {parentDocPath}", ex.Message, ex.StackTrace)
            End Try
        Next

        If errSb.Length > 0 Then
            MsgBox(errSb.ToString())
        End If

        Return dicDocumentDetails
    End Function

    Public Function ReadCustomProperties(ByVal isSheetMetalPart As Boolean, ByRef objDocument As Object) As CustomProperties
        Dim custPropertiesObj = New CustomProperties()
        Try

            Dim objSheetMetalDocument As SheetMetalDocument = Nothing
            Dim objPartDocument As PartDocument = Nothing
            Dim propSets As SolidEdgeFramework.PropertySets = Nothing

            If isSheetMetalPart Then

                objSheetMetalDocument = objDocument
                propSets = objSheetMetalDocument.Properties
            Else

                objPartDocument = objDocument
                propSets = objPartDocument.Properties

            End If

            Dim custProps As Properties = propSets.Item("Custom")

            For Each prop1 As [Property] In custProps

                Try

                    If prop1.Name = "Size" Then

                        custPropertiesObj.size = prop1.Value

                    ElseIf prop1.Name = "Density" Then

                        custPropertiesObj.density = prop1.Value

                    ElseIf prop1.Name = "Accuracy" Then

                        custPropertiesObj.accuracy = prop1.Value

                    ElseIf prop1.Name = "Material Used" Then

                        custPropertiesObj.materialUsed = prop1.Value

                    ElseIf prop1.Name = "Teamcenter Item Type" Then

                        custPropertiesObj.teamCenterItemType = prop1.Value

                    ElseIf prop1.Name = "Part Type" Then

                        custPropertiesObj.partType = prop1.Value

                    ElseIf prop1.Name = "Last Saved Version" Then

                        custPropertiesObj.lastSavedVersion = prop1.Value

                    ElseIf prop1.Name = "MATL SPEC" Then

                        custPropertiesObj.materialSpec = prop1.Value

                    ElseIf prop1.Name = "Grade" Then

                        custPropertiesObj.grade = prop1.Value

                    ElseIf prop1.Name = "Material Thickness" Then

                        custPropertiesObj.materialThickness = prop1.Value

                    ElseIf prop1.Name = "Bend Radius" Then

                        'custPropertiesObj.bendRadius = prop1.Value.ToString()

                    ElseIf prop1.Name = "Gage" Then

                        custPropertiesObj.gageName = prop1.Value.ToString()

                    ElseIf prop1.Name = "Gage Table" Then

                        custPropertiesObj.bendRadius = prop1.Value.ToString()

                    End If
                Catch ex As Exception
                End Try

            Next
#Region "Variable Table"'TEMP12SEPT2023
            Dim objVariables As SolidEdgeFramework.Variables = Nothing
            Dim objVariable As SolidEdgeFramework.variable = Nothing
            objVariables = objSheetMetalDocument.Variables
            Dim value As Double = 0

            For Each objVariable In objVariables
                Try

                    If objVariable.DisplayName.Contains("RadiusGlobal") Then
                        Dim myvalue As String = Nothing
                        objVariable.GetValue(myvalue)
                        myvalue = Convert.ToDouble(myvalue.Trim.Replace("in", ""))
                        If value = 0 Or myvalue < value Then
                            value = myvalue
                            custPropertiesObj.bendRadius = value
                        End If
                    End If

                Catch ex As Exception

                End Try
            Next
            If value = 0 Then

                For Each objVariable In objVariables
                    Try

                        If objVariable.DisplayName.Contains("BendRadius") Then
                            Dim myvalue As String = Nothing
                            objVariable.GetValue(myvalue)
                            myvalue = Convert.ToDouble(myvalue.Trim.Replace("in", ""))
                            If value = 0 Or myvalue < value Then
                                value = myvalue
                                custPropertiesObj.bendRadius = value
                            End If
                        End If

                    Catch ex As Exception

                    End Try
                Next
            End If
            custPropertiesObj.bendRadius += " in"
#End Region
        Catch ex As Exception
            MessageBox.Show($"Error in reading custom properties ", "Error")
            CustomLogUtil.Log("in reading custom properties", ex.Message, ex.StackTrace)
        End Try

        Return custPropertiesObj
    End Function

    Public Function AddColumns(ByVal dtComments As DataTable) As DataTable

        dtComments = New DataTable("Assembly Data")

        'Dim myDataColumn1 As DataColumn = New DataColumn()
        'myDataColumn1 = New DataColumn()
        'myDataColumn1.ColumnName = "Select"
        'myDataColumn1.DefaultValue = "0"
        'myDataColumn1.DataType = System.Type.GetType("System.Boolean")
        'dtComments.Columns.Add(myDataColumn1)

        Dim column As New DataColumn With {
            .DataType = System.Type.[GetType]("System.Int32"),
            .Caption = "Sr",
            .ColumnName = "Sr",
            .AutoIncrement = True,
            .AutoIncrementSeed = 1,
            .AutoIncrementStep = 1
        }
        dtComments.Columns.Add(column)

        Dim commentsColumnsNames As Array
        commentsColumnsNames = System.Enum.GetNames(GetType(AssemblyColumns))
        Dim columnName As String
        For Each columnName In commentsColumnsNames
            dtComments.Columns.Add(columnName, GetType(String))
        Next

        'Dim myDataColumn As DataColumn = New DataColumn()
        'myDataColumn = New DataColumn()
        'myDataColumn.ColumnName = "IsEdited"
        'myDataColumn.DefaultValue = "0"
        'myDataColumn.DataType = System.Type.GetType("System.Boolean")
        'dtComments.Columns.Add(myDataColumn)

        Dim myDataColumn1 As New DataColumn With {
            .ColumnName = "IsValidData",
            .DefaultValue = "0",
            .DataType = System.Type.GetType("System.Boolean")
        }
        dtComments.Columns.Add(myDataColumn1)

        Return dtComments
    End Function

    Private Sub BtnOpenDocument_Click(sender As Object, e As EventArgs) Handles btnOpenDocument.Click

        OpenDocument()

    End Sub

    Private Sub OpenDocument()

        Try
            Dim rInd As Integer = dgvDocumentDetails.CurrentCell.RowIndex

            Dim dt As DataTable = dgvDocumentDetails.DataSource

            'Dim x As Integer
            Dim docPath As String = dgvDocumentDetails.Rows(rInd).Cells(AssemblyColumns.DocumentPath.ToString()).Value

            ' Dim docPath As String = dt.Rows(rInd)(AssemblyColumns.DocumentPath.ToString()).ToString()

            objApp.Documents.Open(docPath)
        Catch ex As Exception
            MessageBox.Show($"Error in open document ", "Error")

            CustomLogUtil.Log("in open document", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub BtnBrowseExcel_Click(sender As Object, e As EventArgs) Handles btnBrowseExcel.Click
        Try
            Using dialog As New OpenFileDialog
                dialog.Filter = "Excel files (*.xlsx)|*.xlsx"
                If dialog.ShowDialog() <> DialogResult.OK Then Return
                txtExcelPath.Text = dialog.FileName
            End Using

            'If IO.File.Exists(txtExcelPath.Text) Then

            '    If GlobalEntity.dictRawMaterials.Count = 0 Then
            '        GlobalEntity.dictRawMaterials = ExcelUtil.ReadRawMaterials2(txtExcelPath.Text)
            '    End If
            '    dt = GlobalEntity.dictRawMaterials("BECMaterials").Tables(0)
            '    FillCategory(dt, True)

            '    Dim ds As DataSet = GlobalEntity.dictRawMaterials("Structure")
            '    dtstructure = ds.Tables(0)
            '    Dim ds2 As DataSet = GlobalEntity.dictRawMaterials("SheetMetal")
            '    dtsheetmetal = ds2.Tables(0)

            'Else
            '    MsgBox("Please select material excel details.")
            'End If
        Catch ex As Exception

            CustomLogUtil.Log("While fatching BEC Material Excel Path", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub FillAssemblyDataExcelBased()

        If IO.File.Exists(txtExcelPath.Text) Then

            If GlobalEntity.dictRawMaterials.Count = 0 Then
                GlobalEntity.dictRawMaterials = ExcelUtil.ReadRawMaterials2(txtExcelPath.Text)
            End If

            dt = dtsheetmetal.Copy 'GlobalEntity.dictRawMaterials("BECMaterials").Tables(0)

            FillCategory(dt, True)

            SetControlBrowseExcel()

            Dim dtAssemblyData As DataTable = dgvDocumentDetails.DataSource
            Dim dtExcelData As DataTable = dt 'GlobalEntity.dictRawMaterials("BECMaterials").Tables(0)
            ValidateAssemblyData(dtAssemblyData, dtExcelData)

            FillSearchCombo()
        Else
            MsgBox("Please select material excel details.")
        End If

    End Sub

    Private Sub FillSearchCombo()
        ComboBoxFields.Items.Clear()
        ComboBoxFields.Items.Add("Select")

        For Each columns As System.Windows.Forms.DataGridViewColumn In dgvDocumentDetails.Columns
            ComboBoxFields.Items.Add(columns.Name.ToString())
        Next
        If ComboBoxFields.Items.Count > 0 Then
            ComboBoxFields.SelectedItem = ComboBoxFields.Items(0) 'ComboBoxFields.Items.Count - 1)
        End If
    End Sub

    Public Sub FillCategory(Optional ByVal dt As DataTable = Nothing, Optional ByVal isSheetMetalPart As Boolean = True)

        Try
            cmbCategory.Items.Clear()
            Dim categoryList As New List(Of String) From {
                "Structure", "SheetMetal"}


            For Each categoryName As String In categoryList

                If Not cmbCategory.Items.Contains(categoryName) Then
                    cmbCategory.Items.Add(categoryName)
                End If

            Next

            If cmbCategory.Items.Count > 0 Then
                cmbCategory.SelectedItem = cmbCategory.Items(0)

            End If

            If isSheetMetalPart Then
                cmbCategory.Text = "SheetMetal"
            Else
                cmbCategory.Text = "Structure"
            End If
        Catch ex As Exception
            MessageBox.Show($"Error in fill category ", "Error")
            CustomLogUtil.Log("in fill category", ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub DgvDocumentDetails_SelectionChanged(sender As Object, e As EventArgs) Handles dgvDocumentDetails.SelectionChanged

        If dgvDocumentDetails.CurrentCell Is Nothing Then
            Exit Sub
        End If

        Try
            btnApply.Enabled = False
            btnClose.Enabled = False
            dgvDocumentDetails.Enabled = False
            btnRefresh.Enabled = False

            Dim rInd As Integer = dgvDocumentDetails.CurrentCell.RowIndex
            Dim srNo As Integer = Integer.Parse(dgvDocumentDetails.Rows(rInd).Cells("Sr").Value.ToString())

            ResetCurrentPartDetails()

            SetSelectedPartData(srNo - 1)

            Validation()
        Catch ex As Exception
        Finally
            btnApply.Enabled = True
            btnClose.Enabled = True
            dgvDocumentDetails.Enabled = True
            btnRefresh.Enabled = True
        End Try

    End Sub

    Private Sub Validation()
        Try
            If cmbMaterialUsed2_Mw.Text = txtMaterialUsed_C.Text Then
                txtMaterialUsed_C.BackColor = Color.White
                txtMaterialUsed_C.ForeColor = Color.Black
            Else
                txtMaterialUsed_C.BackColor = Color.FromArgb(255, 199, 206)
                txtMaterialUsed_C.ForeColor = Color.FromArgb(156, 0, 6)
            End If

            If txtThickness2_Mw.Text = txtThickness_C.Text Then
                txtThickness_C.BackColor = Color.White
                txtThickness_C.ForeColor = Color.Black
            Else
                txtThickness_C.BackColor = Color.FromArgb(255, 199, 206)
                txtThickness_C.ForeColor = Color.FromArgb(156, 0, 6)
            End If

            If txtBECMaterial2_Mw.Text = txtCurrentMaterial_C.Text Then
                txtCurrentMaterial_C.BackColor = Color.White
                txtCurrentMaterial_C.ForeColor = Color.Black
            Else
                txtCurrentMaterial_C.BackColor = Color.FromArgb(255, 199, 206)
                txtCurrentMaterial_C.ForeColor = Color.FromArgb(156, 0, 6)
            End If
            Dim ListOfCurrentParts As New List(Of TextBox)
            ListOfCurrentParts.Clear()
            ListOfCurrentParts.Add(txtMaterialUsed_C)
            ListOfCurrentParts.Add(txtSize_C)
            ListOfCurrentParts.Add(txtGrade_C)
            ListOfCurrentParts.Add(txtGageName_C)
            ListOfCurrentParts.Add(txtThickness_C)
            ListOfCurrentParts.Add(txtPartType_C)
            ListOfCurrentParts.Add(txtBendRadius_C)
            ListOfCurrentParts.Add(txtMaterialSpec_C)
            ListOfCurrentParts.Add(txtCurrentMaterial_C)
            For i = 0 To ListOfCurrentParts.Count - 1
                Dim part As String = ListOfCurrentParts.Item(i).Text

                If part.Contains("Missing") Or part = "" Or part.Contains("Not") Or part.Contains("Not".ToUpper) Or part.Contains("Not".ToLower) Then
                    ListOfCurrentParts.Item(i).BackColor = Color.FromArgb(255, 199, 206)
                    ListOfCurrentParts.Item(i).ForeColor = Color.FromArgb(156, 0, 6)
                Else
                    ListOfCurrentParts.Item(i).BackColor = Color.White
                    ListOfCurrentParts.Item(i).ForeColor = Color.Black
                End If
            Next
        Catch ex As Exception
            MessageBox.Show($"Error in checking the assembly document", "Error")
            CustomLogUtil.Log("in checking the assembly document", ex.Message, ex.StackTrace)
        End Try
        BendRadiusColorValidate()
    End Sub

    Private Sub SetSelectedPartData(ByVal rInd As Integer)
        Try

            'rbCurrentDetails.Checked = True
            'rbCurrentDetails.PerformClick()
            'rInd -= 1
            Dim dtAssemblyData As DataTable = dgvDocumentDetails.DataSource

            Dim docPath As String = dtAssemblyData.Rows(rInd)(AssemblyColumns.DocumentPath.ToString()).ToString()

            Dim isSheetMetalPart As Boolean = IsSheetMetalOcc(docPath)

            If docPath.ToUpper().EndsWith(".ASM") Then

                'Do nothing
            Else

                txtSize_C.Text = dtAssemblyData.Rows(rInd)(AssemblyColumns.Size.ToString()).ToString()
                txtGrade_C.Text = dtAssemblyData.Rows(rInd)(AssemblyColumns.Grade.ToString()).ToString()
                txtGageName_C.Text = dtAssemblyData.Rows(rInd)(AssemblyColumns.GageName.ToString()).ToString()
                txtGageTable.Text = String.Empty ' dtAssemblyData.Rows(rInd)(AssemblyColumns.DocumentPath.ToString()).ToString()
                txtThickness_C.Text = dtAssemblyData.Rows(rInd)(AssemblyColumns.MaterialThickness.ToString()).ToString()
                txtBendRadius_C.Text = dtAssemblyData.Rows(rInd)(AssemblyColumns.BendRadius.ToString()).ToString()
                txtPartType_C.Text = dtAssemblyData.Rows(rInd)(AssemblyColumns.PartType.ToString()).ToString()
                txtMaterialUsed_C.Text = dtAssemblyData.Rows(rInd)(AssemblyColumns.MaterialUsed.ToString()).ToString()
                txtCurrentMaterial_C.Text = dtAssemblyData.Rows(rInd)(AssemblyColumns.BECMaterial.ToString()).ToString()
                txtMaterialSpec_C.Text = dtAssemblyData.Rows(rInd)(AssemblyColumns.MaterialSpec.ToString()).ToString()
                'Dim materialName As String = GetCurrentMaterialName(objSheetMetalDocument, objDocument, isSheetMetalPart)

                'Dim currentGageName As String = GetCurrentGageName(objSheetMetalDocument, objDocument, isSheetMetalPart)

                'Dim rInd As Integer = dgvDocumentDetails.CurrentCell.RowIndex

                'Dim dtAssemblyData As DataTable = dgvDocumentDetails.DataSource

                'Dim docPath As String = dtAssemblyData.Rows(rInd)(AssemblyColumns.DocumentPath.ToString()).ToString()

                FillCategory(dt, isSheetMetalPart)

            End If
        Catch ex As Exception
            'MsgBox($"Error in set selected part data {ex.Message}{vbNewLine}{ex.StackTrace}")
        End Try
    End Sub

    Private Sub RbCurrentDetails_CheckedChanged(sender As Object, e As EventArgs)

        'Exit Sub
        '' ResetCurrentPartDetails()

        ''SetMWControlsVisibility(rbCurrentDetails.Checked)

        'If rbCurrentDetails.Checked = True Then

        '    dgvDocumentDetails_SelectionChanged(sender, e)

        'End If

    End Sub

    Private Sub ResetCurrentPartDetails()
        txtMaterialSpec_C.Text = String.Empty
        txtSize_C.Text = String.Empty
        txtGrade_C.Text = String.Empty
        txtGageName_C.Text = String.Empty
        txtGageTable.Text = String.Empty
        txtThickness_C.Text = String.Empty
        txtBendRadius_C.Text = String.Empty
        txtPartType_C.Text = String.Empty
        txtMaterialSpec_C.Text = String.Empty
        txtCurrentMaterial_C.Text = String.Empty

    End Sub

    Private Sub RbMaterialWise_CheckedChanged(sender As Object, e As EventArgs)

        'If dgvDocumentDetails.CurrentCell Is Nothing Then
        '    Exit Sub
        'End If
        'Exit Sub

        'If rbMaterialWise.Checked Then

        '    Dim rInd As Integer = dgvDocumentDetails.CurrentCell.RowIndex

        '    Dim dtAssemblyData As DataTable = dgvDocumentDetails.DataSource

        '    Dim docPath As String = dtAssemblyData.Rows(rInd)(AssemblyColumns.DocumentPath.ToString()).ToString()

        '    Dim isSheetMetalPart As Boolean = IsSheetMetalOcc(docPath)

        '    FillCategory(dt, isSheetMetalPart)

        '    'SetMaterialWiseData(rInd)

        '    'SetMWControlsVisibility(True)

        'End If

    End Sub

    Private Sub SetMWControlsVisibility()

        'flag = True
        'If Not rbMaterialWise.Checked Then

        '    cmbMaterialUsed2_Mw.Enabled = True
        '    txtSize2_Mw.Enabled = True
        '    txtGrade2_Mw.Enabled = True
        '    txtGageName_Mw.Enabled = True
        '    txtThickness2_Mw.Enabled = True
        '    txtBendRadius_Mw.Enabled = True
        '    txtPartType_Mw.Enabled = True
        '    txtMaterialSpec2_Mw.Enabled = True
        '    txtBECMaterial2_Mw.Enabled = True

        '    txtMaterialUsed_C.Enabled = False
        '    txtSize_C.Enabled = False
        '    txtGrade_C.Enabled = False
        '    txtGageName_C.Enabled = False
        '    txtThickness_C.Enabled = False
        '    txtBendRadius_C.Enabled = False
        '    txtPartType_C.Enabled = False
        '    txtMaterialSpec_C.Enabled = False
        '    txtCurrentMaterial_C.Enabled = False

        'Else

        '    txtMaterialUsed_C.Enabled = True
        '    txtSize_C.Enabled = True
        '    txtGrade_C.Enabled = True
        '    txtGageName_C.Enabled = True
        '    txtThickness_C.Enabled = True
        '    txtBendRadius_C.Enabled = True
        '    txtPartType_C.Enabled = True
        '    txtMaterialSpec_C.Enabled = True
        '    txtCurrentMaterial_C.Enabled = True

        '    cmbMaterialUsed2_Mw.Enabled = False
        '    txtSize2_Mw.Enabled = False
        '    txtGrade2_Mw.Enabled = False
        '    txtGageName_Mw.Enabled = False
        '    txtThickness2_Mw.Enabled = False
        '    txtBendRadius_Mw.Enabled = False
        '    txtPartType_Mw.Enabled = False
        '    txtMaterialSpec2_Mw.Enabled = False
        '    txtBECMaterial2_Mw.Enabled = False
        '    'txtBECMaterial_Pw.Enabled = flag
        'End If

        'txtBECMaterial2_Mw.Visible = True

    End Sub

    Private Sub SetMaterialWiseData(ByVal rInd As Integer)
        Try
            Dim dt As DataTable = dgvDocumentDetails.DataSource

            Dim docPath As String = dt.Rows(rInd)(AssemblyColumns.DocumentPath.ToString()).ToString()

            If Not docPath.ToUpper().EndsWith(".ASM") Then
                cmbMaterialUsed2_Mw.Text = dt.Rows(rInd)(AssemblyColumns.MaterialUsed.ToString()).ToString()
            End If
        Catch ex As Exception
            MessageBox.Show($"Error in set material wise data", "Error")
            CustomLogUtil.Log("in set material wise data", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub CmbCategory_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCategory.SelectedIndexChanged
        Try
            CmbValidations()

            If cmbCategory.Text = "Structure" Then
                dt = dtstructure.Copy()
            ElseIf cmbCategory.Text = "SheetMetal" Then
                dt = dtsheetmetal.Copy()
            Else
                Exit Sub
            End If

            If cmbCategory.Text = "SheetMetal" Then
                cmbMaterialUsed2_Mw.Items.Clear()
                For i = 0 To dt.Rows.Count - 1
                    'MsgBox(dt.Rows(i)(4).ToString())
                    If Not cmbMaterialUsed2_Mw.Items.Contains(dt.Rows(i)(4).ToString()) Then
                        cmbMaterialUsed2_Mw.Items.Add(dt.Rows(i)(4).ToString())
                    End If
                Next
                cmbMaterialUsed2_Mw.Text = cmbMaterialUsed2_Mw.Items(0)
            End If

            FillMaterialUsed2(dt)

            If cmbMaterialUsed2_Mw.Items.Contains(txtMaterialUsed_C.Text) Then
                cmbMaterialUsed2_Mw.Text = txtMaterialUsed_C.Text
            End If

            'rbMaterialWise.Enabled = True
            'btnApply.Enabled = True
            'rbMaterialWise.Checked = True
        Catch ex As Exception
            MessageBox.Show($"Error in category selection change", "Error")
            CustomLogUtil.Log("in category selection change", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Public Sub CmbValidations()
        If cmbCategory.SelectedItem = "Structure" Then

            txtSize2_Mw.Enabled = False
            txtBendRadius_Mw.Enabled = False
            txtGageName_Mw.Enabled = False
            txtThickness2_Mw.Enabled = False
            cmbBendType_Mw.Enabled = False
        ElseIf cmbCategory.SelectedItem = "SheetMetal" Then

            txtSize2_Mw.Enabled = True
            txtBendRadius_Mw.Enabled = True
            txtGageName_Mw.Enabled = True
            txtThickness2_Mw.Enabled = True
            cmbBendType_Mw.Enabled = True
        End If
    End Sub

    Public Sub FillMaterialUsed2(ByVal dt As DataTable)

        Try

            cmbMaterialUsed2_Mw.Items.Clear()

            Dim materialUsedCol As String = ExcelUtil.ExcelSheetColumns.Material_Used.ToString()

            Dim categoryCol As String = ExcelUtil.ExcelSheetColumns.Category.ToString()
            Dim filter As String = String.Empty
            If cmbCategory.Text = "SheetMetal" Then
                'Dim value = "Sheet"
                Dim value = "SheetMetal"
                filter = $"Convert([{categoryCol}], 'System.String') = '{value}'"
            Else
                filter = $"Convert([{categoryCol}], 'System.String') = '{cmbCategory.Text}'"
            End If
            Dim dv As New DataView(dt) With {
                .RowFilter = filter'$"{categoryCol}='{cmbCategory.Text}'"
            }

            Dim materialUsedList As New List(Of String)()
            Dim dtSize As DataTable = dv.ToTable()
            For Each drv As DataRowView In dv

                Dim materialUsed1 As String = drv(ExcelUtil.ExcelSheetColumns.Material_Used.ToString)

                If Not materialUsedList.Contains(materialUsed1) Then
                    materialUsedList.Add(drv(ExcelUtil.ExcelSheetColumns.Material_Used.ToString))
                End If

            Next

            For Each materialUsedName As String In materialUsedList

                If Not cmbMaterialUsed2_Mw.Items.Contains(materialUsedName) Then
                    cmbMaterialUsed2_Mw.Items.Add(materialUsedName)
                End If

            Next

            If cmbMaterialUsed2_Mw.Items.Count > 0 Then
                cmbMaterialUsed2_Mw.SelectedItem = cmbMaterialUsed2_Mw.Items(0)
            End If

            Dim mySource As New AutoCompleteStringCollection()
            mySource.AddRange(materialUsedList.ToArray)
            cmbMaterialUsed2_Mw.AutoCompleteMode = AutoCompleteMode.SuggestAppend
            cmbMaterialUsed2_Mw.AutoCompleteSource = AutoCompleteSource.CustomSource
            cmbMaterialUsed2_Mw.AutoCompleteCustomSource = mySource
        Catch ex As Exception
            MessageBox.Show($"Error in fill material used2", "Error")
            CustomLogUtil.Log("in fill material used2", ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub CmbMaterialUsed2_Mw_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbMaterialUsed2_Mw.SelectedIndexChanged

        FillMaterialUsedWiseDetails(dt, cmbMaterialUsed2_Mw.Text)

        'SetMWControlsVisibility(rbCurrentDetails.Checked)

        FillBendType__AccordingToPriority()

        If dgvDocumentDetails.Rows.Count > 0 Then
            Validation()
        End If

    End Sub
    Private Sub FillBendType__AccordingToPriority()
        Try


            Dim typeCol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Type.ToString()
            Dim categoryCol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Category.ToString()
            Dim Becmaterialusedcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Used.ToString()
            Dim Thicknesscol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Thickness.ToString()
            Dim BECMaterialcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.BEC_Material.ToString()
            Dim GageTablecol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Gage_Table.ToString()
            Dim Becmaterialspecol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Specification.ToString()
            Dim Priorityspecol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Priority.ToString()
            Dim Priority As Integer = 0
            Dim value As Integer = 0
            Dim CategoryValue = "SheetMetal"
            If cmbCategory.Text = CategoryValue Then
                Dim dv As DataView = New DataView(dt)

                Dim filter As String = $"Convert([{Becmaterialusedcol}], 'System.String') = '{cmbMaterialUsed2_Mw.Text}'" ' $"Convert([{Becmaterialusedcol}], 'System.String') = '{cmbMaterialUsed2_Mw.Text}'and Convert([{Becmaterialspecol}], 'System.String')='{txtMaterialSpec2_Mw.Text}'and Convert([{BECMaterialcol}], 'System.String')='{txtBECMaterial2_Mw.Text}' and Convert([{Thicknesscol}], 'System.String')='{txtThickness2_Mw.Text}' "
                dv.RowFilter = filter

                cmbBendType_Mw.Items.Clear()

                For Each drv As DataRowView In dv
                    If Priority = 0 Then
                        Priority = Convert.ToInt32(drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Priority.ToString).ToString())
                    End If
                    value = Convert.ToInt32(drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Priority.ToString).ToString())


                    If (Priority > value) Then
                        Priority = value
                    End If
                    Dim BendType As String = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Bend_Type.ToString).ToString()
                    If Not cmbBendType_Mw.Items.Contains(BendType) Then
                        cmbBendType_Mw.Items.Add(BendType)
                    End If

                Next

                If dv.Count = 1 And cmbBendType_Mw.Items.Count > 0 Then
                    cmbBendType_Mw.SelectedItem = cmbBendType_Mw.Items(0)
                    'txtSize2_Mw.Text = dv(ExcelUtil.ExcelSheetColumns.Size.ToString).ToString()
                    'txtGrade2_Mw.Text = dv(ExcelUtil.ExcelSheetColumns.Grade.ToString).ToString()
                    'txtGageName_Mw.Text = dv(ExcelUtil.ExcelSheetColumns.Gage_Name.ToString).ToString()
                    'txtGageTable.Text = dv(ExcelUtil.ExcelSheetColumns.Gage_Table.ToString).ToString()
                    'txtThickness2_Mw.Text = dv(ExcelUtil.ExcelSheetColumns.Thickness.ToString).ToString()
                    'txtBendRadius_Mw.Text = dv(ExcelUtil.ExcelSheetColumns.Bend_Radius.ToString).ToString()
                    'txtPartType_Mw.Text = dv(ExcelUtil.ExcelSheetColumns.Type.ToString).ToString()
                    'txtMaterialSpec2_Mw.Text = dv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString).ToString()
                    'txtBECMaterial2_Mw.Text = dv(ExcelUtil.ExcelSheetColumns.BEC_Material.ToString).ToString()
                Else
                    dv.RowFilter = $"Convert([{Becmaterialusedcol}], 'System.String') = '{cmbMaterialUsed2_Mw.Text}' and {Priorityspecol}='{Priority}'" '$"{Becmaterialusedcol}='{cmbMaterialUsed2_Mw.Text}' and {Becmaterialspecol}='{txtMaterialSpec2_Mw.Text}'and {BECMaterialcol}='{txtBECMaterial2_Mw.Text}' and {Thicknesscol}='{txtThickness2_Mw.Text}' and {Priorityspecol}='{Priority}'"
                    For Each drv As DataRowView In dv
                        Dim BendType As String = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Bend_Type.ToString).ToString()


                        cmbBendType_Mw.SelectedItem = BendType
                        txtSize2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Size.ToString).ToString()
                        txtGrade2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Grade.ToString).ToString()
                        txtGageName_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Gage_Name.ToString).ToString()
                        txtGageTable.Text = drv(ExcelUtil.ExcelSheetColumns.Gage_Table.ToString).ToString()
                        txtThickness2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Thickness.ToString).ToString()
                        txtBendRadius_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Bend_Radius.ToString).ToString()
                        txtPartType_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Type.ToString).ToString()
                        txtMaterialSpec2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString).ToString()
                        txtBECMaterial2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.BEC_Material.ToString).ToString()
                    Next

                End If
            ElseIf Not cmbCategory.Text = CategoryValue Then
                Dim dv As New DataView(dt)

                Dim materialUsedCol As String = ExcelUtil.ExcelSheetColumns.Material_Used.ToString()
                Dim bendTypeCol As String = ExcelUtil.ExcelSheetColumns.Bend_Type.ToString()
                dv.RowFilter = $"Convert([{materialUsedCol}], 'System.String') ='{cmbMaterialUsed2_Mw.Text}'" '$"{materialUsedCol}='{cmbMaterialUsed2_Mw.Text}'And {bendTypeCol}='{cmbBendType_Mw.Text}'"

                For Each drv As DataRowView In dv

                    txtSize2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Size.ToString).ToString()
                    txtGrade2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Grade.ToString).ToString()

                    txtThickness2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Thickness.ToString).ToString()

                    txtPartType_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Type.ToString).ToString()
                    txtMaterialSpec2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString).ToString()
                    txtBECMaterial2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.BEC_Material.ToString).ToString()
                    Exit For

                Next
            End If
        Catch ex As Exception
            MessageBox.Show($"While fetching Bend Type", "Error")
            CustomLogUtil.Log("While fetching Bend Type", ex.Message, ex.StackTrace)
        End Try
    End Sub
    Private Sub FillMaterialUsedWiseDetails(ByVal dt As DataTable, ByVal materialUsed As String)

        Try

            txtSize2_Mw.Text = String.Empty
            txtGrade2_Mw.Text = String.Empty
            txtThickness2_Mw.Text = String.Empty
            txtPartType_Mw.Text = String.Empty
            txtMaterialSpec2_Mw.Text = String.Empty
            txtBECMaterial2_Mw.Text = String.Empty
            txtGageName_Mw.Text = String.Empty
            txtGageTable.Text = String.Empty
            txtBendRadius_Mw.Text = String.Empty
            cmbBendType_Mw.Items.Clear()
            Dim dv As New DataView(dt)

            Dim materialUsedCol As String = ExcelUtil.ExcelSheetColumns.Material_Used.ToString()
            'Temp18APR2023
            Dim filter As String = $"Convert([{materialUsedCol}], 'System.String') = '{cmbMaterialUsed2_Mw.Text}'"
            'Dim filter As String = $"Material_Used = '{(cmbMaterialUsed2_Mw.Text)}'"

            Try
                dv.RowFilter = filter '"Material_Used = 185-00030"
            Catch ex As Exception
                Debug.Print("aaa")
            End Try


            Dim BenType1 As String = String.Empty
            For Each drv As DataRowView In dv

                txtSize2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Size.ToString).ToString()
                txtGrade2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Grade.ToString).ToString()
                txtPartType_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Type.ToString).ToString()
                txtMaterialSpec2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString).ToString()
                txtBECMaterial2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.BEC_Material.ToString).ToString()
                txtThickness2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Thickness.ToString).ToString()



                If cmbCategory.Text = "SheetMetal" Then
                    txtGageName_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Gage_Name.ToString).ToString()
                    txtGageTable.Text = drv(ExcelUtil.ExcelSheetColumns.Gage_Table.ToString).ToString()
                    txtBendRadius_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Bend_Radius.ToString).ToString()

                    BenType1 = drv(ExcelUtil.ExcelSheetColumns.Bend_Type.ToString).ToString()
                    If Not cmbBendType_Mw.Items.Contains(BenType1) Then

                        cmbBendType_Mw.Items.Add(BenType1)

                    End If
                Else
                    Exit For
                End If



                'Exit For

            Next

            'If cmbBendType_Mw.Items.Count > 0 Then
            '    cmbBendType_Mw.SelectedItem = cmbBendType_Mw.Items(0)
            '    If cmbBendType_Mw.Items.Contains("") And cmbBendType_Mw.Items.Contains("NONE") Then
            '        cmbBendType_Mw.Items.Remove("")
            '        cmbBendType_Mw.Items.Remove("NONE")
            '    ElseIf cmbBendType_Mw.Items.Contains("") Then
            '        cmbBendType_Mw.Items.Remove("")
            '    ElseIf cmbBendType_Mw.Items.Contains("None") Then
            '        cmbBendType_Mw.Items.Remove("NONE")
            '    End If

            If cmbCategory.Text = "SheetMetal" Then
                cmbBendType_Mw.SelectedItem = cmbBendType_Mw.Items(0)
            End If

            'End If

        Catch ex As Exception
            MessageBox.Show($"Error in fill material used wise details", "Error")
            CustomLogUtil.Log("in fill material used wise details", ex.Message, ex.StackTrace)
        End Try
    End Sub
    Private Sub FillMaterialUsedWiseDetailsAfterBendType(ByVal dt As DataTable, ByVal materialUsed As String)

        Try

            txtSize2_Mw.Text = String.Empty
            txtGrade2_Mw.Text = String.Empty
            txtThickness2_Mw.Text = String.Empty
            txtPartType_Mw.Text = String.Empty
            txtMaterialSpec2_Mw.Text = String.Empty
            txtBECMaterial2_Mw.Text = String.Empty
            txtGageName_Mw.Text = String.Empty
            txtGageTable.Text = String.Empty
            txtBendRadius_Mw.Text = String.Empty

            Dim dv As New DataView(dt)

            Dim materialUsedCol As String = ExcelUtil.ExcelSheetColumns.Material_Used.ToString()
            Dim bendTypeCol As String = ExcelUtil.ExcelSheetColumns.Bend_Type.ToString()
            dv.RowFilter = $"{materialUsedCol}='{cmbMaterialUsed2_Mw.Text}'And {bendTypeCol}='{cmbBendType_Mw.Text}'"

            For Each drv As DataRowView In dv

                txtSize2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Size.ToString).ToString()
                txtGrade2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Grade.ToString).ToString()
                txtGageName_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Gage_Name.ToString).ToString()
                txtGageTable.Text = drv(ExcelUtil.ExcelSheetColumns.Gage_Table.ToString).ToString()
                txtThickness2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Thickness.ToString).ToString()
                txtBendRadius_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Bend_Radius.ToString).ToString()
                txtPartType_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Type.ToString).ToString()
                txtMaterialSpec2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString).ToString()
                txtBECMaterial2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.BEC_Material.ToString).ToString()
                Exit For

            Next
        Catch ex As Exception
            MessageBox.Show($"Error in fill material used wise details", "Error")
            CustomLogUtil.Log("in fill material used wise details", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub DgvDocumentDetails_DoubleClick(sender As Object, e As EventArgs) Handles dgvDocumentDetails.DoubleClick
        OpenDocument()
    End Sub

    Private Sub BtnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click

        If Not dictMaterials.ContainsKey(materialLib) Then
            MessageBox.Show($"BEC material library {materialLib} is not linked", "Material library", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        Dim rInd As Integer = dgvDocumentDetails.CurrentCell.RowIndex
        Dim doc As Object = objApp.ActiveDocument
        Dim docPath As String = dtAssemblyData.Rows(rInd)(AssemblyColumns.DocumentPath.ToString()).ToString()

        If Not docPath = doc.FullName Then

            'MsgBox($"Part document: {doc.FullName}  is different than {vbNewLine}{vbNewLine}selected document: {docPath}.{vbNewLine}Please select again")

            MessageBox.Show($"Active document and selected document from list are different{vbNewLine}{vbNewLine}Active    : {IO.Path.GetFileName(doc.FullName)}{vbNewLine}Selected: {IO.Path.GetFileName(docPath)}", "Wrong selection", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Dim isSheetMetalPart = IsSheetMetalDocument(docPath)

        ApplyPartDetails(docPath, isSheetMetalPart, dictMaterials)

    End Sub

    Private Sub ApplyMaterial(ByVal isSheetMetalPart As Boolean, ByVal dictMaterials As Dictionary(Of String, List(Of String)))
        Try

            Dim materialName As String = txtBECMaterial2_Mw.Text

            Dim lstString As List(Of String) = dictMaterials(materialLib)


            If lstString.Contains(materialName) Then

                ' Set active document handle

                If Not isSheetMetalPart Then
                    objDocument = objApp.ActiveDocument
                    objMatTable.SetActiveDocument(objDocument)

                    objMatTable.ApplyMaterialToDoc(objDocument, materialName, materialLib)
                Else

                    objSheetMetalDocument = objApp.ActiveDocument
                    objMatTable.SetActiveDocument(objSheetMetalDocument)

                    objMatTable.ApplyMaterialToDoc(objSheetMetalDocument, materialName, materialLib)
                End If
            Else
                MsgBox($"Material does not exist in {materialLib}")

            End If
        Catch ex As Exception
            CustomLogUtil.Log("Material does not exist in", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub ApplyPartDetails(ByVal docPath As String, ByVal isSheetMetalPart As Boolean, ByVal dictMaterials As Dictionary(Of String, List(Of String)))

        Dim currentCustProperties As CustomProperties = GetCurrentPartProperties()

        Dim newCustProperties As CustomProperties = GetNewPartProperties()

        ApplyMaterial(isSheetMetalPart, dictMaterials)

        ApplyCustomProperties(newCustProperties)

        Try
            If isSheetMetalPart Then

                'objMatTable.ApplyGageFromLibraryToDoc(objSheetMetalDocument, txtGageName_Mw.Text, cmbMaterialLib.Text)

                ''objMatTable.ApplyGageFromGageTableToDoc(objSheetMetalDocument, "20 Gage Al", "aluminum 6061")
                objMatTable.SetDocumentToGageTableAssociation(objSheetMetalDocument, newCustProperties.gageName, txtGageTable.Text, True, True)
            Else
                'objMatTable.ApplyGageFromGageTableToDoc(objDocument, txtGageName_Mw.Text, "Gagetable")
                ' objMatTable.SetActiveDocument(objDocument)
                'objMatTable.ApplyGageFromLibraryToDoc(objDocument, custPropertiesObj.gageName, cmbMaterialLib.Text)

                objMatTable.SetDocumentToGageTableAssociation(objDocument, newCustProperties.gageName, txtGageTable.Text, True, True)
                ' ApplyMaterial()
                'objMatTable.UpdateOODMaterialAndGageProperties(objDocument, True, True)

            End If

            ' ApplyMaterial()

            MsgBox("Process completed.")

            CustomLogUtil.Heading("AssemblyAutomationForm Process Done.....")
        Catch ex As Exception
            MessageBox.Show($"Error in apply", "Error")

            CustomLogUtil.Log("While Cliking On Apply Button", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub ApplyCustomProperties(ByVal custPropertiesObj As CustomProperties)
        Try

            If objDocument IsNot Nothing Then

                Dim propSets As SolidEdgeFramework.PropertySets = objDocument.Properties

                Dim custProps As Properties = propSets.Item("Custom")

                Dim lstProps As New List(Of String)()
                For Each prop1 As [Property] In custProps
                    lstProps.Add(prop1.Name)
                Next

                For Each pName As String In lstProps
                    If Not lstProps.Contains(pName) Then
                        custProps.Add(pName, String.Empty)
                    End If
                Next

                Try

                    If Not lstProps.Contains("JBL") Then
                        custProps.Add("JBLL", String.Empty)
                    End If
                    custProps.Save()
                Catch ex As Exception

                End Try

                For Each prop1 As [Property] In custProps

                    Try

                        If prop1.Name = "Size" Then

                            prop1.Value = custPropertiesObj.size

                        ElseIf prop1.Name = "Density" Then

                            prop1.Value = custPropertiesObj.density

                        ElseIf prop1.Name = "Accuracy" Then

                            prop1.Value = custPropertiesObj.accuracy

                        ElseIf prop1.Name = "Material Used" Then

                            prop1.Value = custPropertiesObj.materialUsed

                        ElseIf prop1.Name = "Teamcenter Item Type" Then

                            prop1.Value = custPropertiesObj.teamCenterItemType

                        ElseIf prop1.Name = "Part Type" Then

                            prop1.Value = custPropertiesObj.partType

                        ElseIf prop1.Name = "Last Saved Version" Then

                            prop1.Value = custPropertiesObj.lastSavedVersion

                        ElseIf prop1.Name = "MATL SPEC" Then

                            prop1.Value = custPropertiesObj.materialSpec

                        ElseIf prop1.Name = "Grade" Then

                            prop1.Value = custPropertiesObj.grade

                        ElseIf prop1.Name = "Material Thickness" Then

                            prop1.Value = custPropertiesObj.materialThickness

                        ElseIf prop1.Name = "Bend Radius" Then

                            prop1.Value = custPropertiesObj.bendRadius

                        End If
                    Catch ex As Exception
                    End Try

                Next
                propSets.Save()

            End If

            If objSheetMetalDocument IsNot Nothing Then

                Dim propSets As SolidEdgeFramework.PropertySets = objSheetMetalDocument.Properties

                Dim custProps As Properties = propSets.Item("Custom")

                For Each prop1 As [Property] In custProps

                    Try

                        If prop1.Name = "Size" Then

                            prop1.Value = custPropertiesObj.size

                        ElseIf prop1.Name = "Density" Then

                            prop1.Value = custPropertiesObj.density

                        ElseIf prop1.Name = "Accuracy" Then

                            prop1.Value = custPropertiesObj.accuracy

                        ElseIf prop1.Name = "Material Used" Then

                            prop1.Value = custPropertiesObj.materialUsed

                        ElseIf prop1.Name = "Teamcenter Item Type" Then

                            prop1.Value = custPropertiesObj.teamCenterItemType

                        ElseIf prop1.Name = "Part Type" Then

                            prop1.Value = custPropertiesObj.partType

                        ElseIf prop1.Name = "Last Saved Version" Then

                            prop1.Value = custPropertiesObj.lastSavedVersion

                        ElseIf prop1.Name = "MATL SPEC" Then

                            prop1.Value = custPropertiesObj.materialSpec

                        ElseIf prop1.Name = "Grade" Then

                            prop1.Value = custPropertiesObj.grade

                        ElseIf prop1.Name = "Material Thickness" Then

                            prop1.Value = custPropertiesObj.materialThickness

                        ElseIf prop1.Name = "Bend Radius" Then

                            prop1.Value = custPropertiesObj.bendRadius
                        End If
                    Catch ex As Exception
                    End Try

                Next
                propSets.Save()

            End If

            ' MsgBox("Property updation completed.")
        Catch ex As Exception
            CustomLogUtil.Log("While Applying Custom Properties", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Function IsSheetMetalDocument(ByVal docPath As String) As Boolean

        Dim isSheetMetalPart As Boolean = False

        If IO.Path.GetFileName(docPath).ToUpper.EndsWith(".PSM") Then

            isSheetMetalPart = True

        End If

        Return isSheetMetalPart

    End Function

    Private Function GetCurrentPartProperties() As CustomProperties
        Dim custProperties As New CustomProperties With {
            .size = txtSize_C.Text,
            .grade = txtGrade_C.Text,
            .gageName = txtGageName_C.Text,
            .gageTable = txtGageTable.Text,
            .materialThickness = txtThickness_C.Text,
            .bendRadius = txtBendRadius_C.Text,
            .partType = txtPartType_C.Text,
            .materialUsed = txtMaterialUsed_C.Text,
            .materialName = txtCurrentMaterial_C.Text,
            .materialSpec = txtMaterialSpec_C.Text,
            .category = cmbCategory.Text
        }

        Return custProperties
    End Function

    Private Function GetNewPartProperties() As CustomProperties
        Dim custProperties As New CustomProperties With {
            .size = txtSize2_Mw.Text,
            .grade = txtGrade2_Mw.Text,
            .gageName = txtGageName_Mw.Text,
            .gageTable = txtGageTable.Text,
            .materialThickness = txtThickness2_Mw.Text,
            .bendRadius = txtBendRadius_Mw.Text,
            .partType = txtPartType_Mw.Text,
            .materialUsed = cmbMaterialUsed2_Mw.Text,
            .materialName = txtBECMaterial2_Mw.Text,
            .materialSpec = txtMaterialSpec2_Mw.Text,
            .category = cmbCategory.Text,
            .bendType = cmbBendType_Mw.Text
        }

        Return custProperties
    End Function

    Private Sub AssemblyAutomationForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'System.Windows.Forms.Application.Exit()
    End Sub

    Private Sub BtnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click

        Try
            'ResetFrom()

            If Not IsAssemblyDocument() Then
                MsgBox("Please open assembly document")
                Exit Sub
            Else
                objAssemblyDocument = objApp.ActiveDocument
            End If

            dtAssemblyData = Nothing

            dtAssemblyData = SetAssemblyDetails(dtAssemblyData, objAssemblyDocument)

            dgvDocumentDetails.DataSource = dtAssemblyData
            Dim dtExcelData As DataTable = dt
            ValidateAssemblyData(dtAssemblyData, dtExcelData)
        Catch ex As Exception

        End Try
    End Sub
    Private Sub ResetFrom()
        'objApp = Nothing
        objDocument = Nothing
        objSheetMetalDocument = Nothing
        objAssemblyDocument = Nothing
        dtAssemblyData = New DataTable()
        objMatTable = Nothing
        dt = New DataTable()
        dictMaterials = New Dictionary(Of String, List(Of String))()
        listOfLibraries = Nothing
        dgvDocumentDetails.DataSource = Nothing

        cmbCategory.Items.Clear()
        cmbCategory.Text = String.Empty

        cmbMaterialUsed2_Mw.Items.Clear()
        cmbMaterialUsed2_Mw.Text = String.Empty

        txtExcelPath.Text = String.Empty
        txtSize2_Mw.Text = String.Empty
        txtGrade2_Mw.Text = String.Empty
        txtGageName_Mw.Text = String.Empty
        txtGageTable.Text = String.Empty
        txtThickness2_Mw.Text = String.Empty
        txtBendRadius_Mw.Text = String.Empty
        txtPartType_Mw.Text = String.Empty
        txtMaterialSpec2_Mw.Text = String.Empty
        txtBECMaterial2_Mw.Text = String.Empty

        lblAssemblyName.Text = String.Empty
        lblAssemblyPath.Text = String.Empty

        SetControlFirst()

        dtAssemblyData = AddColumns(dtAssemblyData)

        SetSolidEdgeInstance()

        SetMaterialTable()

        GetMaterialLibraryList()

        dictMaterials = GetMaterialCollection(listOfLibraries)

        SetControls(dgvDocumentDetails)
    End Sub
    Private Sub ComboBoxFields_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxFields.SelectedIndexChanged
        Try
            dgvDocumentDetails.DataSource = dtAssemblyData
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

    Public Sub ControlsEnability(ByVal flag As Boolean)

        btnOpenDocument.Enabled = flag
        cmbCategory.Enabled = flag
        btnRefresh.Enabled = flag
        btnApply.Enabled = flag
        cmbCategory.Enabled = flag
        cmbMaterialUsed2_Mw.Enabled = flag

    End Sub

    Private Sub BtnSearchFile_Click(sender As Object, e As EventArgs) Handles btnSearchFile.Click
        TestRemoveCode()
        TestRemoveCode()
    End Sub

    Private Sub TestRemoveCode()
        Try
            ControlsEnability(False)
            If txtSearch.Text.Trim = String.Empty Then
                ' dgvDocumentDetails.DataSource = Nothing
                dgvDocumentDetails.DataSource = dtAssemblyData
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
                Dim dt As DataTable = dtAssemblyData
                Dim DV As New DataView(dt)
                Try
                    DV.RowFilter = String.Format("" + ComboBoxFields.Text + " LIKE '%{0}%'", txtSearch.Text)
                Catch ex As Exception
                    DV.RowFilter = String.Format("" + ComboBoxFields.Text + "={0}", txtSearch.Text)
                End Try
                dgvDocumentDetails.DataSource = DV.ToTable()
            End If
        Catch ex As Exception
            MessageBox.Show("Unable to search.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CustomLogUtil.Log($"Unable to Search{ex.Message}{ex.StackTrace}")
        Finally

            ControlsEnability(True)
            SetGridValidationColor()

        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub cmbBendType_Mw_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbBendType_Mw.SelectedIndexChanged
        FillMaterialUsedWiseDetailsAfterBendType(dt, cmbMaterialUsed2_Mw.Text)
        If dgvDocumentDetails.Rows.Count > 0 Then
            Validation()
        End If
    End Sub

    Public Sub BendRadiusColorValidate() 'TEMP12SEPT2023
        Try
            If (Not txtBendRadius_C.Text = "" And Not txtBendRadius_Mw.Text = "") Then


                Dim CurrentBendRadius As Double = Convert.ToDouble(txtBendRadius_C.Text.Trim.Replace("in", ""))
                Dim ExcelBendRadius As Double = Convert.ToDouble(txtBendRadius_Mw.Text.Trim.Replace("in", ""))
                If CurrentBendRadius > ExcelBendRadius Then
                    txtBendRadius_C.BackColor = Color.FromArgb(255, 255, 0)
                    txtBendRadius_C.ForeColor = Color.Black
                ElseIf CurrentBendRadius < ExcelBendRadius Then
                    txtBendRadius_C.BackColor = Color.FromArgb(255, 199, 206)
                    txtBendRadius_C.ForeColor = Color.FromArgb(156, 0, 6)
                Else
                    txtBendRadius_C.BackColor = Color.White
                    txtBendRadius_C.ForeColor = Color.Black

                End If
            End If

        Catch ex As Exception

        End Try
    End Sub
End Class