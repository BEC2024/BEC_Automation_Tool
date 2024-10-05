Imports System.Runtime.InteropServices
Imports System.Text
'Imports SolidEdge.Framework.Interop
Imports SolidEdgePart
Imports WK.Libraries.BetterFolderBrowserNS
Imports SolidEdgeFramework
Imports SolidEdgeDraft
Imports Microsoft.Office.Interop
Imports SolidEdgeAssembly

Public Class AssemblyBomForm

    Public Enum AssemblyColumns

        'ParentDocumentName
        'DocumentName
        PartType
        'Size
        'Grade
        'GageName
        'MaterialThickness
        'BendRadius
        'MaterialUsed
        'MaterialSpec
        'BECMaterial
        'DocumentPath
        Material_Used
        Material_Spec
        Standard_ThicknessorLength
        Material
        Size
        Description
        ' Description2
        Document_Number
        Quantity
        Length 'Flat_Pattern_Model_CutSizeX

        'Total_CutSizeX
        Width 'Flat_Pattern_Model_CutSizeY

        'Total_CutSizeY
        Stock_Allowance

        Total_Length
        Total_Width
        Area
        OrderArea_Length
        OrderArea_LengthFT
        Extension
        'MATL_SPEC

        'Material_Thickness


        'Title
        Comments
        Document_Status
        Hardware

    End Enum

    Dim objApp As SolidEdgeFramework.Application = Nothing
    Dim objDocument As SolidEdgeFramework.SolidEdgeDocument = Nothing
    Dim objSheetMetalDocument As SheetMetalDocument = Nothing
    Dim objAssemblyDocument As SolidEdgeAssembly.AssemblyDocument
    Dim dtAssemblyData As DataTable = Nothing
    Dim objMatTable As SolidEdgeFramework.MatTable = Nothing
    Dim dt As New DataTable("")
    Dim dictMaterials As New Dictionary(Of String, List(Of String))()
    Dim listOfLibraries As Object = Nothing
    Dim numMaterials As Long
    Dim materialLib As String = "BEC MATERIAL LIBRARY"

    Dim dicQty As New Dictionary(Of String, Integer)()
    Dim dicAssemblyQty As New Dictionary(Of String, Integer)()
    Dim dicPartAssemblyList As New Dictionary(Of String, List(Of String))()

    Dim sbQty As New StringBuilder()
    Dim dtBOMInputData As New DataTable()

    Dim waitFormObj As Wait

    Dim assemblyNumber As Integer = 0
    Dim assemblyChildDocList As New Dictionary(Of String, List(Of String))()
    Dim ListOfColumns As New List(Of String)
    'Dim documentQty As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)()

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

#Region "Load"

    Private Function GetBOMData() As DataTable
        Dim filepath As String = txtBecMaterialExcel.Text '$"{System.IO.Directory.GetCurrentDirectory}\RawMaterialBOM.xlsx"
        'Dim dictBom As Dictionary(Of String, DataSet) = ExcelUtil.ReadRawMaterials2BOM(filepath) '"C:\Users\vimalb\Downloads\Bom_20220503-1.xlsx"

        'TEMP02MAY2023
        Dim dictBom As Dictionary(Of String, DataSet) = ExcelUtil.TestCreatePart(filepath)
        Dim Structure1 As DataTable = dictBom.Values(0).Tables(0).Copy
        Dim SheetMetal As DataTable = dictBom.Values(1).Tables(0).Copy
        dictBom.Clear()
        Dim ds As New DataSet

        'Structure1.Columns.Add(New DataColumn("Thickness1"))
        'Structure1.Columns("Thickness1").DataType = GetType(String)
        ds.Tables.Add(Structure1)
        SheetMetal.Columns.Add("Linear_Length")
        SheetMetal.Rows.RemoveAt(0)
        ds.Tables.Add(SheetMetal)
        'Dim ds As DataSet = dictBom.Values(0)

        ds = RemoveExtraColumnsFromDs(ds)

        Dim dt As DataTable = ds.Tables(1)
        dt.Merge(ds.Tables(0), True, MissingSchemaAction.Ignore)
        Return dt
    End Function
    Public Sub ListofExcelCloumnsForBOM()
        ListOfColumns.Clear()
        ListOfColumns.Add("Category")
        ListOfColumns.Add("Type")
        ListOfColumns.Add("Material_Used")
        ListOfColumns.Add("BEC_Material")
        ListOfColumns.Add("Size")
        ListOfColumns.Add("Stock_Clearance")
        ListOfColumns.Add("Thickness")
        ListOfColumns.Add("Linear_Length")
    End Sub
    Public Function RemoveExtraColumnsFromDs(ByVal ds As DataSet)
        ListofExcelCloumnsForBOM()

        For i = 0 To ds.Tables.Count - 1
            Dim count As Integer = ds.Tables(i).Columns.Count - 1
            For j = 0 To count
                If j > count Then
                    count = ds.Tables(i).Columns.Count - 1
                    j = 0
                    If count + 1 = ListOfColumns.Count Then
                        Exit For
                    End If
                End If
                Dim ColumnName As String = ds.Tables(i).Columns(j).ColumnName.ToString()
                If Not ListOfColumns.Contains(ColumnName) Then
                    ds.Tables(i).Columns.Remove(ColumnName)
                    count = ds.Tables(i).Columns.Count - 1
                End If

            Next
        Next
        Return ds
    End Function
    Private Function IsValid() As Boolean

        SetSolidEdgeInstance()
        DisableBtn()
        If objApp Is Nothing Then
            Return False
        Else
            Return True
        End If
    End Function
    Public Sub DisableBtn()
        If objApp Is Nothing Then
            btnClose.Enabled = False
            btnExportExcel.Enabled = False
            btnTemplateLocation.Enabled = False
            btnGetData.Enabled = False
            btnBrowseRawMaterialBOM.Enabled = False
        Else

            btnClose.Enabled = True
            btnExportExcel.Enabled = True
            btnTemplateLocation.Enabled = True
            btnGetData.Enabled = True
            btnBrowseRawMaterialBOM.Enabled = True
        End If
    End Sub

    Private Sub AssemblyBomForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CustomLogUtil.Heading("Raw Material Estimation Form Open.....")
        If IsValid() Then

        Else
            MessageBox.Show("Please Open Solid-Edge Assembly and Restart the Application", "Message")
            CustomLogUtil.Log("Please Open Solid-Edge Assembly and Restart the Application", "", "")
        End If
        'dtBOMInputData = GetBOMData()

        'SetSolidEdgeInstance()

        'SetMaterialTable()

        'GetMaterialLibraryList()

        'dictMaterials = GetMaterialCollection(listOfLibraries)

        'dtAssemblyData = AddColumns(dtAssemblyData)

        'SetControls(dgvDocumentDetails)


        txtRawMaterialEstimationReportDirPath.Text = Config.configObj.rawMaterialEstimationReportDirPath

        txtBecMaterialExcel.Text = Config.configObj.becMaterialExcelPath

    End Sub

    Private Sub SetSolidEdgeInstance()
        Try
            objApp = Marshal.GetActiveObject("SolidEdge.Application")
        Catch ex As Exception
            'MessageBox.Show($"Error in fetching the Solid-Edge instance {ex.Message} {vbNewLine} {ex.StackTrace}", "Message")
            'CustomLogUtil.Log("in Fetching the Solid-Edge Instance", ex.Message, ex.StackTrace)
        End Try

    End Sub

    Private Sub SetMaterialTable()
        Try
            objMatTable = objApp.GetMaterialTable()
        Catch ex As Exception
            MessageBox.Show($"Error in the Setting Material Tabl", "Error")
            CustomLogUtil.Log("in the Setting Material Table", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub GetMaterialLibraryList()
        Try
            objMatTable.GetMaterialLibraryList(listOfLibraries, numMaterials)
        Catch ex As Exception
            MessageBox.Show($"Error in fetching the Material Library List", "Error")
            CustomLogUtil.Log("in Fetching the Material Library List", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Function GetMaterialCollection(ByVal listOfLibraries As Object) As Dictionary(Of String, List(Of String))
        Dim dictMaterials As New Dictionary(Of String, List(Of String))()
        For Each libr As String In listOfLibraries

            Try
                Dim listOfMaterials1 As Object = Nothing
                Dim numMaterials1 As Long
                objMatTable.GetMaterialListFromLibrary(libr, numMaterials1, listOfMaterials1)
                If Not numMaterials1 = 0 And listOfMaterials1 IsNot Nothing Then
                    Dim lstMaterials As New List(Of String)()
                    For Each m1 As String In listOfMaterials1
                        lstMaterials.Add(m1)
                    Next
                    dictMaterials.Add(libr, lstMaterials)
                End If

            Catch ex As Exception
                MessageBox.Show($"While getting Material Collection", "Error")
                CustomLogUtil.Log($"While getting Material Collection", ex.Message, ex.StackTrace)
            End Try

        Next

        Return dictMaterials
    End Function

    Public Function AddColumns(dtComments As DataTable) As DataTable

        dtComments = New DataTable("Assembly Data")

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

        Return dtComments

    End Function

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

            DataGridViewComments.ReadOnly = True
        Catch ex As Exception
            MessageBox.Show($"While Setting the Controls", "Error")
            CustomLogUtil.Log($"While Setting the Controls", ex.Message, ex.StackTrace)
        End Try
    End Sub

#End Region

    Private Sub BtnGetData_Click(sender As Object, e As EventArgs) Handles btnGetData.Click

        WaitStartSave()

        'Move code from load here

        dtBOMInputData = GetBOMData()


        SetMaterialTable()

        GetMaterialLibraryList()

        dictMaterials = GetMaterialCollection(listOfLibraries)

        dtAssemblyData = AddColumns(dtAssemblyData)

        SetControls(dgvDocumentDetails)

        '=====

        btnGetData.Enabled = False
        btnExportExcel.Enabled = False

        If Not IsAssemblyDocument() Then

            MessageBox.Show("Please open assembly document", "Message")
            Exit Sub
        Else
            objAssemblyDocument = objApp.ActiveDocument
        End If

        dtAssemblyData = Nothing

        dtAssemblyData = GetAssemblyDetails(dtAssemblyData, objAssemblyDocument)

        dgvDocumentDetails.DataSource = dtAssemblyData

        btnGetData.Enabled = True
        btnExportExcel.Enabled = True

        WaitEndSave()

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
            CustomLogUtil.Log("While Checking the Assembly Document", ex.Message, ex.StackTrace)
        End Try

        Return res

    End Function

    Public Function GetSortingDt(ByVal dtAssemblyData As DataTable) As DataTable
        dtAssemblyData.DefaultView.Sort = $"{AssemblyBomForm.AssemblyColumns.Extension} Desc , {AssemblyBomForm.AssemblyColumns.Material_Used} ASC" ', {AssemblyBomForm.AssemblyColumns.Description.ToString()} DESC "
        dtAssemblyData = dtAssemblyData.DefaultView.ToTable

        'Dim materialUsedList As List(Of String) = From row In dtAssemblyData.AsEnumerable()
        '                                          Select row.Field(Of String)(AssemblyColumns.Material_Used.ToString()) Distinct.ToList()

        'Dim dv As DataView = New DataView(dtAssemblyData)
        'dv.RowFilter = $""

        Return dtAssemblyData
    End Function

    Public Function SetSrNo(ByVal dtAssemblyData As DataTable) As DataTable
        Try

            Dim rCnt As Integer = 1
            For Each dr As DataRow In dtAssemblyData.Rows
                dr("Sr") = rCnt
                rCnt += 1
            Next
        Catch ex As Exception

        End Try

        Return dtAssemblyData
    End Function

    Public Function UpdateProfileLength(ByVal dtAssemblyData As DataTable) As DataTable
        Try

            Dim materialUsedList As List(Of String) = From row In dtAssemblyData.AsEnumerable()
                                                      Select row.Field(Of String)(AssemblyColumns.Material_Used.ToString()) Distinct.ToList()

            For Each materialUsed As String In materialUsedList

                Dim rows As DataRow() = dtAssemblyData.Select($"{AssemblyColumns.Material_Used} ='{materialUsed}'")
                Dim totalArea As Double = 0

                For Each row As DataRow In rows

                    Try
                        totalArea += Double.Parse(row(AssemblyColumns.Area.ToString()))
                    Catch ex As Exception
                    End Try

                Next

                Dim totalLength As Double = 0
                For Each row As DataRow In rows

                    Try
                        totalLength += Double.Parse(row(AssemblyColumns.Total_Length.ToString()))
                    Catch ex As Exception
                    End Try

                Next


                Dim totalwidth As Double = 0
                For Each row As DataRow In rows

                    Try
                        totalwidth += Double.Parse(row(AssemblyColumns.Total_Width.ToString()))
                    Catch ex As Exception
                    End Try

                Next

                totalLength = Math.Round(totalLength, 3)
                totalArea = Math.Round(totalArea, 3)
                totalwidth = Math.Round(totalwidth, 3)
                Dim cnt1 As Integer = 1
                For Each row As DataRow In rows
                    If cnt1 = 1 Then

                        If (row(AssemblyColumns.Total_Width.ToString())).ToString.ToUpper() = "NA" Or (row(AssemblyColumns.Total_Width.ToString())).ToString.ToUpper() = "N/A" Or (row(AssemblyColumns.Total_Width.ToString())).ToString.ToUpper() = "" Then
                            row(AssemblyColumns.OrderArea_Length.ToString()) = totalLength.ToString()
                            row(AssemblyColumns.OrderArea_LengthFT.ToString()) = Math.Round(Convert.ToDouble(totalLength.ToString()) / 12, 3)
                            'ElseIf (row(AssemblyColumns.Total_Length.ToString())).ToString.ToUpper() = "NA" Or (row(AssemblyColumns.Total_Length.ToString())).ToString.ToUpper() = "N/A" Or (row(AssemblyColumns.Total_Length.ToString())).ToString.ToUpper() = "" Then
                            '    row(AssemblyColumns.OrderArea_Length.ToString()) = .ToString()
                        Else

                            If totalArea = 0 Then
                                totalArea = totalwidth
                            End If
                            row(AssemblyColumns.OrderArea_Length.ToString()) = totalArea.ToString()
                            row(AssemblyColumns.OrderArea_LengthFT.ToString()) = Math.Round(Convert.ToDouble(totalArea.ToString()) / 12, 3)
                        End If

                        'If totalArea = 0 Then
                        '    row(AssemblyColumns.OrderArea_Length.ToString()) = row(AssemblyColumns.OrderArea_Length.ToString()) = (Double.Parse(row(AssemblyColumns.Total_Length.ToString())) + Double.Parse(row(AssemblyColumns.Total_Width.ToString()))).ToString()
                        'End If
                    Else
                        row(AssemblyColumns.OrderArea_Length.ToString()) = String.Empty
                        row(AssemblyColumns.Material_Used.ToString()) = String.Empty
                        row(AssemblyColumns.Material.ToString()) = String.Empty
                        row(AssemblyColumns.Size.ToString()) = String.Empty
                    End If


                    cnt1 += 1
                Next

            Next

            'Dim cnt As Integer = 0
            'Dim profileLength As Double = 0
            'For Each dr As DataRow In dtAssemblyData.Rows

            '    Dim materialUsed As String = dr(AssemblyColumns.Material_Used.ToString())

            '    Dim totalLength As Double = dr(AssemblyColumns.Total_Length.ToString())

            '    Dim nextMaterialUsed As String = String.Empty

            '    If cnt + 1 < dtAssemblyData.Rows.Count Then
            '        profileLength = profileLength + totalLength
            '        nextMaterialUsed = dtAssemblyData.Rows(cnt + 1)(AssemblyColumns.Material_Used.ToString()).ToString()
            '    End If

            '    If Not materialUsed = nextMaterialUsed Then
            '        dr(AssemblyColumns.Profile_Length.ToString()) = profileLength.ToString
            '        profileLength = 0
            '    End If

            '    cnt = cnt + 1
            'Next


        Catch ex As Exception
            MessageBox.Show($"Error in Update Profile Length ", "Error")
            CustomLogUtil.Log("in Update Profile Length", ex.Message, ex.StackTrace)
        End Try
        Return dtAssemblyData
    End Function

    Public Function SetProfileLength(ByVal dtAssemblyData As DataTable) As DataTable
        Try
            For Each dr As DataRow In dtAssemblyData.Rows

                If dr(AssemblyColumns.Material_Used.ToString()) = String.Empty Then
                    Continue For
                End If
                Try
                    Dim totalLength As Double = Double.Parse(dr(AssemblyColumns.Total_Length.ToString()))
                    Dim totalWidth As Double = Double.Parse(dr(AssemblyColumns.Total_Width.ToString()))

                    If totalLength > 0 And totalWidth > 0 Then
                        dr(AssemblyColumns.OrderArea_Length.ToString()) = String.Empty
                    End If
                Catch ex As Exception

                End Try

            Next

        Catch ex As Exception
            MessageBox.Show($"Error in Set Profile Length ", "Error")
            CustomLogUtil.Log("in set Profile Length", ex.Message, ex.StackTrace)
        End Try
        Return dtAssemblyData
    End Function

    Public Function GetAssemblyDetails(ByVal dtAssemblyData As DataTable, ByVal objAssemblyDocument As AssemblyDocument)

        Try

            SetDocumentNames()

            dtAssemblyData = AddColumns(dtAssemblyData)

            Dim parentDocPath As String = objAssemblyDocument.FullName

            Dim objOccurrences As SolidEdgeAssembly.Occurrences = objAssemblyDocument.Occurrences

            Dim dicDocumentDetails As New Dictionary(Of String, CustomProperties)()

            'dicDocumentDetails = GetAllDocumentsDetailsNew1(parentDocPath, objOccurrences, dicDocumentDetails)

            '  dicDocumentDetails = GetAllDocumentsDetails(parentDocPath, objOccurrences, dicDocumentDetails)
            Dim dt As DataTable = Getpartlist()


            dicDocumentDetails = GetDocumentDetails(dt)

            dtAssemblyData = GetDocumentDetailsDatTable(dicDocumentDetails, dtAssemblyData)

            ' dtAssemblyData = GetSortingDt(dtAssemblyData)

            dtAssemblyData = SetSrNo(dtAssemblyData)

            dtAssemblyData = UpdateProfileLength(dtAssemblyData)

            'dtAssemblyData = SetProfileLength(dtAssemblyData)
            'Try
            '    Dim totalLength As Double
            '    Dim lastlength As Double
            '    Dim MaterialUsed As String

            '    Dim i As Integer = dtAssemblyData.Rows.Count
            '    For i = 0 To dtAssemblyData.Rows.Count Step 1
            '        '

            '        If dtAssemblyData.Rows(i)(AssemblyColumns.Material_Used.ToString()).ToString().Trim() = "185-00005" Then
            '            Debug.Print("aaaa")
            '        End If
            '        If dtAssemblyData.Rows(i)(AssemblyColumns.Material_Used.ToString()).ToString().Trim() = "185-00041" Then
            '            Debug.Print("aaaa")
            '        End If

            '        Dim currentMaterialUsed As String = String.Empty
            '        Try
            '            currentMaterialUsed = dtAssemblyData.Rows(i)(AssemblyColumns.Material_Used.ToString()).ToString().Trim()
            '        Catch ex As Exception

            '        End Try

            '        If MaterialUsed = String.Empty Then
            '            MaterialUsed = dtAssemblyData.Rows(i)(AssemblyColumns.Material_Used.ToString()).ToString().Trim()
            '            If i <> 0 Then
            '                Try
            '                    totalLength = Double.Parse(dtAssemblyData.Rows(i - 1)(AssemblyColumns.Total_Length.ToString()))
            '                Catch ex As Exception

            '                End Try

            '            End If

            '        End If

            '        If MaterialUsed = dtAssemblyData.Rows(i)(AssemblyColumns.Material_Used.ToString()).ToString().Trim() Then
            '            Try
            '                totalLength = totalLength + Double.Parse(dtAssemblyData.Rows(i)(AssemblyColumns.Total_Length.ToString()))
            '            Catch ex As Exception

            '            End Try

            '        End If

            '        If MaterialUsed <> dtAssemblyData.Rows(i)(AssemblyColumns.Material_Used.ToString()).ToString().Trim() Then
            '            dtAssemblyData.Rows(i - 1)(AssemblyColumns.Profile_Length.ToString()) = totalLength
            '            totalLength = 0
            '            MaterialUsed = String.Empty
            '        End If
            '        If dtAssemblyData.Rows.Count = i + 1 Then
            '            dtAssemblyData.Rows(i)(AssemblyColumns.Profile_Length.ToString()) = totalLength
            '        End If
            '    Next
            'Catch ex As Exception
            '    Debug.Print(ex.Message + vbNewLine + ex.StackTrace)
            'End Try
        Catch ex As Exception

            MessageBox.Show($"error in set assembly details", "Error")
            CustomLogUtil.Log("While setting Assembly Details", ex.Message, ex.StackTrace)

        End Try

        Return dtAssemblyData

    End Function

    Private Sub SetDocumentNames()

        lblAssemblyName.Text = IO.Path.GetFileName(objAssemblyDocument.FullName)
        ' lblAssemblyPath.Text = objAssemblyDocument.FullName

    End Sub
    Private Function GetAllDocumentsDetails(ByVal parentDocPath As String, ByRef objOccurrences As SolidEdgeAssembly.Occurrences, ByRef dicDocumentDetails As Dictionary(Of String, CustomProperties))

        Dim errSb As New StringBuilder()
        For Each objOccu As Occurrence In objOccurrences

            Dim objDocument As Object = objOccu.OccurrenceDocument

            Try

                Dim docFullPath As String = objDocument.FullName
                Dim documentName As String = IO.Path.GetFileName(docFullPath)
                Dim custPropertiesObj As New CustomProperties()

                If Not IsAssemblyOcc(documentName) Then

                    Dim isSheetMetalPart As Boolean = IsSheetMetalOcc(documentName)

                    'Dim materialName As String = GetMaterialName(isSheetMetalPart, objOccu)

                    custPropertiesObj = ReadCustomProperties(isSheetMetalPart, objDocument)

                    custPropertiesObj.documentPath = docFullPath
                    custPropertiesObj.parentDocumentPath = parentDocPath
                    'custPropertiesObj.materialName = materialName

                    'If dicQty.ContainsKey(docFullPath) Then
                    '    custPropertiesObj.quantity = dicQty(docFullPath)
                    'End If

                    Dim gageName As String = GetGageName(isSheetMetalPart, objOccu)

                    custPropertiesObj.gageName = gageName

                    If dicQty.ContainsKey(docFullPath) Then
                        Try
                            Dim cnt As Integer = dicQty(docFullPath)
                            dicQty(docFullPath) = cnt + 1
                        Catch ex As Exception
                        End Try
                    Else
                        dicDocumentDetails.Add(docFullPath, custPropertiesObj)
                        dicQty.Add(docFullPath, 1)

                    End If
                Else

                    If Not dicDocumentDetails.ContainsKey(docFullPath) Then

                        dicDocumentDetails.Add(docFullPath, custPropertiesObj)

                    End If

                    Dim asmDoc2 As AssemblyDocument = objOccu.OccurrenceDocument

                    GetAllDocumentsDetails(asmDoc2.Path, asmDoc2.Occurrences, dicDocumentDetails)

                End If
            Catch ex As Exception
                errSb.AppendLine($"{parentDocPath},{ex.Message},{ex.StackTrace}")
            End Try
        Next

        If errSb.Length > 0 Then
            MsgBox(errSb.ToString())
        End If

        Return dicDocumentDetails
    End Function
    Private Function GetAllDocumentsDetails_(ByVal parentDocPath As String, ByRef objOccurrences As SolidEdgeAssembly.Occurrences, ByRef dicDocumentDetails As Dictionary(Of String, CustomProperties))

        Dim errSb As New StringBuilder()

        For Each objOccu As Occurrence In objOccurrences

            Dim objDocument As Object = Nothing

            Try
                objDocument = objOccu.OccurrenceDocument
            Catch ex As Exception
            End Try

            If objDocument Is Nothing Then
                Continue For
            End If

            Try

                Dim docFullPath As String = objDocument.FullName
                Dim documentName As String = IO.Path.GetFileName(docFullPath)
                waitFormObj.SetProgressInformationMessage(documentName)

                Dim custPropertiesObj As New CustomProperties()

                If Not IsAssemblyOcc(documentName) Then

                    Dim isSheetMetalPart As Boolean = IsSheetMetalOcc(documentName)

                    'Dim materialName As String = GetMaterialName(isSheetMetalPart, objOccu)

                    custPropertiesObj = ReadCustomProperties(isSheetMetalPart, objDocument)

                    custPropertiesObj.documentPath = docFullPath
                    custPropertiesObj.parentDocumentPath = parentDocPath
                    'custPropertiesObj.materialName = materialName

                    Dim gageName As String = GetGageName(isSheetMetalPart, objOccu)

                    custPropertiesObj.gageName = gageName

                    If dicQty.ContainsKey(docFullPath) Then
                        Try
                            Dim cnt As Integer = dicQty(docFullPath)
                            dicQty(docFullPath) = cnt + 1
                        Catch ex As Exception
                        End Try
                    Else
                        dicDocumentDetails.Add(docFullPath, custPropertiesObj)
                        dicQty.Add(docFullPath, 1)
                    End If

                    '====================
                    If dicPartAssemblyList.ContainsKey(docFullPath) Then
                        Try
                            'Dim cnt As Integer = dicQty(docFullPath)
                            'dicQty(docFullPath) = cnt + 1

                            Dim assemblyPathList As List(Of String) = dicPartAssemblyList(docFullPath)
                            If Not assemblyPathList.Contains(parentDocPath) Then
                                assemblyPathList.Add(parentDocPath)
                            End If
                            dicPartAssemblyList(docFullPath) = assemblyPathList

                        Catch ex As Exception
                        End Try
                    Else
                        'dicDocumentDetails.Add(docFullPath, custPropertiesObj)
                        'dicQty.Add(docFullPath, 1)

                        Dim assemblyPathList As New List(Of String) From {parentDocPath}
                        'assemblyPathList.Add(parentDocPath)
                        dicPartAssemblyList.Add(docFullPath, assemblyPathList)

                    End If

                Else

                    assemblyNumber += 1
                    waitFormObj.SetProgressCountMessage(assemblyNumber.ToString)
                    waitFormObj.SetWaitMessage($"{IO.Path.GetFileName(docFullPath)} In progress..")

                    'Assembly property does not required in report

                    'If Not dicDocumentDetails.ContainsKey(docFullPath) Then
                    '    dicDocumentDetails.Add(docFullPath, custPropertiesObj)
                    'End If

                    If dicAssemblyQty.ContainsKey(docFullPath) Then

                        Try
                            Dim cnt As Integer = dicAssemblyQty(docFullPath)
                            dicAssemblyQty(docFullPath) = cnt + 1
                            Continue For

                        Catch ex As Exception
                        End Try
                    Else

                        'dicDocumentDetails.Add(docFullPath, custPropertiesObj)
                        dicAssemblyQty.Add(docFullPath, 1)

                    End If

                    Dim asmDoc2 As AssemblyDocument = objOccu.OccurrenceDocument
                    GetAllDocumentsDetails(asmDoc2.Path, asmDoc2.Occurrences, dicDocumentDetails)

                End If
            Catch ex As Exception

                errSb.AppendLine($"{parentDocPath},{ex.Message},{ex.StackTrace}")

            End Try
        Next

        If errSb.Length > 0 Then
            MsgBox(errSb.ToString())
        End If

        Return dicDocumentDetails
    End Function


    Dim dicAssembly_PartQty As New Dictionary(Of String, Dictionary(Of String, Integer))()

    Private Function GetAllDocumentsDetailsNew1(ByVal parentDocPath As String, ByRef objOccurrences As SolidEdgeAssembly.Occurrences, ByRef dicDocumentDetails As Dictionary(Of String, CustomProperties))

        Dim errSb As New StringBuilder()

        For Each objOccu As Occurrence In objOccurrences

            Dim objDocument As Object = Nothing
            Try
                objDocument = objOccu.OccurrenceDocument
            Catch ex As Exception
            End Try

            If objDocument Is Nothing Then
                Continue For
            End If

            Try

                Dim docFullPath As String = objDocument.FullName
                Dim documentName As String = IO.Path.GetFileName(docFullPath)
                waitFormObj.SetProgressInformationMessage(documentName)

                Dim custPropertiesObj As New CustomProperties()

                If Not IsAssemblyOcc(documentName) Then

                    Dim isSheetMetalPart As Boolean = IsSheetMetalOcc(documentName)
                    custPropertiesObj = ReadCustomProperties(isSheetMetalPart, objDocument)
                    custPropertiesObj.documentPath = docFullPath
                    custPropertiesObj.parentDocumentPath = parentDocPath
                    custPropertiesObj.gageName = GetGageName(isSheetMetalPart, objOccu)

                    If dicAssembly_PartQty.ContainsKey(parentDocPath) Then

                        Dim dicPartQty As Dictionary(Of String, Integer) = dicAssembly_PartQty(parentDocPath)
                        If dicPartQty.ContainsKey(docFullPath) Then
                            Try
                                Dim cnt As Integer = dicPartQty(docFullPath)
                                dicPartQty(docFullPath) = cnt + 1
                            Catch ex As Exception
                            End Try
                        Else
                            dicPartQty.Add(docFullPath, 1)
                        End If

                        dicAssembly_PartQty(parentDocPath) = dicPartQty

                    Else
                        Dim dicPartQty As New Dictionary(Of String, Integer) From {
                            {docFullPath, 1}
                        }
                        dicAssembly_PartQty.Add(parentDocPath, dicPartQty)
                    End If


                    If dicQty.ContainsKey(docFullPath) Then
                        Try
                            Dim cnt As Integer = dicQty(docFullPath)
                            dicQty(docFullPath) = cnt + 1
                        Catch ex As Exception
                        End Try
                    Else
                        dicDocumentDetails.Add(docFullPath, custPropertiesObj)
                        dicQty.Add(docFullPath, 1)
                    End If

                    '====================
                    If dicPartAssemblyList.ContainsKey(docFullPath) Then
                        Try
                            'Dim cnt As Integer = dicQty(docFullPath)
                            'dicQty(docFullPath) = cnt + 1

                            Dim assemblyPathList As List(Of String) = dicPartAssemblyList(docFullPath)
                            If Not assemblyPathList.Contains(parentDocPath) Then
                                assemblyPathList.Add(parentDocPath)
                            End If
                            dicPartAssemblyList(docFullPath) = assemblyPathList

                        Catch ex As Exception
                        End Try
                    Else
                        Dim assemblyPathList As New List(Of String) From {
                            parentDocPath
                        }
                        dicPartAssemblyList.Add(docFullPath, assemblyPathList)

                    End If

                Else

                    assemblyNumber += 1
                    waitFormObj.SetProgressCountMessage(assemblyNumber.ToString)
                    waitFormObj.SetWaitMessage($"{IO.Path.GetFileName(docFullPath)} In progress..")


                    If dicAssemblyQty.ContainsKey(docFullPath) Then

                        Try
                            Dim cnt As Integer = dicAssemblyQty(docFullPath)
                            dicAssemblyQty(docFullPath) = cnt + 1
                            Continue For
                        Catch ex As Exception
                        End Try
                    Else
                        dicAssemblyQty.Add(docFullPath, 1)
                    End If

                    Dim asmDoc2 As AssemblyDocument = objOccu.OccurrenceDocument
                    GetAllDocumentsDetailsNew1(asmDoc2.FullName, asmDoc2.Occurrences, dicDocumentDetails)

                End If
            Catch ex As Exception

                errSb.AppendLine($"{parentDocPath},{ex.Message},{ex.StackTrace}")

            End Try
        Next

        If errSb.Length > 0 Then
            MsgBox(errSb.ToString())
        End If

        Return dicDocumentDetails
    End Function
    Private Function IsAssemblyOcc(ByVal documentName As String) As Boolean

        Dim isAssemblyDoc1 As Boolean = False

        If documentName.ToUpper.EndsWith(".ASM") Then
            isAssemblyDoc1 = True
        End If

        Return isAssemblyDoc1
    End Function

    Private Function IsSheetMetalOcc(ByVal documentName As String) As Boolean

        Dim isSheetMetalPart As Boolean = False

        If documentName.ToUpper.EndsWith(".PSM") Then
            isSheetMetalPart = True
        End If

        Return isSheetMetalPart
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

    Private Function GetGageName(ByVal isSheetMetalPart As Boolean, ByRef objOccu As Occurrence) As String
        Dim currentGaugeName As String = String.Empty
        Try
            If isSheetMetalPart Then

                Dim sheetMetalDoc As SolidEdgePart.SheetMetalDocument = objOccu.OccurrenceDocument

                objMatTable.GetCurrentGageName(sheetMetalDoc, currentGaugeName)
            Else

                Dim partDoc As SolidEdgePart.PartDocument = objOccu.OccurrenceDocument

                objMatTable.GetCurrentGageName(partDoc, currentGaugeName)

            End If
        Catch ex As Exception

        End Try


        Return currentGaugeName
    End Function

    Public Function ReadCustomProperties(ByVal isSheetMetalPart As Boolean, ByRef objDocument As Object) As CustomProperties

        Dim custPropertiesObj = New CustomProperties()
        Try

            Dim objSheetMetalDocument As SheetMetalDocument = Nothing
            Dim objPartDocument As PartDocument = Nothing
            Dim propSets As SolidEdgeFramework.PropertySets = Nothing

            Try
                If isSheetMetalPart Then

                    objSheetMetalDocument = objDocument
                    propSets = objSheetMetalDocument.Properties
                Else

                    objPartDocument = objDocument
                    propSets = objPartDocument.Properties

                End If
            Catch ex As Exception

            End Try


            If propSets IsNot Nothing Then

                Dim custProps As SolidEdgeFramework.Properties = propSets.Item("Custom")

                For Each prop1 As SolidEdgeFramework.[Property] In custProps

                    Try
                        If prop1.Name = AssemblyBomForm.AssemblyColumns.Document_Number.ToString() Then

                            custPropertiesObj.document_number = prop1.Value

                        ElseIf prop1.Name = AssemblyBomForm.AssemblyColumns.Material_Used.ToString().Replace("_", " ") Then

                            custPropertiesObj.materialUsed = prop1.Value

                        ElseIf prop1.Name = AssemblyBomForm.AssemblyColumns.Material.ToString() Then

                            custPropertiesObj.materialName = prop1.Value

                            'ElseIf prop1.Name = AssemblyBomForm.AssemblyColumns.MATL_SPEC.ToString().Replace("_", " ") Then

                            '    custPropertiesObj.materialSpec = prop1.Value

                            'ElseIf prop1.Name = AssemblyBomForm.AssemblyColumns.Material_Thickness.ToString().Replace("_", " ") Then

                            '    custPropertiesObj.materialThickness = prop1.Value

                        ElseIf prop1.Name = AssemblyBomForm.AssemblyColumns.Quantity.ToString() Then

                            custPropertiesObj.quantity = prop1.Value

                        ElseIf prop1.Name = "Flat_Pattern_Model_CutSizeX" Then 'AssemblyBomForm.AssemblyColumns.Length.ToString() Then

                            custPropertiesObj.Flat_Pattern_Model_CutSizeX = prop1.Value

                        ElseIf prop1.Name = "Flat_Pattern_Model_CutSizeY" Then 'AssemblyBomForm.AssemblyColumns.Width.ToString() Then

                            custPropertiesObj.Flat_Pattern_Model_CutSizeY = prop1.Value

                        ElseIf prop1.Name = "Comments" Then

                            custPropertiesObj.comments = prop1.Value

                        End If
                    Catch ex As Exception
                    End Try

                Next

                Dim custProps2 As Properties = propSets.Item("SummaryInformation")
                For Each prop2 As [Property] In custProps2

                    Try
                        If prop2.Name = "Comments" Then

                            custPropertiesObj.comments = prop2.Value

                        End If
                    Catch ex As Exception
                    End Try

                Next
            End If


        Catch ex As Exception
            MsgBox($"Error in reading custom properties", "Error")
            CustomLogUtil.Log("While Reading Custom Properties", ex.Message, ex.StackTrace)
        End Try

        Return custPropertiesObj
    End Function

    Private Function GetBOMInputData(ByVal materialUsed As String, ByVal colName As String, ByVal dt As DataTable) As String

        Dim colVal As String = String.Empty
        Dim dv As New DataView(dt)
        Dim materialUsedCol As String = ExcelUtil.ExcelSheetColumnsBOM.Material_Used.ToString()
        Dim filter As String = $"Convert([{materialUsedCol}], 'System.String') = '{materialUsed}'"
        dv.RowFilter = filter

        For Each drv As DataRowView In dv
            If Not IsDBNull(drv(colName)) Then
                colVal = drv(colName).ToString()
                Exit For
            End If

        Next

        Return colVal
    End Function

    Private Function GetStockAllowance(ByRef dr As DataRow) As Double
        Dim stockAllowance As Double = 0
        Try
            Dim str As String = dr(AssemblyColumns.Stock_Allowance.ToString())
            str = str.Replace("%", "")
            str = str.Trim()
            stockAllowance = Double.Parse(str)
        Catch ex As Exception
        End Try
        Return stockAllowance
    End Function

    Private Function GetFlatPatternX(ByRef custProp As CustomProperties) As Double
        Dim flatPatternX As Double = 0
        Try
            flatPatternX = Double.Parse(custProp.Flat_Pattern_Model_CutSizeX.ToString().ToUpper().Replace("IN.", "").Replace("IN", "").Trim())
        Catch ex As Exception
        End Try
        Return flatPatternX

    End Function

    Private Function GetFlatPatternY(ByRef custProp As CustomProperties) As Double
        Dim flatPatternY As Double = 0
        Try
            flatPatternY = Double.Parse(custProp.Flat_Pattern_Model_CutSizeY.ToString().ToUpper().Replace("IN.", "").Replace("IN", "").Trim())
        Catch ex As Exception
        End Try
        Return flatPatternY

    End Function

    Private Function GetPercentage(ByVal val As Double, ByVal stockAllowance As Double) As Double

        Dim percentageVal As Double = (stockAllowance * val / 100)
        Return percentageVal

    End Function


    Private Function GetPartQty(ByVal partPath As String) As Integer

        Dim qty As Integer = dicQty(partPath)

        Try
            'dicAssembly_PartQty

            For Each kvp As KeyValuePair(Of String, Dictionary(Of String, Integer)) In dicAssembly_PartQty

                Dim assemblyPath As String = kvp.Key
                Dim dicAssemblyPartCnt As Dictionary(Of String, Integer) = kvp.Value

                If dicAssemblyPartCnt.ContainsKey(partPath) Then

                    Dim partCnt As Integer = dicAssemblyPartCnt(partPath)
                    Dim assemblyCnt As Integer = dicAssemblyQty(assemblyPath)
                    qty += partCnt * assemblyCnt
                End If
            Next

        Catch ex As Exception
        End Try
        'qty = 0

        ''For Each kvp1 As KeyValuePair(Of String, List(Of String)) In dicPartAssemblyList

        ''Next

        'If dicPartAssemblyList.ContainsKey(partPath) Then


        '    Dim partAssemblyList As List(Of String) = dicPartAssemblyList(partPath)

        '    For Each assemblyPath As String In partAssemblyList

        '        Dim assemblyQty As Integer = dicAssemblyQty(assemblyPath)
        '        Dim dicPartQtyInAssembly As Dictionary(Of String, Integer) = dicAssembly_PartQty(assemblyPath)
        '        Dim partQtyInAssembly As Integer = dicPartQtyInAssembly(partPath)

        '        qty = qty + (partQtyInAssembly * assemblyQty)

        '    Next


        'End If



        Return qty
    End Function
    Private Function GetDocumentDetailsDatTable(ByVal dicDocumentDetails As Dictionary(Of String, CustomProperties), ByVal dtAssemblyData As DataTable) As DataTable
        Try
            Dim null = "Not Found"
            Dim NA = "N/A"
            Dim dtBothAvailable As DataTable = dtAssemblyData.Copy()
            dtBothAvailable.TableName = "1"
            dtBothAvailable.Rows.Clear()

            Dim dtSingleAvailable As DataTable = dtAssemblyData.Copy()
            dtSingleAvailable.TableName = "2"
            dtSingleAvailable.Rows.Clear()

            Dim dtBothNotAvailable As DataTable = dtAssemblyData.Copy()
            dtBothNotAvailable.TableName = "3"
            dtBothNotAvailable.Rows.Clear()

            Dim length As String = String.Empty
            Dim Width As String = String.Empty
            'length
            Dim flatPatternX As Double

            'Width
            Dim flatPatternY As Double
            Try

                For Each kvp As KeyValuePair(Of String, CustomProperties) In dicDocumentDetails

                    Dim custProp As CustomProperties = kvp.Value

                    Dim dr As DataRow = dtAssemblyData.NewRow()

                    If custProp.documentPath.ToUpper().EndsWith(".ASM") Then
                        Continue For
                    End If

                    'If custProp.parentDocumentPath = String.Empty Then
                    '    Continue For
                    'End If

                    ' dr(AssemblyColumns.Document_Number.ToString()) = IO.Path.GetFileName(custProp.documentPath) 'MiliComment
                    Try
                        'Material Used
                        If custProp.materialUsed Is Nothing Or custProp.materialUsed = String.Empty Or custProp.materialUsed = "" Then
                            dr(AssemblyColumns.Material_Used.ToString()) = null '"PURCHASED"
                        Else
                            If custProp.materialUsed.ToString.Contains(",") Then
                                Dim values As String() = custProp.materialUsed.ToString.Split(New Char() {","c})
                                dr(AssemblyColumns.Material_Used.ToString()) = custProp.materialUsed
                                custProp.materialUsed = values(0)
                            Else
                                dr(AssemblyColumns.Material_Used.ToString()) = custProp.materialUsed
                            End If

                        End If

                        Try
                            'If custProp.materialUsed.ToString.Contains(",") Then
                            '    Dim values As String() = custProp.materialUsed.ToString.Split(New Char() {","c})
                            '    custProp.materialUsed = values(0)
                            '    Dim v2 As String = values(1)
                            '    If v2.Contains("x") Then
                            '        Dim str2 As String() = v2.Split(New Char() {"x"c})
                            '        length = str2(0)
                            '        If length = "" Then
                            '            flatPatternX = 0
                            '        Else
                            '            flatPatternX = Convert.ToDouble(length)
                            '        End If
                            '        Width = str2(1)
                            '        If Width = "" Then
                            '            flatPatternY = 0
                            '        Else
                            '            flatPatternY = Convert.ToDouble(Width)
                            '        End If
                            '    ElseIf v2.Contains("X") Then
                            '        Dim str2 As String() = v2.Split(New Char() {"X"c})
                            '        length = str2(0)
                            '        If length = "" Or length = " " Or length = String.Empty Then
                            '            flatPatternX = 0
                            '        Else

                            '            If length.Contains("/") Or length.Contains("\") Then
                            '                Dim str1 As String() = length.Split(New Char() {"/"c})
                            '                Dim value1 As Double = Convert.ToDouble(str1(0))
                            '                Dim value2 As Double = Convert.ToDouble(str1(1))
                            '                flatPatternX = value1 / value2
                            '            Else
                            '                flatPatternX = length
                            '            End If
                            '        End If
                            '        Width = str2(1)

                            '        If Width = "" Or Width = " " Or Width = String.Empty Then
                            '            flatPatternY = 0
                            '        Else
                            '            If Width.Contains("/") Or Width.Contains("\") Then
                            '                Dim str1 As String() = Width.Split(New Char() {"/"c})
                            '                Dim value1 As Double = Convert.ToDouble(str1(0))
                            '                Dim value2 As Double = Convert.ToDouble(str1(1))
                            '                flatPatternY = value1 / value2

                            '            Else
                            '                flatPatternY = Width
                            '            End If


                            '        End If
                            '    End If

                            '    '    dr(AssemblyColumns.Length.ToString()) = length
                            '    '    dr(AssemblyColumns.Width.ToString()) = Width

                            '    '    If custProp.materialUsed = "PL" Then
                            '    '        Dim ThicknessValue As String = "T=" + GetBOMInputData(custProp.materialUsed, ExcelUtil.ExcelSheetColumnsBOM.Thickness.ToString().ToString(), dtBOMInputData)
                            '    '        dr(AssemblyColumns.Standard_ThicknessorLength.ToString()) = If(ThicknessValue = "T=", null, ThicknessValue)
                            '    '        dr(AssemblyColumns.Length.ToString()) = custProp.Linear_Length
                            '    '        dr(AssemblyColumns.Width.ToString()) = null
                            '    '    Else
                            '    '        Dim LengthValue As String = "L=" + GetBOMInputData(custProp.materialUsed, ExcelUtil.ExcelSheetColumnsBOM.Linear_Length.ToString().ToString(), dtBOMInputData)
                            '    '        dr(AssemblyColumns.Standard_ThicknessorLength.ToString()) = If(LengthValue = "L=", null, LengthValue)

                            '    '    End If
                            '    'Else
                            '    '    If custProp.materialUsed.ToString().Contains("PL") Then
                            '    '        Dim ThicknessValue As String = "T=" + GetBOMInputData(custProp.materialUsed, ExcelUtil.ExcelSheetColumnsBOM.Thickness.ToString().ToString(), dtBOMInputData)
                            '    '        dr(AssemblyColumns.Standard_ThicknessorLength.ToString()) = If(ThicknessValue = "T=", null, ThicknessValue)
                            '    '        dr(AssemblyColumns.Length.ToString()) = custProp.Linear_Length
                            '    '        dr(AssemblyColumns.Width.ToString()) = null
                            '    '    Else
                            '    '        Dim LengthValue As String = "L=" + GetBOMInputData(custProp.materialUsed, ExcelUtil.ExcelSheetColumnsBOM.Linear_Length.ToString().ToString(), dtBOMInputData)
                            '    '        dr(AssemblyColumns.Standard_ThicknessorLength.ToString()) = If(LengthValue = "L=", null, LengthValue)
                            '    '        Dim l1 = custProp.Flat_Pattern_Model_CutSizeX.ToString().Trim().ToUpper().Replace("IN.", "").Replace("IN", "").Trim()
                            '    '        dr(AssemblyColumns.Length.ToString()) = If(l1 = "", null, l1)
                            '    '        Dim w1 = custProp.Flat_Pattern_Model_CutSizeY.ToString().Trim().ToUpper().Replace("IN.", "").Replace("IN", "").Trim()
                            '    '        dr(AssemblyColumns.Width.ToString()) = If(w1 = "", null, w1)
                            '    '    End If

                            'End If


                            Dim category = GetBOMInputData(custProp.materialUsed, ExcelUtil.ExcelSheetColumnsBOM.Category.ToString(), dtBOMInputData)
                            If category = "Structure" Then
                                Dim LengthValue As String = "L=" + GetBOMInputData(custProp.materialUsed, ExcelUtil.ExcelSheetColumnsBOM.Linear_Length.ToString().ToString(), dtBOMInputData)
                                dr(AssemblyColumns.Standard_ThicknessorLength.ToString()) = If(LengthValue = "L=", null, LengthValue)

                                dr(AssemblyColumns.Length.ToString()) = If(Not custProp.Linear_Length = null, custProp.Linear_Length, null)

                                dr(AssemblyColumns.Width.ToString()) = NA 'null

                                'for Total Length And Width Count
                                If custProp.Linear_Length = "" Or custProp.Linear_Length = " " Or custProp.Linear_Length = String.Empty Then
                                    flatPatternX = 0
                                Else
                                    flatPatternX = length
                                End If
                                flatPatternY = 0

                            ElseIf category = "Plate" Or category = "Sheet" Or category = "SheetMetal" Then
                                Dim ThicknessValue As String = "T=" + GetBOMInputData(custProp.materialUsed, ExcelUtil.ExcelSheetColumnsBOM.Thickness.ToString().ToString(), dtBOMInputData)
                                dr(AssemblyColumns.Standard_ThicknessorLength.ToString()) = If(ThicknessValue = "T=", null, ThicknessValue)

                                Dim l1 = custProp.Flat_Pattern_Model_CutSizeX.ToString().Trim().ToUpper().Replace("IN.", "").Replace("IN", "").Trim()
                                'TEMP29SEP2023
                                If l1.Contains("/") Then
                                    Dim values As String() = l1.ToString.Split(New Char() {" "c})
                                    Dim dotVal As String = values(values.Length() - 1)
                                    l1 = values(0)
                                    values = dotVal.ToString.Split(New Char() {"/"c})
                                    dotVal = Convert.ToDouble(values(0)) / Convert.ToDouble(values(1))
                                    l1 = Convert.ToDouble(l1) + Convert.ToDouble(dotVal)
                                End If
                                dr(AssemblyColumns.Length.ToString()) = If(l1 = "", null, l1)


                                Dim w1 = custProp.Flat_Pattern_Model_CutSizeY.ToString().Trim().ToUpper().Replace("IN.", "").Replace("IN", "").Trim()
                                'TEMP29SEP2023
                                If w1.Contains("/") Then
                                    Dim values As String() = w1.ToString.Split(New Char() {" "c})
                                    Dim dotVal As String = values(values.Length() - 1)
                                    w1 = values(0)
                                    values = dotVal.ToString.Split(New Char() {"/"c})
                                    dotVal = Convert.ToDouble(values(0)) / Convert.ToDouble(values(1))
                                    w1 = Convert.ToDouble(w1) + Convert.ToDouble(dotVal)
                                End If
                                dr(AssemblyColumns.Width.ToString()) = If(w1 = "", null, w1)

                                'for Total Length And Width Count
                                If l1 = "" Or l1 = " " Or l1 = String.Empty Then
                                    flatPatternX = 0
                                Else

                                    flatPatternX = CDbl(Val(l1))
                                End If

                                If w1 = "" Or w1 = " " Or w1 = String.Empty Then
                                    flatPatternY = 0
                                Else

                                    flatPatternY = CDbl(Val(w1))
                                End If
                            End If

                        Catch ex As Exception
                            MsgBox("Material Used : " + custProp.materialUsed + vbNewLine + "Please Check Material Used Property. Formate is Incorrect ")
                        End Try

                        Dim PartType As String = GetBOMInputData(custProp.materialUsed, ExcelUtil.ExcelSheetColumnsBOM.Type.ToString().ToString(), dtBOMInputData)
                        If Not PartType = Nothing Or Not PartType = "" Then
                            dr(AssemblyColumns.PartType.ToString()) = PartType
                        Else
                            dr(AssemblyColumns.PartType.ToString()) = null 'pending
                            PartType = dr(AssemblyColumns.PartType.ToString())
                        End If

                        'BEC Number=Document Number
                        dr(AssemblyColumns.Document_Number.ToString()) = custProp.document_number

                        dr(AssemblyColumns.Material_Spec.ToString()) = If(custProp.materialSpec = "", null, custProp.materialSpec)


                        dr(AssemblyColumns.Material.ToString()) = If(custProp.material = "", null, custProp.material)  'GetBOMInputData(custProp.materialUsed, ExcelUtil.ExcelSheetColumnsBOM.BEC_Material.ToString().ToString(), dtBOMInputData) ' custProp.materialName
                        'dr(AssemblyColumns.Size.ToString()) = GetBOMInputData(custProp.materialUsed, ExcelUtil.ExcelSheetColumnsBOM.Size.ToString(), dtBOMInputData) ' custProp.materialName
                        dr(AssemblyColumns.Description.ToString()) = If(custProp.description = "", null, custProp.description) 'GetBOMInputData(custProp.materialUsed, ExcelUtil.ExcelSheetColumnsBOM.Type.ToString(), dtBOMInputData)
                        dr(AssemblyColumns.Quantity.ToString()) = custProp.quantity 'GetPartQty(custProp.documentPath) '





                        ''StandardThickness/Value
                        'If PartType = Nothing Or PartType = "N/A" Then
                        '    dr(AssemblyColumns.Standard_ThicknessorLength.ToString()) = null
                        'ElseIf PartType = "PL" Then
                        '    Dim ThicknessValue As String = "T=" + GetBOMInputData(custProp.materialUsed, ExcelUtil.ExcelSheetColumnsBOM.Thickness.ToString().ToString(), dtBOMInputData)
                        '    dr(AssemblyColumns.Standard_ThicknessorLength.ToString()) = If(ThicknessValue = "T=", null, ThicknessValue)
                        '    dr(AssemblyColumns.Length.ToString()) = custProp.Linear_Length
                        '    dr(AssemblyColumns.Width.ToString()) = null
                        'Else
                        '    Dim LengthValue As String = "L=" + GetBOMInputData(custProp.materialUsed, ExcelUtil.ExcelSheetColumnsBOM.Linear_Length.ToString().ToString(), dtBOMInputData)
                        '    dr(AssemblyColumns.Standard_ThicknessorLength.ToString()) = If(LengthValue = "L=", null, LengthValue)

                        'End If




                        Dim stockallowance1 = GetBOMInputData(custProp.materialUsed, ExcelUtil.ExcelSheetColumnsBOM.Stock_Clearance.ToString(), dtBOMInputData)
                        dr(AssemblyColumns.Stock_Allowance.ToString()) = If(stockallowance1 = "", "-", stockallowance1)
                        dr(AssemblyColumns.Comments.ToString()) = custProp.comments
                        'extension
                        dr(AssemblyColumns.Extension.ToString()) = IO.Path.GetExtension(custProp.documentPath).ToUpper()
                    Catch ex As Exception

                    End Try


                    'Stock allowance
                    Dim stockAllowance As Double = GetStockAllowance(dr)





                    'Qty
                    Dim qty As Integer = custProp.quantity 'dicQty(custProp.documentPath) 'mili comment

                    'Total Length
                    Dim totalLength As Double = 0

                    'Total Width
                    Dim totalWidth As Double = 0



                    'Total Length and Total Width

                    If flatPatternX > 0 And flatPatternY > 0 Then
                        If flatPatternX > flatPatternY Then
                            Dim valY As Double = qty * flatPatternY
                            Dim perValY As Double = GetPercentage(valY, stockAllowance)
                            dr(AssemblyColumns.Total_Length.ToString()) = valY + perValY
                            totalLength = valY + perValY

                            totalLength = Math.Round(totalLength, 3)

                            Dim valX As Double = qty * flatPatternX
                            Dim perValX As Double = GetPercentage(valX, stockAllowance)
                            dr(AssemblyColumns.Total_Width.ToString()) = valX + perValX
                            totalWidth = valX + perValX

                            totalWidth = Math.Round(totalWidth, 3)
                        Else
                            Dim valY As Double = qty * flatPatternY
                            Dim perValY As Double = GetPercentage(valY, stockAllowance)
                            dr(AssemblyColumns.Total_Width.ToString()) = valY + perValY
                            totalWidth = valY + perValY

                            totalWidth = Math.Round(totalWidth, 3)

                            Dim valX As Double = qty * flatPatternX
                            Dim perValX As Double = GetPercentage(valX, stockAllowance)
                            dr(AssemblyColumns.Total_Length.ToString()) = valX + perValX
                            totalLength = valX + perValX

                            totalLength = Math.Round(totalLength, 3)

                        End If
                    Else
                        If flatPatternY = 0 Then
                            Dim valX As Double = qty * flatPatternX
                            Dim perValX As Double = GetPercentage(valX, stockAllowance)
                            dr(AssemblyColumns.Total_Length.ToString()) = valX + perValX
                            totalLength = valX + perValX

                            totalLength = Math.Round(totalLength, 3)

                            totalWidth = 0
                        End If

                    End If

                    dr(AssemblyColumns.Total_Length.ToString()) = If(totalLength.ToString = "", null, totalLength)
                    If totalWidth > 0 Then
                        dr(AssemblyColumns.Total_Width.ToString()) = totalWidth
                        dr(AssemblyColumns.Area.ToString()) = Math.Round(totalLength * totalWidth, 3)
                    Else
                        dr(AssemblyColumns.OrderArea_Length.ToString()) = If(totalLength.ToString = "", null, totalLength)
                        dr(AssemblyColumns.Total_Width.ToString()) = dr(AssemblyColumns.Width.ToString()) '= ""
                    End If

                    'If dr(AssemblyColumns.Material_Used.ToString()) = "PURCHASED" Then
                    'dr(AssemblyColumns.Description.ToString()) = custProp.description
                    'End If
                    dr(AssemblyColumns.Document_Status.ToString()) = custProp.document_status
                    dr(AssemblyColumns.Hardware.ToString()) = custProp.hardware
                    If flatPatternX > 0 And flatPatternY > 0 Then

                        Dim dr1 As DataRow = dtBothAvailable.NewRow()
                        dr1.ItemArray() = dr.ItemArray()
                        dtBothAvailable.Rows.Add(dr1)

                        'dtBothAvailable.Rows.Add(dr)
                    ElseIf flatPatternX > 0 Or flatPatternY > 0 Then

                        Dim dr2 As DataRow = dtSingleAvailable.NewRow()
                        dr2.ItemArray() = dr.ItemArray()
                        dtSingleAvailable.Rows.Add(dr2)

                        'dtSingleAvailable.Rows.Add(dr)
                    Else

                        Dim dr3 As DataRow = dtBothNotAvailable.NewRow()
                        dr3.ItemArray() = dr.ItemArray()
                        dtBothNotAvailable.Rows.Add(dr3)

                        'dtBothNotAvailable.Rows.Add(dr)
                    End If

                    'dtAssemblyData.Rows.Add(dr)

                Next
            Catch ex As Exception

            End Try

            dtBothAvailable = GetSortingDt(dtBothAvailable)
            dtAssemblyData.Merge(dtBothAvailable)

            dtSingleAvailable = GetSortingDt(dtSingleAvailable)
            dtAssemblyData.Merge(dtSingleAvailable)

            dtBothNotAvailable = GetSortingDt(dtBothNotAvailable)
            dtAssemblyData.Merge(dtBothNotAvailable)



        Catch ex As Exception
            MsgBox($"Error in Get Document Details DatTable ", "Error")
            CustomLogUtil.Log("in Get DocumentDetails DatTable", ex.Message, ex.StackTrace)
        End Try
        Return dtAssemblyData
    End Function

    Private Function GetDocumentDetailsDatTable_old(ByVal dicDocumentDetails As Dictionary(Of String, CustomProperties), ByVal dtAssemblyData As DataTable) As DataTable

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

                dr(AssemblyColumns.Document_Number.ToString()) = IO.Path.GetFileName(custProp.documentPath)
                If custProp.materialUsed Is Nothing Or custProp.materialUsed = String.Empty Then
                    dr(AssemblyColumns.Material_Used.ToString()) = "PURCHASED"
                Else
                    dr(AssemblyColumns.Material_Used.ToString()) = custProp.materialUsed
                End If

                dr(AssemblyColumns.Material.ToString()) = GetBOMInputData(custProp.materialUsed, ExcelUtil.ExcelSheetColumnsBOM.BEC_Material.ToString(), dtBOMInputData) ' custProp.materialName
                dr(AssemblyColumns.Size.ToString()) = GetBOMInputData(custProp.materialUsed, ExcelUtil.ExcelSheetColumnsBOM.Size.ToString(), dtBOMInputData) ' custProp.materialName
                'dr(AssemblyColumns.Description.ToString()) = GetBOMInputData(custProp.materialUsed, ExcelUtil.ExcelSheetColumnsBOM.Type.ToString(), dtBOMInputData)
                dr(AssemblyColumns.Quantity.ToString()) = dicQty(custProp.documentPath)
                dr(AssemblyColumns.Length.ToString()) = custProp.Flat_Pattern_Model_CutSizeX.ToString().Trim().ToUpper().Replace("IN.", "").Replace("IN", "").Trim() ' custProp.Flat_Pattern_Model_CutSizeX
                dr(AssemblyColumns.Width.ToString()) = custProp.Flat_Pattern_Model_CutSizeY.ToString().Trim().ToUpper().Replace("IN.", "").Replace("IN", "").Trim() 'custProp.Flat_Pattern_Model_CutSizeY
                dr(AssemblyColumns.Stock_Allowance.ToString()) = GetBOMInputData(custProp.materialUsed, ExcelUtil.ExcelSheetColumnsBOM.Stock_Clearance.ToString(), dtBOMInputData)
                Dim stockAllowance As Double = GetStockAllowance(dr)

                Dim flatPatternX As Double = 0
                Dim flatPatternY As Double = 0
                Dim qty As Integer = dicQty(custProp.documentPath)
                Try
                    Dim Total_CutSizeX As Double = 0
                    flatPatternX = Double.Parse(custProp.Flat_Pattern_Model_CutSizeX.ToString().ToUpper().Replace("IN.", "").Replace("IN", "").Trim())

                    Total_CutSizeX = qty * flatPatternX
                    dr(AssemblyColumns.Total_Length.ToString()) = Total_CutSizeX.ToString()
                Catch ex As Exception
                End Try

                If dr(AssemblyColumns.Material_Used.ToString()) = " SAHRAN1/4X2X2-A588" Then
                    Debug.Print("aaaa")
                End If

                Try

                    Dim Total_CutSizeY As Double = 0
                    flatPatternY = Double.Parse(custProp.Flat_Pattern_Model_CutSizeY.ToString().ToUpper().Replace("IN.", "").Replace("IN", "").Trim())

                    If flatPatternX > flatPatternY Then
                        dr(AssemblyColumns.Total_Length.ToString()) = (qty * flatPatternY) + ((stockAllowance * (qty * flatPatternY)) / 100) 'stockClearance ' 100*stockClearan/qty*fal
                        dr(AssemblyColumns.Total_Width.ToString()) = (flatPatternX) + ((stockAllowance * (flatPatternX)) / 100)
                    Else
                        dr(AssemblyColumns.Total_Length.ToString()) = (qty * flatPatternX) + ((stockAllowance * (qty * flatPatternX)) / 100)
                        dr(AssemblyColumns.Total_Width.ToString()) = (flatPatternY) + ((stockAllowance * flatPatternY) / 100)
                    End If

                    Dim area As Double = Double.Parse(dr(AssemblyColumns.Total_Length.ToString())) * Double.Parse(dr(AssemblyColumns.Total_Width.ToString()))
                    dr(AssemblyColumns.Area.ToString()) = area.ToString()
                Catch ex As Exception
                    dr(AssemblyColumns.Total_Width.ToString()) = custProp.Flat_Pattern_Model_CutSizeY.ToString().ToUpper()
                    Try
                        dr(AssemblyColumns.Area.ToString()) = (Double.Parse(dr(AssemblyColumns.Total_Length.ToString())) * stockAllowance) / 100
                    Catch ex1 As Exception

                    End Try

                End Try

                If dr(AssemblyColumns.Material_Used.ToString()) = "PURCHASED" Then
                    dr(AssemblyColumns.Description.ToString()) = custProp.description
                Else
                    dr(AssemblyColumns.Description.ToString()) = custProp.description
                End If

                dr(AssemblyColumns.Comments.ToString()) = custProp.comments

                dtAssemblyData.Rows.Add(dr)

            Next
        Catch ex As Exception

        End Try

        Return dtAssemblyData
    End Function

    Private Sub BtnExportExcel_Click(sender As Object, e As EventArgs) Handles btnExportExcel.Click

        ' If Not txtRawMaterialEstimationReportDirPath.Text = "" Then
        btnGetData.Enabled = False
            btnExportExcel.Enabled = False
        Try
            Dim reportDir As String = txtRawMaterialEstimationReportDirPath.Text

            'Try
            '    Dim BetterFolderBrowser As New BetterFolderBrowser()

            '    BetterFolderBrowser.Title = "Select folders"

            '    BetterFolderBrowser.RootFolder = "C:\\"

            '    BetterFolderBrowser.Multiselect = False
            '    If BetterFolderBrowser.ShowDialog() = DialogResult.OK Then
            '        reportDir = BetterFolderBrowser.SelectedFolder
            '    End If
            'Catch ex As Exception
            '    If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            '        Dim path As String = FolderBrowserDialog1.SelectedPath
            '        reportDir = path
            '    End If
            'End Try

            If Not IO.Directory.Exists(reportDir) Then
                MessageBox.Show("Please select valid directory location", "Message")
                btnExportExcel.Enabled = True
                Exit Sub
            End If
            WaitStartSave()
            Dim oExcelUtil As New ExcelUtil()
            '1st Oct 2024
            'Dim excelPath As String = reportDir + "\" + "RawMaterialEstimationReport" + "_" & IO.Path.GetFileNameWithoutExtension(lblAssemblyName.Text) & "_" & System.DateTime.Now.ToString("yyyyMMdd") & "_" & System.DateTime.Now.Hour.ToString & System.DateTime.Now.Minute.ToString & System.DateTime.Now.Second.ToString
            Dim excelPath As String = reportDir + "\" + "RawMaterialEstimationReport" + "_" & IO.Path.GetFileNameWithoutExtension(lblAssemblyName.Text) & "_" & System.DateTime.Now.ToString("MMM_d_yyyy_HH_mm")
            Dim dtData As DataTable = dgvDocumentDetails.DataSource
            Dim dtSheet_Plate_Structure As DataTable = dgvDocumentDetails.DataSource
            Dim dtStdPartsAndHardware As DataTable = dtSheet_Plate_Structure.Copy
            Dim dtMisc As DataTable = dtSheet_Plate_Structure.Copy

            dtSheet_Plate_Structure.TableName = "Sheet_Plate_Structure"
            dtStdPartsAndHardware.TableName = "StdPartsAndHardware"
            dtMisc.TableName = "Misc"

            Try
                Set_All_DT(dtSheet_Plate_Structure, dtStdPartsAndHardware, dtMisc)
            Catch ex As Exception
                MsgBox("Error While Editing Datatable " + ex.Message + ex.StackTrace)
            End Try

            Try
                'dtData.Columns.Remove(AssemblyBomForm.AssemblyColumns.Extension.ToString())
                Dim e1 As New ExcelUtil()
                'e1.SaveExcelReport(excelPath, dtData)
                e1.SaveStdPartsAndMiscExcelReport_RAW(excelPath, dtMisc)
                e1.SaveStdPartsAndMiscExcelReport_RAW(excelPath, dtStdPartsAndHardware)
                e1.SaveSPSExcelReport_RAW(excelPath, dtSheet_Plate_Structure)
                'HighlightNotFound(excelPath)
                CustomLogUtil.Heading("Raw Material Estimation Report successfully created")
                If IO.Directory.Exists(reportDir) Then
                    Process.Start(reportDir)
                End If

            Catch ex As Exception
                MsgBox("Error While Generating Report " + ex.Message + ex.StackTrace)
            End Try

        Catch ex As Exception

        End Try

        btnGetData.Enabled = True
        btnExportExcel.Enabled = True
        WaitEndSave()
        ' End If
    End Sub
    Public Sub HighlightNotFound(ByVal excelPath As String)
        Dim Proceed As Boolean = False
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkBooks As Excel.Workbooks = Nothing
        Dim xlWorkBook As Excel.Workbook = Nothing
        Dim xlWorkSheet As Excel.Worksheet = Nothing
        Dim xlWorkSheets As Excel.Sheets = Nothing
        Dim xlCells As Excel.Range = Nothing
        xlApp = New Excel.Application
        xlApp.DisplayAlerts = False
        xlWorkBooks = xlApp.Workbooks
        xlWorkBook = xlWorkBooks.Open(excelPath)
        xlWorkSheets = xlWorkBook.Sheets
        For x As Integer = 1 To xlWorkSheets.Count
            xlWorkSheet = CType(xlWorkSheets(x), Excel.Worksheet)
            If xlWorkSheet.Name = "Misc" Then

                Proceed = True
                Exit For
            End If
            Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkSheet)
            xlWorkSheet = Nothing
        Next
        If Proceed Then
            xlWorkSheet.Activate()
        End If

        Dim Rcnt = xlWorkSheet.UsedRange.Rows.Count
        Dim Ccnt = xlWorkSheet.UsedRange.Columns.Count
        For i = 2 To Rcnt
            For j = 1 To 6
                If xlWorkSheet.Cells(i)(j).Value = "" Or xlWorkSheet.Cells(i)(j).Value = "N/A" Or xlWorkSheet.Cells(i)(j).Value.ToString.Contains("Not") Then
                    Dim currentCell As Excel.Range = xlWorkSheet.Cells(i)(j) 'change this to your desired cell
                    currentCell.Interior.Color = Color.Red

                    xlWorkBook.Save()

                End If

            Next
        Next
        xlApp.DisplayAlerts = False
        xlWorkBook.Save()
        xlWorkBook.Close()
        xlApp.Quit()
        'MsgBox("Process Completed")
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)

    End Sub
    Public Sub RenameColumnsFromSelectedDt(ByVal dtSheet_Plate_Structure As DataTable, ByVal dtStdPartsAndHardware As DataTable, ByVal dtMisc As DataTable)
        dtSheet_Plate_Structure = RenameDtSPSColumns(dtSheet_Plate_Structure)
        dtStdPartsAndHardware = RnameDtStdpatsColumns(dtStdPartsAndHardware)
        dtMisc = renameDtMiscColumns(dtMisc)
    End Sub
    Public Function RenameDtSPSColumns(ByVal dtSheet_Plate_Structure As DataTable)
        Dim listofcolumns As New List(Of String)
        listofcolumns.Add("Sr")
        listofcolumns.Add("Category")
        listofcolumns.Add("BEC Material")
        listofcolumns.Add("Material Specifications")
        listofcolumns.Add("Standard Thickness/Length (In)")
        listofcolumns.Add("Description")
        listofcolumns.Add("Part Number")
        listofcolumns.Add("Quantity")
        listofcolumns.Add("Length")
        listofcolumns.Add("Width")
        listofcolumns.Add("Stock Allowance")
        listofcolumns.Add("Total Length (In)")
        listofcolumns.Add("Total Width (In)")
        listofcolumns.Add("Area (Sq In)")
        listofcolumns.Add("Order Area/Length (In)")
        listofcolumns.Add("Order Area/Length (Ft)")
        'listofcolumns.Add("Material")
        'listofcolumns.Add("Size")
        'listofcolumns.Add("Extension")
        'listofcolumns.Add("Comments")
        'listofcolumns.Add("Document_Status")
        'listofcolumns.Add("Hardware")
        Dim cnt = dtSheet_Plate_Structure.Columns.Count - 1
        For i = 0 To cnt
            dtSheet_Plate_Structure.Columns(i).ColumnName = listofcolumns.Item(i).ToString()
        Next
        Return dtSheet_Plate_Structure
    End Function
    Public Function RnameDtStdpatsColumns(ByVal dtStdPartsAndHardware As DataTable)
        Dim listofcolumns As New List(Of String)
        listofcolumns.Add("sr")
        listofcolumns.Add("Category")

        'commentlater
        'listofcolumns.Add("Material_Used")
        'listofcolumns.Add("Material_Spec")

        'listofcolumns.Add("Standard_ThicknessorLength")


        listofcolumns.Add("Title")
        listofcolumns.Add("Part Number")
        listofcolumns.Add("Quantity")


        'listofcolumns.Add("Length")
        'listofcolumns.Add("Width")
        'listofcolumns.Add("Stock_Allowance")
        'listofcolumns.Add("Total_Length")
        'listofcolumns.Add("Total_Width")
        'listofcolumns.Add("Area")
        'listofcolumns.Add("OrderArea_Length")
        'listofcolumns.Add("OrderArea_LengthFT")
        'listofcolumns.Add("Material")
        'listofcolumns.Add("Size")
        'listofcolumns.Add("Extension")

        listofcolumns.Add("Information")

        'commentlater
        'listofcolumns.Add("Document_Status")
        'listofcolumns.Add("Hardware")
        Dim cnt = dtStdPartsAndHardware.Columns.Count - 1
        For i = 0 To cnt
            Try
                dtStdPartsAndHardware.Columns(i).ColumnName = listofcolumns.Item(i).ToString()
            Catch ex As Exception

            End Try

        Next
        Return dtStdPartsAndHardware
    End Function
    Public Function renameDtMiscColumns(ByVal dtMisc As DataTable)
        Dim listofcolumns As New List(Of String)
        listofcolumns.Add("Sr")
        listofcolumns.Add("Category")


        listofcolumns.Add("BEC Material Code")
        listofcolumns.Add("Material Specification")
        'listofcolumns.Add("Standard_ThicknessorLength")


        listofcolumns.Add("Title")
        listofcolumns.Add("Part Number")
        listofcolumns.Add("Quantity")


        'listofcolumns.Add("Length")
        'listofcolumns.Add("Width")
        'listofcolumns.Add("Stock_Allowance")
        'listofcolumns.Add("Total_Length")
        'listofcolumns.Add("Total_Width")
        'listofcolumns.Add("Area")
        'listofcolumns.Add("OrderArea_Length")
        'listofcolumns.Add("OrderArea_LengthFT")
        'listofcolumns.Add("Material")
        'listofcolumns.Add("Size")
        'listofcolumns.Add("Extension")

        listofcolumns.Add("Information")

        'commentlater
        'listofcolumns.Add("Document_Status")
        'listofcolumns.Add("Hardware")
        Dim cnt = dtMisc.Columns.Count - 1
        For i = 0 To cnt

            dtMisc.Columns(i).ColumnName = listofcolumns.Item(i).ToString()
        Next
        Return dtMisc
    End Function
    Public Function SetSrNoForAllDt(ByVal dt As DataTable) As DataTable
        Dim nf = "Not Found"
        Dim cnt = dt.Rows.Count - 1
        For i = 0 To cnt
            dt.Rows(i)(0) = i + 1
            If (dt.TableName.Contains("Misc")) Then
                dt.Rows(i)(1) = "Misc"
            End If
            If dt.TableName.Contains("Sheet_Plate_Structure") Then
                Dim Value As String
                For j = 0 To dt.Columns.Count - 1
                    Select Case (dt.Columns(j).ColumnName)

                        Case "Length"
                            Value = dt.Rows(i)(j).ToString
                            If Value = "" Or Value = " " Then
                                dt.Rows(i)(j) = nf
                            End If
                        Case "Total Length (In)"
                            Value = dt.Rows(i)(j).ToString
                            If Value = "" Or Value = "0" Then
                                dt.Rows(i)(j) = nf
                            End If
                        Case "Area (Sq In)"
                            Value = dt.Rows(i)(j).ToString
                            If Value = "" Or Value = "0" Or Value = " " Then
                                dt.Rows(i)(j) = nf
                            End If
                        Case "Order Area/Length (In)"
                            Value = dt.Rows(i)(j).ToString
                            If Value = "" Or Value = "0" Or Value = " " Then
                                dt.Rows(i)(j) = nf
                            End If
                        Case "Order Area/Length (Ft)"
                            Value = dt.Rows(i)(j).ToString
                            If Value = "" Or Value = "0" Or Value = " " Then
                                dt.Rows(i)(j) = nf
                            End If
                    End Select
                Next
            End If
        Next
        Return dt

    End Function

    Public Sub Set_All_DT(ByVal dtSheet_Plate_Structure As DataTable, ByVal dtStdPartsAndHardware As DataTable, ByVal dtMisc As DataTable)
        Dim lstOf_StdPart_SR As New List(Of String)
        Dim lstOf_SPS_SR As New List(Of String)
        Dim lstOf_Misc_SR As New List(Of String)
        EditListsOfAllDt(dtSheet_Plate_Structure, lstOf_SPS_SR, dtStdPartsAndHardware, lstOf_StdPart_SR, dtMisc, lstOf_Misc_SR)
        RemoveRowsFromSelectedDt(dtSheet_Plate_Structure, lstOf_SPS_SR, dtStdPartsAndHardware, lstOf_StdPart_SR, dtMisc, lstOf_Misc_SR)
        RemoveColumnsFromSelectedDt(dtSheet_Plate_Structure, dtStdPartsAndHardware, dtMisc)
        RenameColumnsFromSelectedDt(dtSheet_Plate_Structure, dtStdPartsAndHardware, dtMisc)
        dtSheet_Plate_Structure = OrderCategoryForSPSDt(dtSheet_Plate_Structure)
        dtStdPartsAndHardware = OrderCategoryForStdPartsDt(dtStdPartsAndHardware)
        'dtMisc = OrderCategoryForMiscDt(dtMisc)

        SetSrNo(dtStdPartsAndHardware, dtMisc, dtSheet_Plate_Structure)

        ' dtSheet_Plate_Structure = SetTotalLengthAndWidthInDtSps(dtSheet_Plate_Structure)
        'dtStdPartsAndHardware = EditnullvaluesInAllDt(dtStdPartsAndHardware)
        'dtSheet_Plate_Structure = EditnullvaluesInAllDt(dtSheet_Plate_Structure)


    End Sub
    Public Function SetTotalLengthAndWidthInDtSps(ByVal dtSheet_Plate_Structure As DataTable)
        Dim nf = "Not Found"
        Dim na = "N/A"
        Try
            For i = 0 To dtSheet_Plate_Structure.Rows.Count - 1
                Dim val1 As String = dtSheet_Plate_Structure.Rows(i)(9).ToString
                Dim val2 As String = dtSheet_Plate_Structure.Rows(i)(10).ToString
                If val1 = nf And val2 = nf Then
                    dtSheet_Plate_Structure.Rows(i)(11) = nf
                    dtSheet_Plate_Structure.Rows(i)(12) = nf
                    dtSheet_Plate_Structure.Rows(i)(13) = nf
                    dtSheet_Plate_Structure.Rows(i)(14) = nf
                    dtSheet_Plate_Structure.Rows(i)(15) = nf
                ElseIf val1 = "" And val2 = na Then
                    dtSheet_Plate_Structure.Rows(i)(11) = na
                    dtSheet_Plate_Structure.Rows(i)(12) = na
                    dtSheet_Plate_Structure.Rows(i)(13) = na
                    dtSheet_Plate_Structure.Rows(i)(14) = na
                    dtSheet_Plate_Structure.Rows(i)(15) = na

                End If
                For j = 11 To 15
                    Dim val3 As String = dtSheet_Plate_Structure.Rows(i)(j).ToString
                    Try
                        If val3 = "" Or val3 = 0 Then
                            dtSheet_Plate_Structure.Rows(i)(j) = nf
                        End If
                    Catch ex As Exception
                        dtSheet_Plate_Structure.Rows(i)(j) = nf
                    End Try

                Next
            Next

        Catch ex As Exception

        End Try

        'If dtSheet_Plate_Structure.Rows(i)(10) = na Then
        '    dtSheet_Plate_Structure.Rows(i)(13) = na


        'End If
        'If dtSheet_Plate_Structure.Rows(i)(13) = "" And dtSheet_Plate_Structure.Rows(i)(10) = na Or dtSheet_Plate_Structure.Rows(i)(13) = 0 And dtSheet_Plate_Structure.Rows(i)(10) = na Then
        '    dtSheet_Plate_Structure.Rows(i)(13) = na
        '    dtSheet_Plate_Structure.Rows(i)(14) = na
        '    dtSheet_Plate_Structure.Rows(i)(15) = na
        'ElseIf dtSheet_Plate_Structure.Rows(i)(13) = "" And dtSheet_Plate_Structure.Rows(i)(10) = nf Or dtSheet_Plate_Structure.Rows(i)(13) = 0 And dtSheet_Plate_Structure.Rows(i)(10) = nf Then
        '    dtSheet_Plate_Structure.Rows(i)(13) = nf
        '    dtSheet_Plate_Structure.Rows(i)(14) = nf
        '    dtSheet_Plate_Structure.Rows(i)(15) = nf
        'End If

        Return dtSheet_Plate_Structure
    End Function
    Public Function OrderCategoryForMiscDt(ByVal dtMisc As DataTable)
        Dim Rcount = dtMisc.Rows.Count - 1
        Dim notdashcount = 0
        Dim na = "N/A"
        Try
            For i = 0 To Rcount - 1
                If Not dtMisc.Rows(i)(1).ToString.Contains("Misc") Then
                    Dim dr As DataRow = dtMisc.NewRow()
                    For j = 0 To dtMisc.Columns.Count - 1
                        Dim value As String = dtMisc.Rows(i)(j).ToString
                        dr(j) = value

                    Next
                    dtMisc.Rows(i).Delete()
                    dtMisc.Rows.Add(dr)
                Else
                    'For j = 0 To dtSheet_Plate_Structure.Columns.Count - 1
                    '    Select Case (dtSheet_Plate_Structure.Columns(j).ColumnName)

                    '        Case "Width"
                    '            If Not dtSheet_Plate_Structure.Rows(i)(1).ToString.Contains("-") And dtSheet_Plate_Structure.Rows(i)(j) = "" Or dtSheet_Plate_Structure.Rows(i)(j) = "Not Found" Then
                    '                dtSheet_Plate_Structure.Rows(i)(j) = na
                    '            ElseIf dtSheet_Plate_Structure.Rows(i)(1).ToString.Contains("-") And dtSheet_Plate_Structure.Rows(i)(j) = na Then
                    '                dtSheet_Plate_Structure.Rows(i)(j) = "Not Found"

                    '            End If
                    '    End Select
                    'Next
                    notdashcount += 1
                End If
                If notdashcount + i + 1 = Rcount Then
                    Exit For
                End If
            Next
        Catch ex As Exception

        End Try

        Return dtMisc
    End Function
    Public Function OrderCategoryForStdPartsDt(ByVal dtStdPartsAndHardware As DataTable)
        Dim Rcount = dtStdPartsAndHardware.Rows.Count - 1
        Dim notdashcount = 0
        Dim na = "N/A"
        Try
            For i = Rcount To 0 Step -1
                If Not dtStdPartsAndHardware.Rows(i)(1).ToString.Contains("Hardware") Then
                    Dim dr As DataRow = dtStdPartsAndHardware.NewRow()
                    For j = 0 To dtStdPartsAndHardware.Columns.Count - 1
                        Dim value As String = dtStdPartsAndHardware.Rows(i)(j).ToString
                        dr(j) = value

                    Next
                    dtStdPartsAndHardware.Rows(i).Delete()
                    dtStdPartsAndHardware.Rows.Add(dr)
                Else
                    'For j = 0 To dtSheet_Plate_Structure.Columns.Count - 1
                    '    Select Case (dtSheet_Plate_Structure.Columns(j).ColumnName)

                    '        Case "Width"
                    '            If Not dtSheet_Plate_Structure.Rows(i)(1).ToString.Contains("-") And dtSheet_Plate_Structure.Rows(i)(j) = "" Or dtSheet_Plate_Structure.Rows(i)(j) = "Not Found" Then
                    '                dtSheet_Plate_Structure.Rows(i)(j) = na
                    '            ElseIf dtSheet_Plate_Structure.Rows(i)(1).ToString.Contains("-") And dtSheet_Plate_Structure.Rows(i)(j) = na Then
                    '                dtSheet_Plate_Structure.Rows(i)(j) = "Not Found"

                    '            End If
                    '    End Select
                    'Next

                End If
                If i = 0 Then
                    Exit For
                End If
            Next
        Catch ex As Exception

        End Try

        Return dtStdPartsAndHardware
    End Function
    Public Function OrderCategoryForSPSDt(ByVal dtSheet_Plate_Structure As DataTable)
        Dim Rcount = dtSheet_Plate_Structure.Rows.Count - 1
        Dim notdashcount = 0
        Dim na = "N/A"
        Try
            For i = 0 To Rcount - 1
                If Not dtSheet_Plate_Structure.Rows(i)(1).ToString.Contains("-") Then
                    Dim dr As DataRow = dtSheet_Plate_Structure.NewRow()
                    For j = 0 To dtSheet_Plate_Structure.Columns.Count - 1
                        Dim value As String = dtSheet_Plate_Structure.Rows(i)(j).ToString
                        dr(j) = value

                    Next
                    dtSheet_Plate_Structure.Rows(i).Delete()
                    dtSheet_Plate_Structure.Rows.Add(dr)
                Else
                    'For j = 0 To dtSheet_Plate_Structure.Columns.Count - 1
                    '    Select Case (dtSheet_Plate_Structure.Columns(j).ColumnName)

                    '        Case "Width"
                    '            If Not dtSheet_Plate_Structure.Rows(i)(1).ToString.Contains("-") And dtSheet_Plate_Structure.Rows(i)(j) = "" Or dtSheet_Plate_Structure.Rows(i)(j) = "Not Found" Then
                    '                dtSheet_Plate_Structure.Rows(i)(j) = na
                    '            ElseIf dtSheet_Plate_Structure.Rows(i)(1).ToString.Contains("-") And dtSheet_Plate_Structure.Rows(i)(j) = na Then
                    '                dtSheet_Plate_Structure.Rows(i)(j) = "Not Found"

                    '            End If
                    '    End Select
                    'Next
                    notdashcount += 1
                End If
                If notdashcount + i + 1 = Rcount Then
                    Exit For
                End If
            Next
        Catch ex As Exception

        End Try

        Return dtSheet_Plate_Structure
    End Function
    Public Sub SetSrNo(ByVal dtStdPartsAndHardware As DataTable, ByVal dtMisc As DataTable, ByVal dtSheet_Plate_Structure As DataTable)
        dtStdPartsAndHardware = SetSrNoForAllDt(dtStdPartsAndHardware)
        dtMisc = SetSrNoForAllDt(dtMisc)
        dtSheet_Plate_Structure = SetSrNoForAllDt(dtSheet_Plate_Structure)
    End Sub
    Public Sub RemoveColumnsFromSelectedDt(ByVal dtSheet_Plate_Structure As DataTable, ByVal dtStdPartsAndHardware As DataTable, ByVal dtMisc As DataTable)
        dtSheet_Plate_Structure = RemoveDtSPSColumns(dtSheet_Plate_Structure)
        dtStdPartsAndHardware = RemoveDtStdpatsColumns(dtStdPartsAndHardware)
        dtMisc = removeRowsFromDtMisc(dtMisc)
    End Sub

    Public Function RemoveDtSPSColumns(ByVal dtSheet_Plate_Structure As DataTable)
        Dim listofcolumns As New List(Of String)
        'listofcolumns.Add("Sr")
        'listofcolumns.Add("PartType")
        'listofcolumns.Add("Material_Used")
        'listofcolumns.Add("Material_Spec")
        'listofcolumns.Add("Standard_ThicknessorLength")
        'listofcolumns.Add("Description")
        'listofcolumns.Add("Document_Number")
        'listofcolumns.Add("Quantity")
        'listofcolumns.Add("Length")
        'listofcolumns.Add("Width")
        'listofcolumns.Add("Stock_Allowance")
        'listofcolumns.Add("Total_Length")
        'listofcolumns.Add("Total_Width")
        'listofcolumns.Add("Area")
        'listofcolumns.Add("OrderArea_Length")
        'listofcolumns.Add("OrderArea_LengthFT")
        listofcolumns.Add("Material")
        listofcolumns.Add("Size")
        listofcolumns.Add("Extension")
        listofcolumns.Add("Comments")
        listofcolumns.Add("Document_Status")
        listofcolumns.Add("Hardware")
        Dim cnt = dtSheet_Plate_Structure.Columns.Count - 1
        For i = 0 To cnt
            If i > cnt Then
                cnt = dtSheet_Plate_Structure.Columns.Count - 1
                i = 0
            End If
            If listofcolumns.Contains(dtSheet_Plate_Structure.Columns(i).ColumnName) Then
                Dim Value = dtSheet_Plate_Structure.Columns(i).ColumnName
                dtSheet_Plate_Structure.Columns.Remove(Value)
                listofcolumns.Remove(Value)
                cnt = cnt - 1
            End If
            If listofcolumns.Count = 0 Then
                Exit For
            End If

        Next
        Return dtSheet_Plate_Structure
    End Function

    Public Function RemoveDtStdpatsColumns(ByVal dtStdPartsAndHardware As DataTable)
        Dim listofcolumns As New List(Of String)
        'listofcolumns.Add("Sr")
        'listofcolumns.Add("PartType")


        'uncommentlater
        listofcolumns.Add("Material_Used")
        listofcolumns.Add("Material_Spec")


        listofcolumns.Add("Standard_ThicknessorLength")
        'listofcolumns.Add("Description")
        'listofcolumns.Add("Document_Number")
        'listofcolumns.Add("Quantity")
        listofcolumns.Add("Length")
        listofcolumns.Add("Width")
        listofcolumns.Add("Stock_Allowance")
        listofcolumns.Add("Total_Length")
        listofcolumns.Add("Total_Width")
        listofcolumns.Add("Area")
        listofcolumns.Add("OrderArea_Length")
        listofcolumns.Add("OrderArea_LengthFT")
        listofcolumns.Add("Material")
        listofcolumns.Add("Size")
        listofcolumns.Add("Extension")
        'listofcolumns.Add("Comments")

        'uncommentlater
        listofcolumns.Add("Document_Status")
        listofcolumns.Add("Hardware")
        Dim cnt = dtStdPartsAndHardware.Columns.Count - 1
        For i = 0 To cnt
            If i > cnt Then
                cnt = dtStdPartsAndHardware.Columns.Count - 1
                i = 0
            End If
            If listofcolumns.Contains(dtStdPartsAndHardware.Columns(i).ColumnName) Then
                Dim Value = dtStdPartsAndHardware.Columns(i).ColumnName
                dtStdPartsAndHardware.Columns.Remove(Value)
                listofcolumns.Remove(Value)
                cnt = cnt - 1
            End If
            If listofcolumns.Count = 0 Then
                Exit For
            End If

        Next
        Return dtStdPartsAndHardware
    End Function
    Public Function removeRowsFromDtMisc(ByVal dtMisc As DataTable)
        Dim listofcolumns As New List(Of String)
        'listofcolumns.Add("Sr")
        'listofcolumns.Add("PartType")
        'listofcolumns.Add("Material_Used")
        'listofcolumns.Add("Material_Spec")
        listofcolumns.Add("Standard_ThicknessorLength")
        'listofcolumns.Add("Description")
        'listofcolumns.Add("Document_Number")
        'listofcolumns.Add("Quantity")
        listofcolumns.Add("Length")
        listofcolumns.Add("Width")
        listofcolumns.Add("Stock_Allowance")
        listofcolumns.Add("Total_Length")
        listofcolumns.Add("Total_Width")
        listofcolumns.Add("Area")
        listofcolumns.Add("OrderArea_Length")
        listofcolumns.Add("OrderArea_LengthFT")
        listofcolumns.Add("Material")
        listofcolumns.Add("Size")
        listofcolumns.Add("Extension")
        'listofcolumns.Add("Comments")

        'uncommentlater
        listofcolumns.Add("Document_Status")
        listofcolumns.Add("Hardware")
        Dim cnt = dtMisc.Columns.Count - 1
        For i = 0 To cnt
            If i > cnt Then
                cnt = dtMisc.Columns.Count - 1
                i = 0
            End If
            If listofcolumns.Contains(dtMisc.Columns(i).ColumnName) Then
                Dim Value = dtMisc.Columns(i).ColumnName
                dtMisc.Columns.Remove(Value)
                listofcolumns.Remove(Value)
                cnt = cnt - 1
            End If
            If listofcolumns.Count = 0 Then
                Exit For
            End If

        Next
        Return dtMisc
    End Function
    Public Sub EditListsOfAllDt(ByVal dtSheet_Plate_Structure As DataTable, ByVal lstOf_SPS_SR As List(Of String), ByVal dtStdPartsAndHardware As DataTable, ByVal lstOf_StdPart_SR As List(Of String), ByVal dtMisc As DataTable, ByVal lstOf_Misc_SR As List(Of String))
        Try
            Dim dt As DataTable = dtSheet_Plate_Structure.Copy
            Dim dthardware As DataTable = dtSheet_Plate_Structure.Copy
            Dim nf = "Not Found"
            For i = 0 To dt.Rows.Count - 1
                Dim MaterialUsed As String = String.Empty
                Dim MaterialSpec As String = String.Empty
                Dim DocumentStatus As String = String.Empty
                Dim Hardware As String = String.Empty
                Dim Category As String = String.Empty
                For j = 0 To dt.Columns.Count - 1

                    Select Case dt.Columns(j).ColumnName
                        Case "Material_Used"
                            MaterialUsed = dt.Rows(i)(j)
                        Case "Document_Status"
                            DocumentStatus = dt.Rows(i)(j)
                        Case "Hardware"
                            Hardware = dt.Rows(i)(j)
                        Case "Material_Spec"
                            MaterialSpec = dt.Rows(i)(j)
                        Case "PartType"
                            If dt.Rows(i)(j).ToString.Count < 3 And dt.Rows(i)(j).ToString.Contains(nf) Then
                                Category = dt.Rows(i)(j)
                            End If

                    End Select

                Next
                If MaterialUsed = "PURCHASED" Or MaterialSpec = "PURCHASED" Then
                    MaterialUsed = "PURCHASED"
                End If
                If MaterialUsed = "PURCHASED" And DocumentStatus = "Baselined" And Hardware = "True" Then ' Or MaterialSpec = "PURCHASED" And DocumentStatus = "Baselined" And Hardware = "True" Then
                    'dtSheet_Plate_Structure.Rows.RemoveAt(i)
                    lstOf_SPS_SR.Add(dtSheet_Plate_Structure.Rows(i)(0))

                    lstOf_Misc_SR.Add(dtMisc.Rows(i)(0)) 'new


                    dtStdPartsAndHardware.Rows(i)(1) = "Hardware"
                ElseIf MaterialUsed = "PURCHASED" And DocumentStatus = "Baselined" And Hardware = "False" Then  'Or MaterialSpec = "PURCHASED" And DocumentStatus = "Baselined" And Hardware = "False" Then
                    'dtSheet_Plate_Structure.Rows.RemoveAt(i)
                    lstOf_SPS_SR.Add(dtSheet_Plate_Structure.Rows(i)(0))


                    lstOf_Misc_SR.Add(dtMisc.Rows(i)(0)) 'new

                    dtStdPartsAndHardware.Rows(i)(1) = "Standard Part"
                ElseIf MaterialUsed = "PURCHASED" Or DocumentStatus = "Baselined" Then 'Or MaterialSpec = "PURCHASED" Then
                    'dtSheet_Plate_Structure.Rows.RemoveAt(i)
                    lstOf_SPS_SR.Add(dtSheet_Plate_Structure.Rows(i)(0))
                    lstOf_StdPart_SR.Add(dtStdPartsAndHardware.Rows(i)(0))
                    'dtStdPartsAndHardware.Rows.RemoveAt(i)
                    dtMisc.Rows(i)(1) = "Misc"
                Else
                    lstOf_StdPart_SR.Add(dtStdPartsAndHardware.Rows(i)(0))
                    'dtStdPartsAndHardware.Rows.RemoveAt(i)

                    lstOf_Misc_SR.Add(dtMisc.Rows(i)(0)) 'new


                    'dtMisc.Rows.RemoveAt(i)
                End If




            Next


        Catch ex As Exception

        End Try
    End Sub


    Public Sub RemoveRowsFromSelectedDt(ByVal dtSheet_Plate_Structure As DataTable, ByVal lstOf_SPS_SR As List(Of String), ByVal dtStdPartsAndHardware As DataTable, ByVal lstOf_StdPart_SR As List(Of String), ByVal dtMisc As DataTable, ByVal lstOf_Misc_SR As List(Of String))
        Try

            dtSheet_Plate_Structure = RemoveRowsFromDtSPS(dtSheet_Plate_Structure, lstOf_SPS_SR)
            dtStdPartsAndHardware = removeRowsFromDtStdParts(dtStdPartsAndHardware, lstOf_StdPart_SR)
            lstOf_Misc_SR = getlistofDocumentsForDelelteDtMscData(dtSheet_Plate_Structure, dtStdPartsAndHardware, lstOf_Misc_SR)
            dtMisc = removeRowsFromDtMisc(dtMisc, lstOf_Misc_SR)
        Catch ex As Exception

        End Try
    End Sub
    Public Function getlistofDocumentsForDelelteDtMscData(ByVal dtSheet_Plate_Structure As DataTable, ByVal dtStdPartsAndHardware As DataTable, ByVal lstOf_Misc_SR As List(Of String))

        lstOf_Misc_SR.Clear()

        For i = 0 To dtSheet_Plate_Structure.Rows.Count - 1
            lstOf_Misc_SR.Add(dtSheet_Plate_Structure.Rows(i)(0))
        Next

        For i = 0 To dtStdPartsAndHardware.Rows.Count - 1
            lstOf_Misc_SR.Add(dtStdPartsAndHardware.Rows(i)(0))
        Next
        Return lstOf_Misc_SR
    End Function

    Public Function removeRowsFromDtMisc(ByVal dtMisc As DataTable, ByVal lstOf_Misc_SR As List(Of String)) As DataTable


        Dim cnt2 = dtMisc.Rows.Count - 1
        For i = 0 To cnt2
            If i > cnt2 Then
                cnt2 = dtMisc.Rows.Count - 1
                i = 0
                'If cnt2 + 1 = lstOf_StdPart_SR.Count Then
                '    Exit For
                'End If
            End If

            Dim lstcnt = lstOf_Misc_SR.Count - 1
            If lstcnt = -1 Then
                Exit For
            End If
            For j = 0 To lstcnt
                Dim Sr = lstOf_Misc_SR.Item(j)
                If dtMisc.Rows(i)(0) = Sr Then

                    dtMisc.Rows.RemoveAt(i)
                    cnt2 = cnt2 - 1
                    lstOf_Misc_SR.Remove(Sr)


                    lstcnt = lstcnt - 1
                    Exit For
                End If
            Next
        Next

        Return dtMisc
    End Function
    Public Function removeRowsFromDtStdParts(ByVal dtStdPartsAndHardware As DataTable, ByVal lstOf_StdPart_SR As List(Of String)) As DataTable


        Dim cnt2 = dtStdPartsAndHardware.Rows.Count - 1
        For i = 0 To cnt2
            If i > cnt2 Then
                cnt2 = dtStdPartsAndHardware.Rows.Count - 1
                i = 0
                'If cnt2 + 1 = lstOf_StdPart_SR.Count Then
                '    Exit For
                'End If
            End If

            Dim lstcnt = lstOf_StdPart_SR.Count - 1
            If lstcnt = -1 Then
                Exit For
            End If
            For j = 0 To lstcnt
                Dim Sr = lstOf_StdPart_SR.Item(j)
                If dtStdPartsAndHardware.Rows(i)(0) = Sr Then

                    dtStdPartsAndHardware.Rows.RemoveAt(i)
                    cnt2 = cnt2 - 1
                    lstOf_StdPart_SR.Remove(Sr)
                    lstcnt = lstcnt - 1
                    Exit For
                End If
            Next
        Next

        Return dtStdPartsAndHardware
    End Function

    Public Function RemoveRowsFromDtSPS(ByVal dtSheet_Plate_Structure As DataTable, ByVal lstOf_SPS_SR As List(Of String)) As DataTable
        Dim listofnotfound As New List(Of String)
        '---------------------------------------------------
        Dim cnt2 = dtSheet_Plate_Structure.Rows.Count - 1

        cnt2 = dtSheet_Plate_Structure.Rows.Count - 1
        For i = 0 To cnt2
            If i > cnt2 Then
                cnt2 = dtSheet_Plate_Structure.Rows.Count - 1
                i = 0
                'If cnt2 + 1 = lstOf_StdPart_SR.Count Then
                '    Exit For
                'End If
            End If

            Dim lstcnt = lstOf_SPS_SR.Count - 1
            If lstcnt = -1 Then
                Exit For
            End If
            For j = 0 To lstcnt
                Dim Sr = lstOf_SPS_SR.Item(j)

                If dtSheet_Plate_Structure.Rows(i)(0) = Sr Then

                    dtSheet_Plate_Structure.Rows.RemoveAt(i)
                    cnt2 = cnt2 - 1
                    lstOf_SPS_SR.Remove(Sr)
                    lstcnt = lstcnt - 1
                    Exit For
                End If
            Next
        Next
        'remove not found
        '------------------------------------------------------
        cnt2 = dtSheet_Plate_Structure.Rows.Count - 1

        'Add ListofNotFound
        For k = 0 To cnt2




            If dtSheet_Plate_Structure.Rows(k)(1) = "Not Found" Then

                If Not listofnotfound.Contains(dtSheet_Plate_Structure.Rows(k)(0)) Then
                    listofnotfound.Add(dtSheet_Plate_Structure.Rows(k)(0))
                End If

            End If

        Next
        'lstOf_SPS_SR.Clear()
        'lstOf_SPS_SR = listofnotfound

        cnt2 = dtSheet_Plate_Structure.Rows.Count - 1
        For i = 0 To cnt2
            If i > cnt2 Then
                cnt2 = dtSheet_Plate_Structure.Rows.Count - 1
                i = 0
                'If cnt2 + 1 = lstOf_StdPart_SR.Count Then
                '    Exit For
                'End If
            End If

            Dim lstcnt = listofnotfound.Count - 1
            If lstcnt = -1 Then
                Exit For
            End If
            For j = 0 To lstcnt
                Dim Sr = listofnotfound.Item(j)

                If dtSheet_Plate_Structure.Rows(i)(0) = Sr Then

                    dtSheet_Plate_Structure.Rows.RemoveAt(i)
                    cnt2 = cnt2 - 1
                    listofnotfound.Remove(Sr)
                    lstcnt = lstcnt - 1
                    Exit For
                End If
            Next
        Next



        '----------------------------------------------------
        Return dtSheet_Plate_Structure
    End Function

    Private Sub BtnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    '-------------------------------------------------------##############------------------------------------------------------
    '-------------------------------------------------------##############-------------------------------------------------------
    Public Function GetDocumentDetails(ByVal dt As DataTable) As Dictionary(Of String, CustomProperties)
        Dim custPropertiesObj As New CustomProperties()
        Dim dicDocumentDetails As New Dictionary(Of String, CustomProperties)()
        For Each dr As DataRow In dt.Rows
            Try

                Dim ItemNumber As String = dr("Item Number").ToString()
                If ItemNumber = String.Empty Then
                    Return dicDocumentDetails
                End If
                custPropertiesObj.document_number = dr("Document Number")
                custPropertiesObj.Linear_Length = dr("Linear_Length")
                custPropertiesObj.Flat_Pattern_Model_CutSizeX = dr("Flat_Pattern_Model_CutSizeX")
                If dr("Quantity") IsNot Nothing And Not dr("Quantity") = String.Empty Then
                    custPropertiesObj.quantity = dr("Quantity")
                End If

                custPropertiesObj.Flat_Pattern_Model_CutSizeY = dr("Flat_Pattern_Model_CutSizeY")
                custPropertiesObj.description = dr("Title") 'dr("Description")
                custPropertiesObj.comments = dr("Comments")
                custPropertiesObj.materialUsed = dr("Material Used")
                custPropertiesObj.materialSpec = dr("Material Specification")
                custPropertiesObj.partType = dr("Category")
                custPropertiesObj.hardware = dr("Hardware")
                custPropertiesObj.document_status = dr("Status Text")
                custPropertiesObj.material = dr("Material")
                Dim docFullPath As String = dr("File Name (full path)")
                If Not dicDocumentDetails.ContainsKey(docFullPath) Then
                    dicDocumentDetails.Add(docFullPath, custPropertiesObj)
                End If

            Catch ex As Exception

            End Try
            custPropertiesObj = New CustomProperties()
        Next

        Return dicDocumentDetails
    End Function
    Public Function Getpartlist() As DataTable
        Dim objApplication As SolidEdgeFramework.Application = Nothing
        Dim objDocuments As SolidEdgeFramework.Documents = Nothing
        Dim objDraft As SolidEdgeDraft.DraftDocument = Nothing
        Dim objSheet As SolidEdgeDraft.Sheet = Nothing
        Dim objModelLinks As SolidEdgeDraft.ModelLinks = Nothing
        Dim objModelLink As SolidEdgeDraft.ModelLink = Nothing
        Dim objDrawingViews As SolidEdgeDraft.DrawingViews = Nothing
        Dim objDrawingView As SolidEdgeDraft.DrawingView = Nothing

        Try
            '  OleMessageFilter.Register()

            ' Connect to a running instance of Solid Edge
            objApplication = Marshal.GetActiveObject("SolidEdge.Application")

            ' Get a reference to the documents collection
            objDocuments = objApplication.Documents

            ' Add a Draft document      
            objDraft = objDocuments.Add("SolidEdge.DraftDocument")

            ' Get a reference to the active sheet
            objSheet = objDraft.ActiveSheet

            ' Get a reference to the model links collection
            objModelLinks = objDraft.ModelLinks
            Dim filename As String
            Dim file As String = objAssemblyDocument.FullName
            filename = System.IO.Path.Combine(SolidEdgeCommunity.SolidEdgeUtils.GetTrainingFolderPath(), file)

            ' Add a new model link
            objModelLink = objModelLinks.Add(filename)

            ' Get a reference to the drawing views collection
            objDrawingViews = objSheet.DrawingViews


            objDrawingView = objDrawingViews.AddAssemblyView([From]:=objModelLink, Orientation:=SolidEdgeDraft.ViewOrientationConstants.igDimetricTopBackLeftView, Scale:=1.0, x:=0.4, y:=0.4, ViewType:=SolidEdgeDraft.AssemblyDrawingViewTypeConstants.seAssemblyDesignedView)

            ' Add a new drawing view
            'objDrawingView = objDrawingViews.AddAssemblyView(
            '  objModelLink,
            '  SolidEdgeDraft.ViewOrientationConstants.igFrontView,
            '  1,
            '  0.1,
            '  0.1,
            '  SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView)

            ' Assign a caption
            objDrawingView.Caption = "My New Drawing View"

            ' Ensure caption is displayed
            objDrawingView.DisplayCaption = False

        Catch ex As Exception
            Console.WriteLine(ex.Message)
        Finally
            ' OleMessageFilter.Revoke()
        End Try


        '  Dim objApp As SolidEdgeFramework.Application = Nothing
        Dim objDoc As SolidEdgeDraft.DraftDocument = Nothing
        Dim objPartsLists As SolidEdgeDraft.PartsLists = Nothing
        Dim objPartsList As SolidEdgeDraft.PartsList = Nothing

        ' objApp = GetObject(, "SolidEdge.Application")
        objDoc = objApp.ActiveDocument
        objPartsLists = objDoc.PartsLists

        '30th sep 2024
        'objPartsList = objPartsLists.Add(objDrawingView, "BEC RAW", 1, 1)
        objPartsList = objPartsLists.Add(objDrawingView, "BEC RAW", 1, 1)

        ' objSheet.Activate()

        ' objAsm.LinkedDocuments(DesignManager.LinkTypeConstants.seLinkTypeAll)

        objPartsList = objPartsLists.Item(1)
        '  Dim dt As DataTable = objPartsList
        Dim tableCell As SolidEdgeDraft.TableCell = Nothing
        Dim dt As New DataTable()

        Dim cols As TableColumns = objPartsList.Columns
        Dim rows As TableRows = objPartsList.Rows
        For Each tableColumn In cols.OfType(Of SolidEdgeDraft.TableColumn)()
            Dim dtcolums As New DataColumn With {
                .ColumnName = tableColumn.Header
            }
            dt.Columns.Add(dtcolums)


        Next tableColumn


        'For j As Integer = 1 To rows.Count
        '    Dim objrow As TableRow = rows.Item(j)

        '    tableCell = PartsList.Cell(TableRow.Index, TableColumn.Index)


        'Next

        Dim dtrows As DataRow = dt.NewRow()
        dt.Rows.Add(dtrows)

        For Each tableRow In rows.OfType(Of SolidEdgeDraft.TableRow)()
            For Each tableColumn In cols.OfType(Of SolidEdgeDraft.TableColumn)()
                If tableColumn.Show Then
                    'visibleRowCount += 1
                    tableCell = objPartsList.Cell(tableRow.Index, tableColumn.Index)
                    Dim rowindex As Integer = tableRow.Index - 1
                    Dim colindex As Integer = tableColumn.Index - 1
                    Dim tabvalue As String = tableCell.value

                    dt.Rows(rowindex).Item(colindex) = tabvalue
                    ' MsgBox(tableCell.value)
                    'excelRange = excelCells.Item(tableRow.Index + 1, tableColumn.Index)
                    'excelRange.Value = tableCell.value
                End If
            Next tableColumn
            dtrows = dt.NewRow()
            dt.Rows.Add(dtrows)
            ' visibleRowCount = 0

        Next tableRow


        ' objPartsList.CopyToClipboard()
        objApp.Documents.CloseDocument(objDoc.FullName, False, "", False, False)

        Return dt

    End Function

    Private Sub BtnTemplateLocation_Click(sender As Object, e As EventArgs) Handles btnTemplateLocation.Click
        Try
            Dim reportDir As String = ""

            Try
                Dim BetterFolderBrowser As New BetterFolderBrowser With {
                    .Title = "Select folders",
                    .RootFolder = "C:\\",
                    .Multiselect = False
                }
                If BetterFolderBrowser.ShowDialog() = DialogResult.OK Then
                    reportDir = BetterFolderBrowser.SelectedFolder
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

            txtRawMaterialEstimationReportDirPath.Text = reportDir

        Catch ex As Exception

        End Try
    End Sub

    Private Sub BtnBrowseRawMaterialBOM_Click(sender As Object, e As EventArgs) Handles btnBrowseRawMaterialBOM.Click
        Try
            Using dialog As New OpenFileDialog
                dialog.Filter = "Excel files (*.xlsx)|*.xlsx"
                If dialog.ShowDialog() <> DialogResult.OK Then Return
                txtBecMaterialExcel.Text = dialog.FileName
            End Using
        Catch ex As Exception

        End Try
    End Sub
End Class