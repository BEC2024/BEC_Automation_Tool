Imports System.Runtime.InteropServices
'Imports SolidEdge.Framework.Interop
Imports SolidEdgeFramework
Public Class PartAutomationForm

    Dim dt As New DataTable("")
    Dim custPropertiesObj As New CustomProperties()
    Dim objApp As SolidEdgeFramework.Application = Nothing
    Dim objMatTable As SolidEdgeFramework.MatTable = Nothing
    Dim listOfLibraries As Object = Nothing
    Dim numMaterials As Long
    Dim objSheetMetalDocument As SolidEdgePart.SheetMetalDocument = Nothing
    Dim objDocument As SolidEdgePart.PartDocument = Nothing
    Dim objAssemblyDocument As SolidEdgeAssembly.AssemblyDocument
    Dim dicMaterials As New Dictionary(Of String, List(Of String))()
    Dim isSheetMetalPart As Boolean = False
    ' Dim isPartTypeWise As Boolean = False

    Dim isBrowse As Boolean = False

    Dim dicData As New Dictionary(Of String, DataSet)()

    Dim dtstructure As New DataTable("Structure")
    Dim dtsheetmetal As New DataTable("SheetMetal")

    'version 1.0.12
    'Change the code to set the gage using the Material library.

    'version 1.0.13
    'Add the material library selection defaul is BEC Matarial Library

    Private Sub SetBendType(ByVal dt As DataTable)
        Try
            Dim dv As New DataView(dt)
            Dim partTypeCol As String = ExcelUtil.ExcelSheetColumns.Type.ToString()
            Dim sizeCol As String = ExcelUtil.ExcelSheetColumns.Size.ToString()
            Dim gradeCol As String = ExcelUtil.ExcelSheetColumns.Grade.ToString()
            Dim thicknessCol As String = ExcelUtil.ExcelSheetColumns.Thickness.ToString()
            Dim materialUsedCol As String = ExcelUtil.ExcelSheetColumns.Material_Used.ToString()
            Dim bendRadiusCol As String = ExcelUtil.ExcelSheetColumns.Bend_Radius.ToString()
            Dim materialSpecCol As String = ExcelUtil.ExcelSheetColumns.Material_Specification.ToString()
            Dim becMaterialCol As String = ExcelUtil.ExcelSheetColumns.BEC_Material.ToString()
            'Temp18APR2023
            Dim filter As String = $"Convert([{materialUsedCol}], 'System.String') = '{txtMaterialUsed_C.Text}' And Convert([{bendRadiusCol}], 'System.String') = '{txtBendRadius_C.Text}'"
            dv.RowFilter = filter '$"{materialUsedCol}='{txtMaterialUsed_C.Text}' And {bendRadiusCol}='{txtBendRadius_C.Text}'"
            Dim dtSize As DataTable = dv.ToTable()
            For Each drv As DataRowView In dv
                cmbBendTypeGageWise_Pw.Text = drv(ExcelUtil.ExcelSheetColumns.Bend_Type.ToString)
                Exit For
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SetPartTypeWiseDetails(ByVal dt As DataTable)

        Dim dv As New DataView(dt)

        Dim partTypeCol As String = ExcelUtil.ExcelSheetColumns.Type.ToString()
        Dim sizeCol As String = ExcelUtil.ExcelSheetColumns.Size.ToString()
        Dim gradeCol As String = ExcelUtil.ExcelSheetColumns.Grade.ToString()
        Dim thicknessCol As String = ExcelUtil.ExcelSheetColumns.Thickness.ToString()
        Dim materialUsedCol As String = ExcelUtil.ExcelSheetColumns.Material_Used.ToString()
        Dim materialSpecCol As String = ExcelUtil.ExcelSheetColumns.Material_Specification.ToString()
        Dim becMaterialCol As String = ExcelUtil.ExcelSheetColumns.BEC_Material.ToString()
        If cmbCategory.Text = "SheetMetal" Then
            cmbMaterialUsed2_Mw.Items.Clear()
            For i = 0 To dt.Rows.Count - 1
                'MsgBox(dt.Rows(i)(4).ToString())
                If Not cmbMaterialUsed2_Mw.Items.Contains(dt.Rows(i)(4).ToString()) Then
                    cmbMaterialUsed2_Mw.Items.Add(dt.Rows(i)(4).ToString())
                End If
            Next
            'Temp24APR2023
            If txtMaterialUsed_C.Text = "" Then
                cmbMaterialUsed2_Mw.Text = cmbMaterialUsed2_Mw.Items(0)
            Else
                Dim MaterialUsed = txtMaterialUsed_C.Text
                If cmbMaterialUsed2_Mw.Items.Contains(MaterialUsed) Then
                    cmbMaterialUsed2_Mw.Text = txtMaterialUsed_C.Text
                Else
                    cmbMaterialUsed2_Mw.Text = cmbMaterialUsed2_Mw.Items(0)
                End If

            End If

            End If

        'Temp18APR2023
        'Dim filter As String = $"Convert([{materialUsedCol}], 'System.String') = '{txtMaterialUsed_C.Text}'"
        Dim filter As String = $"Convert([{materialUsedCol}], 'System.String') = '{cmbMaterialUsed2_Mw.Text}'"
        dv.RowFilter = filter
        'dv.RowFilter = $"Material_Used = '{cmbMaterialUsed2_Mw.Text}'"
        Dim dtSize As DataTable = dv.ToTable()
        For Each drv As DataRowView In dv

            'Dim thickness1 As String = drv(ExcelUtil.excelSheetColumns.Thickness.ToString)

            'If Not cmbThickness.Items.Contains(thickness1) Then
            '    cmbThickness.Items.Add(thickness1)
            'End If

            cmbPartType_Pw.Text = drv(ExcelUtil.ExcelSheetColumns.Type.ToString)
            If drv(ExcelUtil.ExcelSheetColumns.Size.ToString) IsNot Nothing Then
                cmbSize_Pw.Text = drv(ExcelUtil.ExcelSheetColumns.Size.ToString).ToString()
            Else
                cmbSize_Pw.Text = String.Empty
            End If

            If drv(ExcelUtil.ExcelSheetColumns.Grade.ToString) IsNot Nothing Then
                cmbGrade_Pw.Text = drv(ExcelUtil.ExcelSheetColumns.Grade.ToString).ToString()
            Else
                cmbGrade_Pw.Text = String.Empty
            End If

            If drv(ExcelUtil.ExcelSheetColumns.Thickness.ToString) IsNot Nothing Then
                cmbThickness_Pw.Text = drv(ExcelUtil.ExcelSheetColumns.Thickness.ToString).ToString()
            Else
                cmbThickness_Pw.Text = String.Empty
            End If

            If drv(ExcelUtil.ExcelSheetColumns.Material_Used.ToString) IsNot Nothing Then
                cmbMaterialUsed_Pw.Text = drv(ExcelUtil.ExcelSheetColumns.Material_Used.ToString).ToString()
            Else
                cmbMaterialUsed_Pw.Text = String.Empty
            End If

            If drv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString) IsNot Nothing Then
                cmbMaterialSpec_Pw.Text = drv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString).ToString()
            Else
                cmbMaterialSpec_Pw.Text = String.Empty
            End If

            If drv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString) IsNot Nothing Then
                txtBECMaterial_Pw.Text = drv(ExcelUtil.ExcelSheetColumns.BEC_Material.ToString).ToString()
            Else
                txtBECMaterial_Pw.Text = String.Empty
            End If

            Exit For

        Next

        SetBendType(dt)
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

            '    FillCategory(dt)

            '    'FillPartType(dt)

            '    'FillMaterialUsed2(dt)

            '    'If cmbMaterialUsed2_Mw.Items.Contains(txtMaterialUsed_C.Text) Then
            '    '    cmbMaterialUsed2_Mw.Text = txtMaterialUsed_C.Text
            '    'End If

            '    'SetPartTypeWiseDetails(dt)

            '    'rbPartTypewise.Enabled = True
            '    'rbMaterialWise.Enabled = True
            '    'btnApply.Enabled = True

            '    'rbMaterialWise.Checked = True
            '    'isBrowse = True

            '    'Validation()
            'Else
            '    MsgBox("Please select material excel details.")
            'End If


        Catch ex As Exception
        End Try
    End Sub

    Private Sub FillMaterialLibList(ByVal listOfLibraries As Object)

        For Each lib1 As String In listOfLibraries
            cmbMaterialLib.Items.Add(lib1)
        Next

        If cmbMaterialLib.Items.Contains("BEC MATERIAL LIBRARY") Then
            cmbMaterialLib.Text = "BEC MATERIAL LIBRARY"
        Else

            If cmbMaterialLib.Items.Count > 0 Then
                cmbMaterialLib.SelectedItem = cmbMaterialLib.Items(0)
            End If

        End If
    End Sub

    Private Sub Test()
        Dim document As SolidEdgeAssembly.AssemblyDocument = Nothing
        document = objApp.GetActiveDocument(Of SolidEdgeAssembly.AssemblyDocument)(False)

    End Sub

    Public Sub Closefn(mainObj As MainClass)
        mainObj.SolidEdgeinstance = "Close"
    End Sub

    Dim mainObj As New MainClass

    Private Sub DisableBtn()
        If objApp Is Nothing Then
            gpPartProperties.Enabled = False
            gpDefaultMaterialExcel.Enabled = False
            btnApply.Enabled = False
            btnClose.Enabled = False
            txtGageTable.Enabled = False
            btnShowGuideLines.Enabled = False
            btnRefresh.Enabled = False
        Else
            gpPartProperties.Enabled = True
            gpDefaultMaterialExcel.Enabled = True
            btnApply.Enabled = True
            btnClose.Enabled = True
            txtGageTable.Enabled = True
            btnShowGuideLines.Enabled = True
            btnRefresh.Enabled = True
        End If

    End Sub
    Private Function IsValid() As Boolean

        SetSolidEdgeInstance()
        DisableBtn()
        If objApp Is Nothing Then
            Return False
        Else
            Return True

        End If

    End Function
    Public Sub GetFormLoadData()


        ReadCustomProperties()

        FillCustomProperties(custPropertiesObj)

        dicMaterials = GetMaterialCollection(listOfLibraries)


        Dim materialName As String = GetCurrentMaterialName(objSheetMetalDocument, objDocument, isSheetMetalPart)

        Dim currentGageName As String = GetCurrentGageName(objSheetMetalDocument, objDocument, isSheetMetalPart)

        txtCurrentMaterial_C.Text = materialName

        txtGageName_C.Text = currentGageName

        'Gauge_RND()

        rbPartTypewise.Enabled = False
        ' rbMaterialWise.Enabled = False
        btnApply.Enabled = False

        'Dim res As Boolean = objMatTable.PerformGageDataValidation("C:\Program Files\Siemens\Solid Edge 2022\Preferences\Gagetable.xls", "aluminum 6061", "20 Gage Al")

        'If res Then

        Try
            'objMatTable.ApplyGageFromGageTableToDoc(objSheetMetalDocument, "20 Gage Al", "aluminum 6061")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message")
            Debug.Print($"{ex.Message}{vbNewLine}{ex.StackTrace}")
        End Try

        ' End If

        'Dim bstrGageTableName As String
        'Dim plNumGages As Integer
        'Dim listOfGages As Object
        'objMatTable.GetPSMGaugeListFromExcel(bstrGageTableName, plNumGages, listOfGages)
        'Debug.Print("aaa")

        'objMatTable.ApplyGageFromGageTableToDoc(objSheetMetalDocument, "10 Gage Al", "aluminum 6061")

        'objMatTable.ApplyGageFromLibraryToDoc(objSheetMetalDocument, "14 gage", "BEC MATERIAL LIBRARY")
        '
        'GetPSMGaugeListFromExcel(bstrGageTableName As String, ByRef plNumGages As Integer, ByRef listOfGages As Object)
        ' objMatTable.ApplyGageFromGageTableToDoc(objSheetMetalDocument, "SHEET - 28 GA1", "Test") 'BEC MATERIAL LIBRARY")

        SketchRND()
    End Sub
    Private Sub PartAutomationForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        CustomLogUtil.Heading("Part/ Sheet-Metal Update Form Open.....")

        If IsValid() Then

            txtExcelPath.Text = Config.configObj.becMaterialExcelPath

            SetSolidEdgeInstance()

            If Not mainObj.SolidEdgeinstance = "Close" Then
                SetMaterialTable()

                GetMaterialLibraryList()

                FillMaterialLibList(listOfLibraries)
                'TEMP27APR2023
                GetFormLoadData()

            End If
        Else
            MessageBox.Show("Please open Solid-Edge Part Or SheetMetal and restart The Application", "Message")
            CustomLogUtil.Log("Please open Solid-Edge Part Or SheetMetal and restart The Application", "", "")
        End If

    End Sub

    Private Sub SketchRND()
        Try
            If objDocument Is Nothing Then
            Else

                Dim sketches As SolidEdgePart.Sketchs = objDocument.Sketches
                For Each sk As SolidEdgePart.Sketch In sketches

                    If sk.DisplayName = "Sketch_4" Then

                        Dim pf As Object = sk.Profile
                        Debug.Print("aaaa")
                    End If
                Next
            End If



        Catch ex As Exception

        End Try

    End Sub

    Private Sub RefreshForm()

        'SetSolidEdgeInstance()

        'SetMaterialTable()

        'GetMaterialLibraryList()

        ReadCustomProperties()

        FillCustomProperties(custPropertiesObj)

        dicMaterials = GetMaterialCollection(listOfLibraries)

        Dim materialName As String = GetCurrentMaterialName(objSheetMetalDocument, objDocument, isSheetMetalPart)

        txtCurrentMaterial_C.Text = materialName

        Dim currentGageName As String = GetCurrentGageName(objSheetMetalDocument, objDocument, isSheetMetalPart)

        txtGageName_C.Text = currentGageName

        If GlobalEntity.dictRawMaterials.Count = 0 Then
            GlobalEntity.dictRawMaterials = ExcelUtil.ReadRawMaterials2(txtExcelPath.Text)
        End If


        '===
        'dt = GlobalEntity.dictRawMaterials("BECMaterials").Tables(0)

        'FillPartType(dt)

        'FillMaterialUsed2(dt)

        'If cmbMaterialUsed2_Mw.Items.Contains(txtMaterialUsed_C.Text) Then
        '    cmbMaterialUsed2_Mw.Text = txtMaterialUsed_C.Text
        'End If

        'SetPartTypeWiseDetails(dt)
        '===
        FillPartType(dt)

        FillMaterialUsed2(dt)

        If cmbMaterialUsed2_Mw.Items.Contains(txtMaterialUsed_C.Text) Then
            cmbMaterialUsed2_Mw.Text = txtMaterialUsed_C.Text
        End If

        SetPartTypeWiseDetails(dt)


        rbPartTypewise.Enabled = True
        rbMaterialWise.Enabled = True
        btnApply.Enabled = True

        rbMaterialWise.Checked = True
        isBrowse = True

        Validation()
    End Sub

    Private Sub Gauge_RND()
        Try

            Dim currentGaugeName As String = String.Empty
            objMatTable.GetCurrentGageName(objSheetMetalDocument, currentGaugeName)

            Dim defaultGaugeFileName As String = String.Empty
            objMatTable.GetDefaultGageFileName(defaultGaugeFileName)

            Dim res As Boolean = objMatTable.PerformGageDataValidation(defaultGaugeFileName, "Test", "SHEET - 28 GA1")
            'Dim gaugeTableName As String = String.Empty
            'objMatTable.GetPSMGaugeListFromExcel()

            'Dim listOfGages As Object
            'Dim plNumGages As Integer
            'objMatTable.GetPSMGaugeListFromExcel(bstrGageTableName As String, ByRef plNumGages As Integer, ByRef listOfGages As Object)

            'objMatTable.set
            'objMatTable.ApplyGageFromGageTableToDoc(objSheetMetalDocument, "SHEET - 28 GA - STEEL (A36,A242,A572,A588)", "Table")
            Debug.Print("aaaa")

            Dim gaugeTableName2 As String = String.Empty
            Dim numGauges As Integer
            Dim lstGauges As New List(Of String)()
            objMatTable.GetPSMGaugeListFromExcel("Table", numGauges, lstGauges)
            Debug.Print("aaa")

            'Dim gageName As String
            'Dim gageFilePath As String

            'objMatTable.GetPSMGaugeInfoForDoc(objSheetMetalDocument, gageName, gageFilePath)
        Catch ex As Exception

        End Try
    End Sub

    Private Function GetCurrentMaterialName(ByVal sheetMetalDoc As SolidEdgePart.SheetMetalDocument, ByVal partDoc As SolidEdgePart.PartDocument, ByVal isSheetMetal As Boolean) As String

        Dim materialName As String = String.Empty
        Try
            If isSheetMetal Then
                objMatTable.GetCurrentMaterialName(sheetMetalDoc, materialName)
            Else

                If partDoc Is Nothing Then
                    MessageBox.Show("Please open Part Document", "Message")


                Else
                    objMatTable.GetCurrentMaterialName(partDoc, materialName)
                End If



            End If


        Catch ex As Exception
            MessageBox.Show($"Error While fetching GetCurrentMaterialName {vbNewLine}{ex.Message}{ ex.StackTrace}", "Message")

        End Try

        Return materialName

    End Function

    Private Function GetCurrentGageName(ByVal sheetMetalDoc As SolidEdgePart.SheetMetalDocument, ByVal partDoc As SolidEdgePart.PartDocument, ByVal isSheetMetal As Boolean) As String

        Dim materialName As String = String.Empty

        If isSheetMetal Then
            objMatTable.GetCurrentGageName(sheetMetalDoc, materialName)
        Else
            If partDoc Is Nothing Then


            Else
                objMatTable.GetCurrentGageName(partDoc, materialName)
            End If

        End If

        Return materialName

    End Function

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

    Public Sub FillCustomProperties(ByVal custPropObj As CustomProperties)
        Try

            txtPartType_C.Text = custPropObj.partType
            txtSize_C.Text = custPropObj.size
            txtGrade_C.Text = custPropObj.grade
            txtThickness_C.Text = custPropObj.materialThickness
            txtMaterialUsed_C.Text = custPropObj.materialUsed
            txtMaterialSpec_C.Text = custPropObj.materialSpec
            txtBendRadius_C.Text = custPropertiesObj.bendRadius.ToString()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub GetMaterialLibraryList()

        objMatTable.GetMaterialLibraryList(listOfLibraries, numMaterials)

    End Sub

    Private Sub SetMaterialTable()
        objMatTable = objApp.GetMaterialTable()
    End Sub

    Private Sub SetSolidEdgeInstance()
        Try
            objApp = Marshal.GetActiveObject("SolidEdge.Application")
        Catch ex As Exception
            'MessageBox.Show($"Error in fetching the Solid-Edge instance {vbNewLine}{ex.Message}{ ex.StackTrace}", "Message")
            'CustomLogUtil.Log("While fetching the Solid-Edge instance ", ex.Message, ex.StackTrace)

            'Closefn(mainObj)
        End Try

    End Sub

    Private Function GetPartType() As String

        Dim doc As Object = objApp.ActiveDocument
        Dim fullName As String = doc.FullName
        Return fullName
    End Function

    Private Function IsSheetMetalDocument(ByVal docPath As String) As Boolean

        Dim isSheetMetalPart As Boolean = False

        If IO.Path.GetFileName(docPath).ToUpper.EndsWith(".PSM") Then

            isSheetMetalPart = True

        End If

        Return isSheetMetalPart

    End Function

    Private Sub FillBendType__AccordingToPriority()
        Try

            Dim dv As DataView = New DataView(dt)
            Dim typeCol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Type.ToString()
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


                dv.RowFilter = $"{Becmaterialusedcol}='{cmbMaterialUsed2_Mw.Text}' and {Becmaterialspecol}='{txtMaterialSpec2_Mw.Text}'and {BECMaterialcol}='{txtBECMaterial2_Mw.Text}' and {Thicknesscol}='{txtThickness2_Mw.Text}' "

                cmbBendTypeGageWise_Mw.Items.Clear()

                For Each drv As DataRowView In dv
                    If Priority = 0 Then
                        Priority = Convert.ToInt32(drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Priority.ToString).ToString())
                    End If
                    value = Convert.ToInt32(drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Priority.ToString).ToString())


                    If (Priority > value) Then
                        Priority = value
                    End If
                    Dim BendType As String = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Bend_Type.ToString).ToString()
                    If Not cmbBendTypeGageWise_Mw.Items.Contains(BendType) Then
                        cmbBendTypeGageWise_Mw.Items.Add(BendType)
                    End If

                Next

                If dv.Count = 1 And cmbBendTypeGageWise_Mw.Items.Count > 0 Then
                    cmbBendTypeGageWise_Mw.SelectedItem = cmbBendTypeGageWise_Mw.Items(0)
                Else
                    dv.RowFilter = $"{Becmaterialusedcol}='{cmbMaterialUsed2_Mw.Text}' and {Becmaterialspecol}='{txtMaterialSpec2_Mw.Text}'and {BECMaterialcol}='{txtBECMaterial2_Mw.Text}' and {Thicknesscol}='{txtThickness2_Mw.Text}' and {Priorityspecol}='{Priority}'"
                    For Each drv As DataRowView In dv
                        Dim BendType As String = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Bend_Type.ToString).ToString()

                        cmbBendTypeGageWise_Mw.SelectedItem = BendType
                        txtSize2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Size.ToString).ToString()
                        txtGrade2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Grade.ToString).ToString()
                        txtGageTable.Text = drv(ExcelUtil.ExcelSheetColumns.Gage_Table.ToString).ToString()
                        txtThickness2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Thickness.ToString).ToString()
                        txtBendRadius_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Bend_Radius.ToString).ToString()
                        txtPartType_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Type.ToString).ToString()
                        txtMaterialSpec2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString).ToString()
                        txtBECMaterial2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.BEC_Material.ToString).ToString()
                    Next

                End If
            End If
        Catch ex As Exception
            MessageBox.Show($"While fetching FillGageName", "Error")
            CustomLogUtil.Log("While fetching FillGageName", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Public Sub ReadCustomProperties()

        Try

            custPropertiesObj = New CustomProperties()

            Dim documentPath As String = GetPartType()

            isSheetMetalPart = IsSheetMetalDocument(documentPath)

            If isSheetMetalPart Then

                objSheetMetalDocument = objApp.ActiveDocument
            Else

                objDocument = objApp.ActiveDocument

            End If

            If objDocument IsNot Nothing Then
#Region "Custom Porperty"
                Dim propSets As SolidEdgeFramework.PropertySets = objDocument.Properties

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

                        End If
                    Catch ex As Exception
                    End Try

                Next
#End Region

#Region "Variable Table"'TEMP12SEPT2023
                '5th Nov 2024
                Try
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
                Catch ex As Exception

                End Try
#End Region



            End If

            If objSheetMetalDocument IsNot Nothing Then
#Region "Custom Property"


                Dim propSets As SolidEdgeFramework.PropertySets = objSheetMetalDocument.Properties

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

                            custPropertiesObj.bendRadius = prop1.Value.ToString()

                        End If
                    Catch ex As Exception
                    End Try

                Next
#End Region
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
                End If
        Catch ex As Exception

        End Try
    End Sub

    Public Sub FillMaterialUsed2(ByVal dt As DataTable)

        Try

            cmbMaterialUsed2_Mw.Items.Clear()

            Dim materialUsedCol As String = ExcelUtil.ExcelSheetColumns.Material_Used.ToString()
            Dim becMaterialCol As String = ExcelUtil.ExcelSheetColumns.BEC_Material.ToString()
            Dim materialThicknessCol As String = ExcelUtil.ExcelSheetColumns.BEC_Material.ToString()

            Dim categoryCol As String = ExcelUtil.ExcelSheetColumns.Category.ToString()
            Dim dv As New DataView(dt) With {
            .RowFilter = $"{categoryCol}='{cmbCategory.Text}'"
            }

            Dim materialUsedList As New List(Of String)()
            Dim dtSize As DataTable = dv.ToTable()
            For Each drv As DataRowView In dv

                Dim materialUsed1 As String = drv(ExcelUtil.ExcelSheetColumns.Material_Used.ToString)

                If Not materialUsedList.Contains(materialUsed1) Then
                    materialUsedList.Add(drv(ExcelUtil.ExcelSheetColumns.Material_Used.ToString))
                End If

            Next

            'materialUsedList = dt.AsEnumerable() _
            '                                   .Select(Function(r) r.Field(Of String)(materialUsedCol)) _
            '                                   .Distinct() _
            '                                   .ToList()

            For Each materialUsedName As String In materialUsedList

                If Not cmbMaterialUsed2_Mw.Items.Contains(materialUsedName) Then
                    cmbMaterialUsed2_Mw.Items.Add(materialUsedName)
                End If

            Next

            If cmbMaterialUsed2_Mw.Items.Count > 0 Then
                'Temp24APR2023
                If txtMaterialUsed_C.Text = "" Then
                    cmbMaterialUsed2_Mw.SelectedItem = cmbMaterialUsed2_Mw.Items(0)
                Else
                    cmbMaterialUsed2_Mw.SelectedItem = txtMaterialUsed_C.Text
                End If

                'Else
                '    cmbMaterialUsed2_Mw.SelectedItem = txtMaterialSpec_C.Text
            End If

            Dim mySource As New AutoCompleteStringCollection()
            mySource.AddRange(materialUsedList.ToArray)
            cmbMaterialUsed2_Mw.AutoCompleteMode = AutoCompleteMode.SuggestAppend
            cmbMaterialUsed2_Mw.AutoCompleteSource = AutoCompleteSource.CustomSource
            cmbMaterialUsed2_Mw.AutoCompleteCustomSource = mySource
        Catch ex As Exception

        End Try

    End Sub

    Public Sub FillCategory(Optional ByVal dt As DataTable = Nothing)

        cmbCategory.Items.Clear()
        cmbPartType_Pw.Items.Clear()
        cmbSize_Pw.Items.Clear()
        cmbGrade_Pw.Items.Clear()
        cmbGageName_Pw.Items.Clear()
        txtGageTable.Text = String.Empty
        cmbThickness_Pw.Items.Clear()
        cmbMaterialUsed_Pw.Items.Clear()
        cmbMaterialSpec_Pw.Items.Clear()

        'Dim categoryCol As String = ExcelUtil.ExcelSheetColumns.Category.ToString()

        'Dim categoryList As List(Of String) = dt.AsEnumerable() _
        '                                   .Select(Function(r) r.Field(Of String)(categoryCol)) _
        '                                   .Distinct() _
        '                                   .ToList()

        Dim categoryList As New List(Of String) From {
            "Structure",
            "SheetMetal"
        }

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

    End Sub

    Public Sub FillPartType(ByVal dt As DataTable)

        Try
            cmbPartType_Pw.Items.Clear()
            cmbSize_Pw.Items.Clear()
            cmbGrade_Pw.Items.Clear()
            cmbGageName_Pw.Items.Clear()
            txtGageTable.Text = String.Empty
            cmbThickness_Pw.Items.Clear()
            cmbMaterialUsed_Pw.Items.Clear()
            cmbMaterialSpec_Pw.Items.Clear()

            Dim categoryCol As String = ExcelUtil.ExcelSheetColumns.Category.ToString()
            Dim dv As New DataView(dt) With {
                .RowFilter = $"{categoryCol}='{cmbCategory.Text}'"
            }
            Dim dtSize As DataTable = dv.ToTable()
            For Each drv As DataRowView In dv

                Dim size1 As String = drv(ExcelUtil.ExcelSheetColumns.Type.ToString)

                If Not cmbPartType_Pw.Items.Contains(size1) Then
                    cmbPartType_Pw.Items.Add(drv(ExcelUtil.ExcelSheetColumns.Type.ToString))
                End If

            Next

            If cmbPartType_Pw.Items.Count > 0 Then
                cmbPartType_Pw.SelectedItem = cmbPartType_Pw.Items(0)
            End If

            'Dim partTypeList As List(Of String) = dt.AsEnumerable() _
            '                                   .Select(Function(r) r.Field(Of String)(partTypeCol)) _
            '                                   .Distinct() _
            '                                   .ToList()

            'For Each partName As String In partTypeList

            '    If Not cmbPartType_Pw.Items.Contains(partName) Then
            '        cmbPartType_Pw.Items.Add(partName)
            '    End If

            'Next

            'If cmbPartType_Pw.Items.Count > 0 Then
            '    cmbPartType_Pw.SelectedItem = cmbPartType_Pw.Items(0)
            'End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub CmbPartType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPartType_Pw.SelectedIndexChanged

        FillPartSize(dt, cmbPartType_Pw.Text)

        Validation()

    End Sub

    Private Sub FillPartSize(ByVal dt As DataTable, ByVal partType As String)

        Try

            cmbSize_Pw.Items.Clear()
            cmbGrade_Pw.Items.Clear()
            cmbThickness_Pw.Items.Clear()
            txtBECMaterial_Pw.Text = ""

            cmbGageName_Pw.Items.Clear()
            cmbMaterialUsed_Pw.Items.Clear()
            cmbMaterialSpec_Pw.Items.Clear()
            cmbBendTypeGageWise_Pw.Items.Clear()
            Dim dv As New DataView(dt)
            Dim partTypeCol As String = ExcelUtil.ExcelSheetColumns.Type.ToString()
            Dim sizeCol As String = ExcelUtil.ExcelSheetColumns.Size.ToString()
            Dim categoryCol As String = ExcelUtil.ExcelSheetColumns.Category.ToString()
            Dim MaterialUsedCol As String = ExcelUtil.ExcelSheetColumns.Material_Used.ToString()
            Dim MaterialSpecCol As String = ExcelUtil.ExcelSheetColumns.Material_Specification.ToString
            Dim BecMaterialCol As String = ExcelUtil.ExcelSheetColumns.BEC_Material.ToString


            'SheetMetal
            Dim Grade As String = ExcelUtil.ExcelSheetColumns.Grade.ToString

            Dim gradeCol As String = ExcelUtil.ExcelSheetColumns.Grade.ToString()
            Dim bendTypeCol As String = ExcelUtil.ExcelSheetColumns.Bend_Type.ToString()
            Dim gageNameCol As String = ExcelUtil.ExcelSheetColumns.Gage_Name.ToString()
            Dim thicknessCol As String = ExcelUtil.ExcelSheetColumns.Thickness.ToString()
            Dim bendRadiusCol As String = ExcelUtil.ExcelSheetColumns.Bend_Radius.ToString()


            Dim ListOfMaterialUsed_Pw As New List(Of String)
            Dim ListOfMaterialSpec_Pw As New List(Of String)
            Dim ListOfBECMaterial_Pw As New List(Of String)
            Dim listOfGrade_Pw As New List(Of String)
            Dim none As String = "NONE"
            If cmbPartType_Pw.Text = none Then
                cmbSize_Pw.Items.Add(none)
                cmbGrade_Pw.Items.Add(none)
                cmbBendRadius_Pw.Items.Add(none)
                cmbGageName_Pw.Items.Add(none)
                cmbThickness_Pw.Items.Add(none)
                cmbBendTypeGageWise_Pw.Items.Add(none)
                cmbMaterialUsed_Pw.Items.Add(none)
                cmbMaterialSpec_Pw.Items.Add(none)


                cmbPartType_Pw.SelectedItem = cmbPartType_Pw.Items(0)
                cmbSize_Pw.SelectedItem = cmbSize_Pw.Items(0)
                cmbGrade_Pw.SelectedItem = cmbGrade_Pw.Items(0)
                cmbBendTypeGageWise_Pw.SelectedItem = cmbBendTypeGageWise_Pw.Items(0)

                If cmbGageName_Pw.Items.Count = "0" Then
                    cmbGageName_Pw.Items.Add(none)
                    cmbGageName_Pw.SelectedItem = cmbGageName_Pw.Items(0)
                End If

                If cmbThickness_Pw.Items.Count = "0" Then
                    cmbThickness_Pw.Items.Add(none)

                End If
                cmbThickness_Pw.SelectedItem = cmbThickness_Pw.Items(0)

                If cmbBendRadius_Pw.Items.Count = "0" Then
                    cmbBendRadius_Pw.Items.Add(none)
                End If
                cmbBendRadius_Pw.SelectedItem = cmbBendRadius_Pw.Items(0)
                If cmbMaterialUsed_Pw.Items.Count = "0" Then
                    cmbMaterialUsed_Pw.Items.Add(none)
                End If
                cmbMaterialUsed_Pw.SelectedItem = cmbMaterialUsed_Pw.Items(0)
                If cmbMaterialSpec_Pw.Items.Count = "0" Then
                    cmbMaterialSpec_Pw.Items.Add(none)
                End If
                cmbMaterialSpec_Pw.SelectedItem = cmbMaterialSpec_Pw.Items(0)

                txtBECMaterial_Pw.Text = none
            Else
                dv.RowFilter = $"{partTypeCol}='{cmbPartType_Pw.Text}'"
                Dim dtSize As DataTable = dv.ToTable()
                Dim i As Integer = 0
                For Each drv As DataRowView In dv
                    i = i + 1
                    Dim size1 As String = drv(ExcelUtil.ExcelSheetColumns.Size.ToString).ToString()

                    If Not cmbSize_Pw.Items.Contains(size1) Then
                        cmbSize_Pw.Items.Add(size1)
                    End If

                    'Dim MaterialUsed1 As String = drv(ExcelUtil.ExcelSheetColumns.Material_Used.ToString).ToString()
                    'If Not cmbMaterialUsed_Pw.Items.Contains(MaterialUsed1) Then
                    '    cmbMaterialUsed_Pw.Items.Add(drv(ExcelUtil.ExcelSheetColumns.Material_Used.ToString))
                    'End If
                    'Dim MaterialSpec1 As String = drv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString).ToString()
                    'If Not cmbMaterialSpec_Pw.Items.Contains(MaterialSpec1) Then
                    '    cmbMaterialSpec_Pw.Items.Add(drv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString))
                    'End If
                    ListOfMaterialUsed_Pw.Add(drv(ExcelUtil.ExcelSheetColumns.Material_Used.ToString).ToString)
                    ListOfMaterialSpec_Pw.Add(drv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString).ToString)
                    ListOfBECMaterial_Pw.Add(drv(ExcelUtil.ExcelSheetColumns.BEC_Material.ToString).ToString)

                    'For Structure
                    If cmbCategory.Text = "Structure" Then
                        Dim BecMaterial1 As String = (drv(ExcelUtil.ExcelSheetColumns.BEC_Material.ToString).ToString)
                        txtBECMaterial_Pw.Text = BecMaterial1
                    End If


                    If cmbCategory.Text = "SheetMetal" Then
                        Dim Grade1 As String = drv(ExcelUtil.ExcelSheetColumns.Grade.ToString).ToString()
                        If Not cmbGrade_Pw.Items.Contains(Grade1) Then
                            listOfGrade_Pw.Add(drv(ExcelUtil.ExcelSheetColumns.Grade.ToString).ToString)
                            cmbGrade_Pw.Items.Add(Grade1)
                        End If
                    End If


                Next

                If cmbSize_Pw.Items.Count > 0 Then
                    cmbSize_Pw.SelectedItem = cmbSize_Pw.Items(0)
                End If

                ' cmbGrade_Pw.SelectedItem = cmbGrade_Pw.Items(0)


                cmbMaterialUsed_Pw.SelectedItem = ListOfMaterialUsed_Pw(0)
                cmbMaterialSpec_Pw.SelectedItem = ListOfMaterialSpec_Pw(0)
                If cmbCategory.Text = "SheetMetal" Then
                    txtBECMaterial_Pw.Text = ListOfBECMaterial_Pw(0)

                    For i = 0 To listOfGrade_Pw.Count - 1
                        If Not cmbGrade_Pw.Items.Contains(listOfGrade_Pw(i)) Then
                            cmbGrade_Pw.Items.Add(listOfGrade_Pw(i))
                        End If
                    Next
                End If
                ' cmbGrade_Pw.Text = listOfGrade_Pw(0)

                'If cmbGrade_Pw.Items.Count > 0 Then
                '    cmbGrade_Pw.SelectedItem = cmbGrade_Pw.Items(0)
                'End If
            End If
            'dv.RowFilter = $"{partTypeCol}='{cmbPartType_Pw.Text}' And {categoryCol}='{cmbCategory.Text}'"

        Catch ex As Exception

        End Try
    End Sub
    Private Sub FillMaterialSpecDetails(ByVal dt As DataTable, ByVal partType As String, ByVal partSize As String)

        Try

            'cmbSize_Pw.Items.Clear()
            cmbGrade_Pw.Items.Clear()
            cmbGageName_Pw.Items.Clear()
            txtGageTable.Text = String.Empty
            cmbThickness_Pw.Items.Clear()
            'cmbMaterialUsed_Pw.Items.Clear()
            cmbMaterialSpec_Pw.Items.Clear()
            txtBECMaterial_Pw.Text = String.Empty
            Dim dv As New DataView(dt)
            Dim partTypeCol As String = ExcelUtil.ExcelSheetColumns.Type.ToString()
            Dim sizeCol As String = ExcelUtil.ExcelSheetColumns.Size.ToString()
            Dim categoryCol As String = ExcelUtil.ExcelSheetColumns.Category.ToString()
            Dim MaterialUsedCol As String = ExcelUtil.ExcelSheetColumns.Material_Used.ToString()
            Dim MaterialSpecCol As String = ExcelUtil.ExcelSheetColumns.Material_Specification.ToString
            Dim BecMaterialCol As String = ExcelUtil.ExcelSheetColumns.BEC_Material.ToString

            dv.RowFilter = $"{partTypeCol}='{cmbPartType_Pw.Text}' And {sizeCol}='{cmbSize_Pw.Text}' And {categoryCol}='{cmbCategory.Text}' And {MaterialUsedCol}='{cmbMaterialUsed_Pw.Text}'"
            Dim dtSize As DataTable = dv.ToTable()
            For Each drv As DataRowView In dv

                Dim MaterailSpec1 As String = drv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString).ToString
                cmbMaterialSpec_Pw.Items.Add(MaterailSpec1)
                cmbMaterialSpec_Pw.SelectedItem = cmbMaterialSpec_Pw.Items(0)

                If cmbSize_Pw.Text = "9X9X" And cmbPartType_Pw.Text = "SQUARE BAR" Then
                    Dim BecMaterial1 As String = (drv(ExcelUtil.ExcelSheetColumns.BEC_Material.ToString))
                    txtBECMaterial_Pw.Text = BecMaterial1
                Else
                    txtBECMaterial_Pw.Text = ""
                End If


                'End If

                'If Not cmbSize_Pw.Items.Contains(size1) Then
                '    cmbSize_Pw.Items.Add(drv(ExcelUtil.ExcelSheetColumns.Size.ToString))
                'End If

                'Dim MaterialUsed1 As String = drv(ExcelUtil.ExcelSheetColumns.Material_Used.ToString).ToString()
                'If Not cmbMaterialUsed_Pw.Items.Contains(MaterialUsed1) Then
                '    cmbMaterialUsed_Pw.Items.Add(drv(ExcelUtil.ExcelSheetColumns.Material_Used.ToString))
                'End If
                'Dim MaterialSpec1 As String = drv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString).ToString()
                'If Not cmbMaterialSpec_Pw.Items.Contains(MaterialSpec1) Then
                '    cmbMaterialSpec_Pw.Items.Add(drv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString))
                'End If

            Next

            'If cmbSize_Pw.Items.Count > 0 Then
            '    cmbSize_Pw.SelectedItem = cmbSize_Pw.Items(0)
            'End If

            If cmbMaterialSpec_Pw.Items.Count > 0 Then
                cmbMaterialSpec_Pw.SelectedItem = cmbMaterialSpec_Pw.Items(0)
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub FillPartWiseDetails(ByVal dt As DataTable, ByVal partType As String, ByVal partSize As String)

        Try

            'cmbSize_Pw.Items.Clear()
            cmbGrade_Pw.Items.Clear()
            cmbGageName_Pw.Items.Clear()
            txtGageTable.Text = String.Empty
            cmbThickness_Pw.Items.Clear()
            cmbMaterialUsed_Pw.Items.Clear()
            cmbMaterialSpec_Pw.Items.Clear()
            txtBECMaterial_Pw.Text = String.Empty
            Dim dv As New DataView(dt)
            Dim partTypeCol As String = ExcelUtil.ExcelSheetColumns.Type.ToString()
            Dim sizeCol As String = ExcelUtil.ExcelSheetColumns.Size.ToString()
            Dim categoryCol As String = ExcelUtil.ExcelSheetColumns.Category.ToString()
            Dim MaterialUsedCol As String = ExcelUtil.ExcelSheetColumns.Material_Used.ToString()
            Dim MaterialSpecCol As String = ExcelUtil.ExcelSheetColumns.Material_Specification.ToString
            Dim BecMaterialCol As String = ExcelUtil.ExcelSheetColumns.BEC_Material.ToString

            dv.RowFilter = $"{partTypeCol}='{cmbPartType_Pw.Text}' And {sizeCol}='{cmbSize_Pw.Text}' And {categoryCol}='{cmbCategory.Text}'"
            Dim dtSize As DataTable = dv.ToTable()
            For Each drv As DataRowView In dv

                Dim size1 As String = drv(ExcelUtil.ExcelSheetColumns.Size.ToString).ToString()
                'If partSize = size1 Then
                Dim MaterialUsed1 As String = drv(ExcelUtil.ExcelSheetColumns.Material_Used.ToString).ToString
                cmbMaterialUsed_Pw.Items.Add(MaterialUsed1)
                For i = 0 To cmbMaterialUsed_Pw.Items.Count - 1
                    Dim j = 0
                    If cmbMaterialUsed_Pw.Items(i) = MaterialUsed1 Then
                        j = j + 1
                    End If
                    If j > 1 Then
                        cmbMaterialUsed_Pw.Items.Remove(MaterialUsed1)
                    End If
                Next




                Dim MaterailSpec1 As String = drv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString).ToString
                cmbMaterialSpec_Pw.Items.Add(MaterailSpec1)
                cmbMaterialSpec_Pw.SelectedItem = cmbMaterialSpec_Pw.Items(0)

                If cmbSize_Pw.Text = "9X9X" And cmbPartType_Pw.Text = "SQUARE BAR" Then
                    Dim BecMaterial1 As String = (drv(ExcelUtil.ExcelSheetColumns.BEC_Material.ToString))
                    txtBECMaterial_Pw.Text = BecMaterial1
                Else
                    txtBECMaterial_Pw.Text = ""
                End If


                'End If

                'If Not cmbSize_Pw.Items.Contains(size1) Then
                '    cmbSize_Pw.Items.Add(drv(ExcelUtil.ExcelSheetColumns.Size.ToString))
                'End If

                'Dim MaterialUsed1 As String = drv(ExcelUtil.ExcelSheetColumns.Material_Used.ToString).ToString()
                'If Not cmbMaterialUsed_Pw.Items.Contains(MaterialUsed1) Then
                '    cmbMaterialUsed_Pw.Items.Add(drv(ExcelUtil.ExcelSheetColumns.Material_Used.ToString))
                'End If
                'Dim MaterialSpec1 As String = drv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString).ToString()
                'If Not cmbMaterialSpec_Pw.Items.Contains(MaterialSpec1) Then
                '    cmbMaterialSpec_Pw.Items.Add(drv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString))
                'End If

            Next

            'If cmbSize_Pw.Items.Count > 0 Then
            '    cmbSize_Pw.SelectedItem = cmbSize_Pw.Items(0)
            'End If

            If cmbMaterialUsed_Pw.Items.Count > 0 Then
                cmbMaterialUsed_Pw.SelectedItem = cmbMaterialUsed_Pw.Items(0)
            End If

            If cmbMaterialSpec_Pw.Items.Count > 0 Then
                cmbMaterialSpec_Pw.SelectedItem = cmbMaterialSpec_Pw.Items(0)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CmbSize_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSize_Pw.SelectedIndexChanged
        If cmbCategory.Text = "Structure" Then
            FillGrade(dt, cmbPartType_Pw.Text, cmbSize_Pw.Text)
            FillPartWiseDetails(dt, cmbPartType_Pw.Text, cmbSize_Pw.Text)
        End If
    End Sub
    Private Sub FillGrade(ByVal dt As DataTable, ByVal partType As String, ByVal size As String)

        Try

            cmbGrade_Pw.Items.Clear()
            cmbGageName_Pw.Items.Clear()
            txtGageTable.Text = String.Empty
            cmbThickness_Pw.Items.Clear()
            cmbMaterialUsed_Pw.Items.Clear()
            cmbMaterialSpec_Pw.Items.Clear()

            Dim dv As New DataView(dt)
            Dim partTypeCol As String = ExcelUtil.ExcelSheetColumns.Type.ToString()
            Dim sizeCol As String = ExcelUtil.ExcelSheetColumns.Size.ToString()
            Dim gradeCol As String = ExcelUtil.ExcelSheetColumns.Grade.ToString()

            dv.RowFilter = $"{partTypeCol}='{cmbPartType_Pw.Text}' And {sizeCol}='{cmbSize_Pw.Text}'"
            Dim dtSize As DataTable = dv.ToTable()
            For Each drv As DataRowView In dv

                Dim grade1 As String = drv(ExcelUtil.ExcelSheetColumns.Grade.ToString)

                If Not cmbGrade_Pw.Items.Contains(grade1) Then
                    cmbGrade_Pw.Items.Add(grade1)
                End If

            Next

            If cmbGrade_Pw.Items.Count > 0 Then
                cmbGrade_Pw.SelectedItem = cmbGrade_Pw.Items(0)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CmbGrade_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbGrade_Pw.SelectedIndexChanged
        ' FillThickness_Gage(dt, cmbPartType_Pw.Text, cmbSize_Pw.Text, cmbGrade_Pw.Text)
        ' FillBendType2(dt, cmbPartType_Pw.Text, cmbSize_Pw.Text, cmbGrade_Pw.Text)

        If Not cmbPartType_Pw.Text = "NONE" Then
            If cmbCategory.Text = "SheetMetal" Then
                fill_SM_Pw_Details(dt)
                'FillGrade(dt, cmbPartType_Pw.Text, cmbSize_Pw.Text)
            End If
        End If

    End Sub
    Private Sub fill_SM_Pw_Details(ByVal dt As DataTable)
        cmbBendTypeGageWise_Pw.Items.Clear()
        cmbGageName_Pw.Items.Clear()
        cmbThickness_Pw.Items.Clear()
        cmbBendRadius_Pw.Items.Clear()
        cmbMaterialUsed_Pw.Items.Clear()
        cmbMaterialSpec_Pw.Items.Clear()
        txtBECMaterial_Pw.Clear()


        Dim dv As New DataView(dt)
        Dim partTypeCol As String = ExcelUtil.ExcelSheetColumns.Type.ToString()
        Dim sizeCol As String = ExcelUtil.ExcelSheetColumns.Size.ToString()
        Dim gradeCol As String = ExcelUtil.ExcelSheetColumns.Grade.ToString()
        Dim bendTypeCol As String = ExcelUtil.ExcelSheetColumns.Bend_Type.ToString()
        Dim gageNameCol As String = ExcelUtil.ExcelSheetColumns.Gage_Name.ToString()
        Dim thicknessCol As String = ExcelUtil.ExcelSheetColumns.Thickness.ToString()
        Dim bendRadiusCol As String = ExcelUtil.ExcelSheetColumns.Bend_Radius.ToString()
        Dim MaterialUsedCol As String = ExcelUtil.ExcelSheetColumns.Material_Used.ToString()
        Dim MaterialSpecCol As String = ExcelUtil.ExcelSheetColumns.Material_Specification.ToString()
        Dim BECMaterialCol As String = ExcelUtil.ExcelSheetColumns.BEC_Material.ToString()


        Dim ListOfGageName As New List(Of String)
        Dim ListOfBendRadius As New List(Of String)
        Dim ListOfMaterialSpec As New List(Of String)

        'Temp18APR2023
        Dim filter As String = $"Convert([{partTypeCol}], 'System.String') = '{cmbPartType_Pw.Text}' And Convert([{sizeCol}], 'System.String') = '{cmbSize_Pw.Text}' And Convert([{gradeCol}], 'System.String') = '{cmbGrade_Pw.Text}'"
        dv.RowFilter = filter '$"{partTypeCol}='{cmbPartType_Pw.Text}' And {sizeCol}='{cmbSize_Pw.Text}' And {gradeCol}='{cmbGrade_Pw.Text}'"
        Dim dtSize As DataTable = dv.ToTable()

        For Each drv As DataRowView In dv
            Dim bendType1 As String = drv(ExcelUtil.ExcelSheetColumns.Bend_Type.ToString).ToString()
            Dim gageName1 As String = drv(ExcelUtil.ExcelSheetColumns.Gage_Name.ToString).ToString()
            Dim thickness1 As String = drv(ExcelUtil.ExcelSheetColumns.Thickness.ToString).ToString()
            Dim bendRadius1 As String = drv(ExcelUtil.ExcelSheetColumns.Bend_Radius.ToString).ToString()
            Dim MaterialUsed1 As String = drv(ExcelUtil.ExcelSheetColumns.Material_Used.ToString).ToString()
            Dim MaterialSpec1 As String = drv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString).ToString()
            Dim BECMaterial1 As String = drv(ExcelUtil.ExcelSheetColumns.BEC_Material.ToString).ToString()


            If Not cmbBendTypeGageWise_Pw.Items.Contains(bendType1) Then
                cmbBendTypeGageWise_Pw.Items.Add(bendType1)
            End If

            If Not cmbGageName_Pw.Items.Contains(gageName1) Then
                ListOfGageName.Add(gageName1)
                cmbGageName_Pw.Items.Add(gageName1)
            End If

            If Not cmbThickness_Pw.Items.Contains(thickness1) Then
                cmbThickness_Pw.Items.Add(thickness1)
            End If

            If Not cmbBendRadius_Pw.Items.Contains(bendRadius1) Then
                ListOfBendRadius.Add(bendRadius1)
                cmbBendRadius_Pw.Items.Add(bendRadius1)
            End If

            If Not cmbMaterialUsed_Pw.Items.Contains(MaterialUsed1) Then
                cmbMaterialUsed_Pw.Items.Add(MaterialUsed1)
                cmbMaterialUsed_Pw.SelectedItem = cmbMaterialUsed_Pw.Items(0)
            End If

            If Not cmbMaterialSpec_Pw.Items.Contains(MaterialSpec1) Then
                ListOfMaterialSpec.Add(MaterialSpec1)
                cmbMaterialSpec_Pw.Items.Add(MaterialSpec1)
                cmbMaterialSpec_Pw.SelectedItem = cmbPartType_Pw.Items(0)
            End If
            txtBECMaterial_Pw.Text = BECMaterial1

        Next

        For i = 0 To ListOfGageName.Count - 1
            If Not cmbGageName_Pw.Items.Contains(ListOfGageName(i)) Then
                cmbGageName_Pw.Items.Add(ListOfGageName(i))
            End If
        Next


        For i = 0 To ListOfBendRadius.Count - 1
            If Not cmbBendRadius_Pw.Items.Contains(ListOfBendRadius(i)) Then
                cmbBendRadius_Pw.Items.Add(ListOfBendRadius(i))
            End If
        Next

        For i = 0 To ListOfMaterialSpec.Count - 1
            If Not cmbMaterialSpec_Pw.Items.Contains(ListOfMaterialSpec(i)) Then
                cmbMaterialSpec_Pw.Items.Add(ListOfMaterialSpec(i))

            End If
            cmbMaterialSpec_Pw.SelectedItem = cmbMaterialSpec_Pw.Items(0)
        Next
        If cmbBendTypeGageWise_Pw.Items.Count > 0 Then
            cmbBendTypeGageWise_Pw.SelectedItem = cmbBendTypeGageWise_Pw.Items(0)
        End If

        If cmbThickness_Pw.Items.Count > 0 Then
            cmbThickness_Pw.SelectedItem = cmbThickness_Pw.Items(0)
        End If

        If cmbBendRadius_Pw.Items.Count > 0 Then
            cmbBendRadius_Pw.SelectedItem = cmbBendRadius_Pw.Items(0)
        End If

        For i = 0 To ListOfGageName.Count - 1
            If Not cmbGageName_Pw.Items.Contains(ListOfGageName(i)) Then
                cmbGageName_Pw.Items.Add(ListOfGageName(i))
            End If
            cmbGageName_Pw.SelectedItem = cmbGageName_Pw.Items(0)
        Next
        'cmbGageName_Pw.Text = ListOfGageName(0)
        'cmbThickness_Pw.Text = cmbThickness_Pw.Items(0)
        'cmbBendRadius_Pw.Text = cmbBendRadius_Pw.Items(0)
    End Sub
    Private Sub FillBendType2(ByVal dt As DataTable, ByVal partType As String, ByVal size As String, ByVal grade As String)

        Try

            cmbBendTypeGageWise_Pw.Items.Clear()
            cmbBendTypeGageWise_Pw.Items.Add("NONE")
            'cmbGageName_Pw.Items.Clear()
            'cmbThickness_Pw.Items.Clear()
            'cmbBendRadius_Pw.Items.Clear()
            'cmbMaterialUsed_Pw.Items.Clear()
            'cmbMaterialSpec_Pw.Items.Clear()
            'txtGageTable.Text = String.Empty

            Dim dv As New DataView(dt)
            Dim partTypeCol As String = ExcelUtil.ExcelSheetColumns.Type.ToString()
            Dim sizeCol As String = ExcelUtil.ExcelSheetColumns.Size.ToString()
            Dim gradeCol As String = ExcelUtil.ExcelSheetColumns.Grade.ToString()
            Dim thicknessCol As String = ExcelUtil.ExcelSheetColumns.Thickness.ToString()
            Dim gageCol As String = ExcelUtil.ExcelSheetColumns.Gage_Name.ToString()
            Dim bendRadiusCol As String = ExcelUtil.ExcelSheetColumns.Bend_Radius.ToString()
            Dim gageTableCol As String = ExcelUtil.ExcelSheetColumns.Gage_Table.ToString()
            Dim bendTypeCol As String = ExcelUtil.ExcelSheetColumns.Bend_Type.ToString()

            dv.RowFilter = $"{partTypeCol}='{cmbPartType_Pw.Text}' And {sizeCol}='{cmbSize_Pw.Text}' And {gradeCol}='{cmbGrade_Pw.Text}'"
            Dim dtSize As DataTable = dv.ToTable()
            For Each drv As DataRowView In dv

                Dim bendType As String = drv(ExcelUtil.ExcelSheetColumns.Bend_Type.ToString)

                If Not cmbBendTypeGageWise_Pw.Items.Contains(bendType) Then
                    cmbBendTypeGageWise_Pw.Items.Add(bendType)
                End If
            Next

            If cmbBendTypeGageWise_Pw.Items.Count > 0 Then
                cmbBendTypeGageWise_Pw.SelectedItem = cmbBendTypeGageWise_Pw.Items(0)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub FillThickness_Gage(ByVal dt As DataTable, ByVal partType As String, ByVal size As String, ByVal grade As String)

        Try
            cmbThickness_Pw.Items.Clear()
            cmbMaterialUsed_Pw.Items.Clear()
            cmbMaterialSpec_Pw.Items.Clear()
            cmbGageName_Pw.Items.Clear()
            txtGageTable.Text = String.Empty

            Dim dv As New DataView(dt)
            Dim partTypeCol As String = ExcelUtil.ExcelSheetColumns.Type.ToString()
            Dim sizeCol As String = ExcelUtil.ExcelSheetColumns.Size.ToString()
            Dim gradeCol As String = ExcelUtil.ExcelSheetColumns.Grade.ToString()
            Dim thicknessCol As String = ExcelUtil.ExcelSheetColumns.Thickness.ToString()
            Dim gageCol As String = ExcelUtil.ExcelSheetColumns.Gage_Name.ToString()
            Dim bendRadiusCol As String = ExcelUtil.ExcelSheetColumns.Bend_Radius.ToString()
            Dim gageTableCol As String = ExcelUtil.ExcelSheetColumns.Gage_Table.ToString()

            dv.RowFilter = $"{partTypeCol}='{cmbPartType_Pw.Text}' And {sizeCol}='{cmbSize_Pw.Text}' And {gradeCol}='{cmbGrade_Pw.Text}'"
            Dim dtSize As DataTable = dv.ToTable()
            For Each drv As DataRowView In dv

                Dim thickness1 As String = drv(ExcelUtil.ExcelSheetColumns.Thickness.ToString)

                If Not cmbThickness_Pw.Items.Contains(thickness1) Then
                    cmbThickness_Pw.Items.Add(thickness1)
                End If

                Dim gageName1 As String = drv(ExcelUtil.ExcelSheetColumns.Gage_Name.ToString)

                If Not cmbGageName_Pw.Items.Contains(gageName1) Then
                    cmbGageName_Pw.Items.Add(gageName1)
                End If

                Dim bendRadius1 As String = drv(ExcelUtil.ExcelSheetColumns.Bend_Radius.ToString)

                If Not cmbBendRadius_Pw.Items.Contains(bendRadius1) Then
                    cmbBendRadius_Pw.Items.Add(bendRadius1)
                End If

                txtGageTable.Text = drv(ExcelUtil.ExcelSheetColumns.Gage_Table.ToString)

            Next

            If cmbThickness_Pw.Items.Count > 0 Then
                cmbThickness_Pw.SelectedItem = cmbThickness_Pw.Items(0)
            End If

            If cmbGageName_Pw.Items.Count > 0 Then
                cmbGageName_Pw.SelectedItem = cmbGageName_Pw.Items(0)
            End If

            If cmbBendRadius_Pw.Items.Count > 0 Then
                cmbBendRadius_Pw.SelectedItem = cmbBendRadius_Pw.Items(0)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CmbThickness_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbThickness_Pw.SelectedIndexChanged
        If Not cmbCategory.Text = "SheetMetal" And rbPartTypewise.Checked = True Then
            FillMaterialDetails(dt, cmbPartType_Pw.Text, cmbSize_Pw.Text, cmbGrade_Pw.Text, cmbThickness_Pw.Text)
        End If

    End Sub

    Private Sub FillMaterialDetails(ByVal dt As DataTable, ByVal partType As String, ByVal size As String, ByVal grade As String, ByVal thickness As String)

        Try
            cmbBendRadius_Pw.Items.Clear()
            cmbMaterialUsed_Pw.Items.Clear()
            cmbMaterialSpec_Pw.Items.Clear()
            txtBECMaterial_Pw.Text = String.Empty

            Dim dv As New DataView(dt)
            Dim partTypeCol As String = ExcelUtil.ExcelSheetColumns.Type.ToString()
            Dim sizeCol As String = ExcelUtil.ExcelSheetColumns.Size.ToString()
            Dim gradeCol As String = ExcelUtil.ExcelSheetColumns.Grade.ToString()
            Dim thicknessCol As String = ExcelUtil.ExcelSheetColumns.Thickness.ToString()

            Dim bendRadiusCol As String = ExcelUtil.ExcelSheetColumns.Bend_Radius.ToString()
            Dim materialUsedCol As String = ExcelUtil.ExcelSheetColumns.Material_Used.ToString()
            Dim materialSpecCol As String = ExcelUtil.ExcelSheetColumns.Material_Specification.ToString()
            Dim becMaterialCol As String = ExcelUtil.ExcelSheetColumns.BEC_Material.ToString()

            Dim bendTypeCol As String = ExcelUtil.ExcelSheetColumns.Bend_Type.ToString()
            Dim gageNameCol As String = ExcelUtil.ExcelSheetColumns.Gage_Name.ToString()

            dv.RowFilter = $"{partTypeCol}='{cmbPartType_Pw.Text}' And {sizeCol}='{cmbSize_Pw.Text}' And {gradeCol}='{cmbGrade_Pw.Text}' And {thicknessCol}='{cmbThickness_Pw.Text}' And {bendTypeCol}='{cmbBendTypeGageWise_Pw.Text}' And {gageNameCol}='{cmbGageName_Pw.Text}'"
            Dim dtSize As DataTable = dv.ToTable()
            For Each drv As DataRowView In dv

                Dim materialUsed As String = drv(ExcelUtil.ExcelSheetColumns.Material_Used.ToString)

                If Not cmbMaterialUsed_Pw.Items.Contains(materialUsed) Then
                    cmbMaterialUsed_Pw.Items.Add(materialUsed)
                End If

                Dim materialSpec As String = drv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString)

                If Not cmbMaterialSpec_Pw.Items.Contains(materialSpec) Then
                    cmbMaterialSpec_Pw.Items.Add(materialSpec)
                End If

                Dim bendRadius As String = drv(ExcelUtil.ExcelSheetColumns.Bend_Radius.ToString)

                If Not cmbBendRadius_Pw.Items.Contains(bendRadius) Then
                    cmbBendRadius_Pw.Items.Add(bendRadius)
                End If

                Dim becMaterial As String = drv(ExcelUtil.ExcelSheetColumns.BEC_Material.ToString)
                txtBECMaterial_Pw.Text = becMaterial

            Next

            If cmbMaterialUsed_Pw.Items.Count > 0 Then
                cmbMaterialUsed_Pw.SelectedItem = cmbMaterialUsed_Pw.Items(0)
            End If

            If cmbMaterialSpec_Pw.Items.Count > 0 Then
                cmbMaterialSpec_Pw.SelectedItem = cmbMaterialSpec_Pw.Items(0)
            End If

            If cmbBendRadius_Pw.Items.Count > 0 Then
                cmbBendRadius_Pw.SelectedItem = cmbBendRadius_Pw.Items(0)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CmbMaterialUseds_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbMaterialUsed_Pw.SelectedIndexChanged
        If Not cmbCategory.Text = "SheetMetal" And rbPartTypewise.Checked = True Then
            FillMaterialSpecDetails(dt, cmbPartType_Pw.Text, cmbPartType_Pw.Text)
        End If

        If rbPartTypewise.Checked Then

            txtImageName.Text = String.Empty
            txtImageName.Text = cmbMaterialSpec_Pw.Text

        End If

    End Sub



    Private Sub BtnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click

        SetCustProp()

        ApplyMaterial()

        ApplyCustomProperties()

        Try
            If isSheetMetalPart Then

                'objMatTable.ApplyGageFromLibraryToDoc(objSheetMetalDocument, txtGageName_Mw.Text, cmbMaterialLib.Text)

                ''objMatTable.ApplyGageFromGageTableToDoc(objSheetMetalDocument, "20 Gage Al", "aluminum 6061")


                objMatTable.SetDocumentToGageTableAssociation(objSheetMetalDocument, custPropertiesObj.gageName, txtGageTable.Text, True, True)
            Else
                'objMatTable.ApplyGageFromGageTableToDoc(objDocument, txtGageName_Mw.Text, "Gagetable")

                ' objMatTable.SetActiveDocument(objDocument)

                'objMatTable.ApplyGageFromLibraryToDoc(objDocument, custPropertiesObj.gageName, cmbMaterialLib.Text)

                objMatTable.SetDocumentToGageTableAssociation(objDocument, custPropertiesObj.gageName, txtGageTable.Text, True, True)

                ' ApplyMaterial()

                'objMatTable.UpdateOODMaterialAndGageProperties(objDocument, True, True)

            End If

            ' ApplyMaterial()

            MessageBox.Show("Process completed.", "Message")
            CustomLogUtil.Heading("Part/ SheetMetal Process successfully Completed")
        Catch ex As Exception

            MessageBox.Show($"Error in apply {vbNewLine}{ex.Message}{ ex.StackTrace}", "Message")
            CustomLogUtil.Log("While apply", ex.Message, ex.StackTrace)
        End Try

        PerformRefreshAction()
        'objMatTable.ApplyGageFromGageTableToDoc(objSheetMetalDocument, "11 GAGE", "BEC MATERIAL LIBRARY")
        'objMatTable.ApplyGageFromLibraryToDoc(objSheetMetalDocument, "PLATE - .032", "BEC MATERIAL LIBRARY") ' "Materials-GOST")
        ' objMatTable.ApplyGageFromGageTableToDoc(objSheetMetalDocument, "SHEET - 10 GA", "Materials-GOST")
    End Sub

    Private Function GetGageName() As String

        Dim gageNameOriginal As String = String.Empty
        If rbPartTypewise.Checked Then

            Try
                'cmbThickness_Pw.Items.Clear()

                Dim dv As New DataView(dt)
                Dim partTypeCol As String = ExcelUtil.ExcelSheetColumns.Type.ToString()
                Dim sizeCol As String = ExcelUtil.ExcelSheetColumns.Size.ToString()
                Dim gradeCol As String = ExcelUtil.ExcelSheetColumns.Grade.ToString()
                Dim BendTypeCol As String = ExcelUtil.ExcelSheetColumns.Bend_Type.ToString()
                Dim gageCol As String = ExcelUtil.ExcelSheetColumns.Gage_Name.ToString()
                Dim thicknessCol As String = ExcelUtil.ExcelSheetColumns.Thickness.ToString()

                dv.RowFilter = $"{partTypeCol}='{cmbPartType_Pw.Text}' And {sizeCol}='{cmbSize_Pw.Text}' And {gradeCol}='{cmbGrade_Pw.Text}' And {BendTypeCol}='{cmbBendTypeGageWise_Pw.Text}' And {gageCol}='{cmbGageName_Pw.Text}'"
                Dim dtSize As DataTable = dv.ToTable()
                For Each drv As DataRowView In dv
                    gageNameOriginal = drv(ExcelUtil.ExcelSheetColumns.Gage_Name.ToString)
                Next
            Catch ex As Exception

            End Try
        Else
            Try
                'cmbThickness_Pw.Items.Clear()

                Dim dv As New DataView(dt)
                Dim materialUsedCol As String = ExcelUtil.ExcelSheetColumns.Material_Used.ToString()
                Dim sizeCol As String = ExcelUtil.ExcelSheetColumns.Size.ToString()
                Dim gradeCol As String = ExcelUtil.ExcelSheetColumns.Grade.ToString()
                Dim BendTypeCol As String = ExcelUtil.ExcelSheetColumns.Bend_Type.ToString()
                Dim gageCol As String = ExcelUtil.ExcelSheetColumns.Gage_Name.ToString()
                Dim thicknessCol As String = ExcelUtil.ExcelSheetColumns.Thickness.ToString()


                'temp2Feb2024
                dv.RowFilter = $"{materialUsedCol}='{cmbMaterialUsed2_Mw.Text}' And {gradeCol}='{txtGrade2_Mw.Text}' And {BendTypeCol}='{cmbBendTypeGageWise_Mw.Text}' And {gageCol}='{cmbGageName.Text}'"

                Dim dtSize As DataTable = dv.ToTable()

                For Each drv As DataRowView In dv

                    gageNameOriginal = drv(ExcelUtil.ExcelSheetColumns.Gage_Name.ToString)

                Next
            Catch ex As Exception

            End Try
        End If
        Return gageNameOriginal
    End Function

    Private Sub SetCustProp()

        If rbPartTypewise.Checked Then
            custPropertiesObj.partType = cmbPartType_Pw.Text
            custPropertiesObj.size = cmbSize_Pw.Text
            custPropertiesObj.grade = cmbGrade_Pw.Text
            custPropertiesObj.materialThickness = cmbThickness_Pw.Text
            custPropertiesObj.materialUsed = cmbMaterialUsed_Pw.Text
            custPropertiesObj.materialSpec = cmbMaterialSpec_Pw.Text

            custPropertiesObj.gageName = GetGageName() ' cmbGageName_Pw.Text
        Else
            custPropertiesObj.partType = txtPartType_Mw.Text
            custPropertiesObj.size = txtSize2_Mw.Text
            custPropertiesObj.grade = txtGrade2_Mw.Text
            custPropertiesObj.materialThickness = txtThickness2_Mw.Text
            custPropertiesObj.materialUsed = cmbMaterialUsed2_Mw.Text
            custPropertiesObj.materialSpec = txtMaterialSpec2_Mw.Text

            custPropertiesObj.gageName = GetGageName() ' txtGageName_Mw.Text
        End If

    End Sub

    Private Sub ApplyMaterial()
        Try

            Dim materialName As String = String.Empty

            If rbPartTypewise.Checked Then
                materialName = txtBECMaterial_Pw.Text
            Else
                materialName = txtBECMaterial2_Mw.Text
            End If

            Dim lstString As List(Of String) = dicMaterials(cmbMaterialLib.Text)

            If lstString.Contains(materialName) Then

                ' Set active document handle

                If Not isSheetMetalPart Then
                    objDocument = objApp.ActiveDocument
                    objMatTable.SetActiveDocument(objDocument)

                    objMatTable.ApplyMaterialToDoc(objDocument, materialName, cmbMaterialLib.Text)
                Else

                    objSheetMetalDocument = objApp.ActiveDocument
                    objMatTable.SetActiveDocument(objSheetMetalDocument)

                    objMatTable.ApplyMaterialToDoc(objSheetMetalDocument, materialName, cmbMaterialLib.Text)

                End If
            Else
                MessageBox.Show($"Material does not exist in {cmbMaterialLib.Text}", "Message")
            End If
        Catch ex As Exception
            CustomLogUtil.Log("on applying Material", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub ApplyCustomProperties()
        Try

            If objDocument IsNot Nothing Then

                Dim propSets As SolidEdgeFramework.PropertySets = objDocument.Properties

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
            CustomLogUtil.Log("While applying Custom Properties ", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub RbPartTypewise_CheckedChanged(sender As Object, e As EventArgs) Handles rbPartTypewise.CheckedChanged
        ControlsEnability(True)
        Validation()
        'SheetMetalPartwiseValidation()
    End Sub
    Public Sub SheetMetalPartwiseValidation()
        If cmbCategory.Text = "SheetMetal" And rbPartTypewise.Checked = True Then
            cmbSize_Pw.Enabled = False
            cmbBendRadius_Pw.Enabled = False
            cmbThickness_Pw.Enabled = False
            cmbBendRadius_Pw.Enabled = False
            cmbMaterialUsed_Pw.Enabled = False
            cmbMaterialSpec_Pw.Enabled = False
            txtBECMaterial_Pw.Enabled = False
            cmbBendTypeGageWise_Pw.Enabled = False
        Else
            cmbSize_Pw.Enabled = True
            cmbBendRadius_Pw.Enabled = True
            cmbThickness_Pw.Enabled = True
            cmbBendRadius_Pw.Enabled = True
            cmbMaterialUsed_Pw.Enabled = True
            cmbMaterialSpec_Pw.Enabled = True
            txtBECMaterial_Pw.Enabled = True
            cmbBendTypeGageWise_Pw.Enabled = True
        End If

    End Sub
    Private Sub RbMaterialWise_CheckedChanged(sender As Object, e As EventArgs) Handles rbMaterialWise.CheckedChanged

        ControlsEnability(False)
        Validation()
    End Sub

    Private Sub SetTextBoxTheme(ByRef txt As System.Windows.Forms.TextBox, ByVal err As Boolean)

        If err Then
            txt.BackColor = Color.FromArgb(255, 199, 206)
            txt.ForeColor = Color.FromArgb(156, 0, 6)
        Else
            txt.BackColor = Color.White
            txt.ForeColor = Color.Black
        End If

    End Sub

    Private Sub Validation()

        If Not isBrowse Then
            Exit Sub
        End If

        SetTextBoxTheme(txtPartType_C, False)
        SetTextBoxTheme(txtSize_C, False)
        SetTextBoxTheme(txtGrade_C, False)
        SetTextBoxTheme(txtThickness_C, False)
        SetTextBoxTheme(txtMaterialUsed_C, False)
        SetTextBoxTheme(txtMaterialSpec_C, False)
        SetTextBoxTheme(txtBECMaterial_Pw, False)

        Try
            If rbMaterialWise.Checked Then

                If cmbMaterialUsed2_Mw.Text = txtMaterialUsed_C.Text Then
                    txtMaterialUsed_C.BackColor = Color.White
                    txtMaterialUsed_C.ForeColor = Color.Black
                Else
                    txtMaterialUsed_C.BackColor = Color.FromArgb(255, 199, 206)
                    txtMaterialUsed_C.ForeColor = Color.FromArgb(156, 0, 6)
                End If
                Dim thicknessValC = txtThickness_C.Text
                Dim thicknessValMw = txtThickness2_Mw.Text

                If thicknessValC = thicknessValMw Then
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

                'Temp18APR2023
                If txtBendRadius_Mw.Text = txtBendRadius_C.Text Then
                    txtBendRadius_C.BackColor = Color.White
                    txtBendRadius_C.ForeColor = Color.Black
                Else
                    txtBendRadius_C.BackColor = Color.FromArgb(255, 199, 206)
                    txtBendRadius_C.ForeColor = Color.FromArgb(156, 0, 6)
                End If
                If txtMaterialSpec_C.Text = txtMaterialSpec2_Mw.Text Then
                    txtMaterialSpec_C.BackColor = Color.White
                    txtMaterialSpec_C.ForeColor = Color.Black
                Else
                    txtMaterialSpec_C.BackColor = Color.FromArgb(255, 199, 206)
                    txtMaterialSpec_C.ForeColor = Color.FromArgb(156, 0, 6)
                End If
            Else
                cmbSize_Pw.Enabled = True
                If cmbMaterialUsed2_Mw.Text = txtMaterialUsed_C.Text Then
                    txtMaterialUsed_C.BackColor = Color.White
                    txtMaterialUsed_C.ForeColor = Color.Black
                Else
                    txtMaterialUsed_C.BackColor = Color.FromArgb(255, 199, 206)
                    txtMaterialUsed_C.ForeColor = Color.FromArgb(156, 0, 6)
                End If

                If cmbThickness_Pw.Text = txtThickness_C.Text Then
                    txtThickness_C.BackColor = Color.White
                    txtThickness_C.ForeColor = Color.Black
                Else
                    txtThickness_C.BackColor = Color.FromArgb(255, 199, 206)
                    txtThickness_C.ForeColor = Color.FromArgb(156, 0, 6)
                End If

                If txtBECMaterial_Pw.Text = txtCurrentMaterial_C.Text Then
                    txtCurrentMaterial_C.BackColor = Color.White
                    txtCurrentMaterial_C.ForeColor = Color.Black
                Else
                    txtCurrentMaterial_C.BackColor = Color.FromArgb(255, 199, 206)
                    txtCurrentMaterial_C.ForeColor = Color.FromArgb(156, 0, 6)
                End If

            End If
        Catch ex As Exception
            MessageBox.Show($"Error in validation {vbNewLine}{ex.Message}{ ex.StackTrace}", "Message")
            CustomLogUtil.Log("in validation ", ex.Message, ex.StackTrace)
        End Try
        BendRadiusColorValidate()
    End Sub

    Public Sub BendRadiusColorValidate()  'TEMP12SEPT2023
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
    Private Sub ControlsEnability(ByVal flag As Boolean)

        cmbPartType_Pw.Visible = flag
        cmbSize_Pw.Visible = flag
        cmbGrade_Pw.Visible = flag
        cmbGageName_Pw.Visible = flag
        cmbThickness_Pw.Visible = flag
        cmbBendRadius_Pw.Visible = flag
        cmbMaterialUsed_Pw.Visible = flag
        cmbMaterialSpec_Pw.Visible = flag
        txtBECMaterial_Pw.Visible = flag
        cmbBendTypeGageWise_Pw.Visible = flag

        cmbMaterialUsed2_Mw.Visible = Not flag
        txtSize2_Mw.Visible = Not flag
        txtGrade2_Mw.Visible = Not flag
        'temp2Feb2024
        'txtGageName_Mw.Visible = Not flag
        cmbGageName.Visible = Not flag
        txtThickness2_Mw.Visible = Not flag
        txtBendRadius_Mw.Visible = Not flag
        txtPartType_Mw.Visible = Not flag
        txtMaterialSpec2_Mw.Visible = Not flag
        txtBECMaterial2_Mw.Visible = Not flag
        cmbBendTypeGageWise_Mw.Visible = Not flag


        'If flag Then
        '    lbl1.Text = "Part Type"
        '    'lbl2.Text = "Size"
        '    'lbl3.Text = "Grade"
        '    'lbl5.Text = "Material Thickness (inch)"
        '    lbl7.Text = "Material Used"
        '    'lbl8.Text = "Material Spec"
        '    'lbl9.Text = "BEC Material"
        'Else

        '    lbl1.Text = "Material Used"
        '    'lbl2.Text = "Size"
        '    'lbl3.Text = "Grade"
        '    'lbl5.Text = "Material Thickness (inch)"
        '    lbl7.Text = "Part Type"
        '    'lbl8.Text = "Material Spec"
        '    'lbl9.Text = "BEC Material"

        'End If

    End Sub

    Private Sub CmbMaterialUsed2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbMaterialUsed2_Mw.SelectedIndexChanged

        'temp2Feb2024
        'FillMaterialUsedWiseDetails(dt, cmbMaterialUsed2_Mw.Text)
        'FillBendType__AccordingToPriority()

        If cmbCategory.Text = "Structure" Then
            FillMaterialUsedWiseDetails(dt, cmbMaterialUsed2_Mw.Text)
            FillBendType__AccordingToPriority()
        ElseIf cmbCategory.Text = "SheetMetal" Then
            FillMaterialUsedWiseDetails(dt, cmbMaterialUsed2_Mw.Text)
            FillBendType__AccordingToPriority()
            'FillGageName1()
        End If

    End Sub

    Public Sub ResetMWTxt()
        txtSize2_Mw.Text = String.Empty
        txtGrade2_Mw.Text = String.Empty
        txtThickness2_Mw.Text = String.Empty
        txtPartType_Mw.Text = String.Empty
        txtMaterialSpec2_Mw.Text = String.Empty
        txtBECMaterial2_Mw.Text = String.Empty
        'temp2Feb2024
        'txtGageName_Mw.Text = String.Empty
        cmbGageName.Text = String.Empty
        txtGageTable.Text = String.Empty
        txtBendRadius_Mw.Text = String.Empty
        txtImageName.Text = String.Empty
        cmbBendTypeGageWise_Mw.Text = String.Empty
        cmbBendTypeGageWise_Mw.Items.Clear()
    End Sub

    Private Sub FillMaterialUsedWiseDetails(ByVal dt As DataTable, ByVal materialUsed As String)

        Try

            txtSize2_Mw.Text = String.Empty
            txtGrade2_Mw.Text = String.Empty
            txtThickness2_Mw.Text = String.Empty
            txtPartType_Mw.Text = String.Empty
            txtMaterialSpec2_Mw.Text = String.Empty
            txtBECMaterial2_Mw.Text = String.Empty
            'temp2Feb2024
            'txtGageName_Mw.Text = String.Empty
            cmbGageName.Text = String.Empty
            txtGageTable.Text = String.Empty
            txtBendRadius_Mw.Text = String.Empty
            txtImageName.Text = String.Empty
            cmbBendTypeGageWise_Mw.Text = String.Empty
            cmbBendTypeGageWise_Mw.Items.Clear()

            Dim dv As New DataView(dt)

            Dim materialUsedCol As String = ExcelUtil.ExcelSheetColumns.Material_Used.ToString()
            'Temp18APR2023
            Dim filter As String = $"Convert([{materialUsedCol}], 'System.String') = '{cmbMaterialUsed2_Mw.Text}'"
            dv.RowFilter = filter '$"{materialUsedCol}='{cmbMaterialUsed2_Mw.Text}'"

            For Each drv As DataRowView In dv

                txtSize2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Size.ToString).ToString()
                txtGrade2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Grade.ToString).ToString()
                If cmbCategory.Text = "SheetMetal" Then
                    'temp2Feb2024
                    'txtGageName_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Gage_Name.ToString).ToString()
                    cmbGageName.Text = drv(ExcelUtil.ExcelSheetColumns.Gage_Name.ToString).ToString()
                    txtGageTable.Text = drv(ExcelUtil.ExcelSheetColumns.Gage_Table.ToString).ToString()
                    txtBendRadius_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Bend_Radius.ToString).ToString()
                End If


                txtThickness2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Thickness.ToString).ToString()

                txtPartType_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Type.ToString).ToString()
                txtMaterialSpec2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString).ToString()
                txtBECMaterial2_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.BEC_Material.ToString).ToString()
                'txtImageName.Text = drv(ExcelUtil.ExcelSheetColumns.Image.ToString).ToString()
                Exit For

            Next
            If cmbCategory.Text = "SheetMetal" Then
                For Each drv As DataRowView In dv

                    Dim bendType As String = drv(ExcelUtil.ExcelSheetColumns.Bend_Type.ToString)
                    If Not cmbBendTypeGageWise_Mw.Items.Contains(bendType) Then
                        cmbBendTypeGageWise_Mw.Items.Add(bendType)
                    End If
                Next

                If cmbBendTypeGageWise_Mw.Items.Count > 0 Then
                    cmbBendTypeGageWise_Mw.SelectedItem = cmbBendTypeGageWise_Mw.Items(0)
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub TxtPartType_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPartType_C.KeyPress
        e.Handled = True
    End Sub

    Private Sub TxtSize_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtSize_C.KeyPress
        e.Handled = True
    End Sub

    Private Sub TxtGrade_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtGrade_C.KeyPress
        e.Handled = True
    End Sub

    Private Sub TxtThickness_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtThickness_C.KeyPress
        e.Handled = True
    End Sub

    Private Sub TxtMaterialUsed_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtMaterialUsed_C.KeyPress
        e.Handled = True
    End Sub

    Private Sub TxtMaterialSpec_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtMaterialSpec_C.KeyPress
        e.Handled = True
    End Sub

    Private Sub TxtCurrentMaterial_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCurrentMaterial_C.KeyPress
        e.Handled = True
    End Sub

    Private Sub TxtSize2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtSize2_Mw.KeyPress
        e.Handled = True
    End Sub

    Private Sub TxtGrade2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtGrade2_Mw.KeyPress
        e.Handled = True
    End Sub

    Private Sub TxtThickness2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtThickness2_Mw.KeyPress
        e.Handled = True
    End Sub

    Private Sub TxtPartType2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPartType_Mw.KeyPress
        e.Handled = True
    End Sub

    Private Sub TxtMaterialSpec2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtMaterialSpec2_Mw.KeyPress
        e.Handled = True
    End Sub

    Private Sub TxtBECMaterial2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtBECMaterial2_Mw.KeyPress
        e.Handled = True
    End Sub

    Private Sub BtnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        PerformRefreshAction()
    End Sub

    Public Sub PerformRefreshAction()
        ResetMWTxt()
        'TEMP27APR2023
        GetFormLoadData()
        GetData()

        SetAutoSuggestForGageName()
        ''OldCode
        'RefreshForm()
        'cmbCategory.SelectedItem = cmbCategory.Text
        'Validation()
    End Sub

    Private Sub TxtBECMaterial_Pw_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtBECMaterial_Pw.KeyPress
        e.Handled = True
    End Sub

    Private Sub TxtGageTable_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtGageTable.KeyPress
        e.Handled = True
    End Sub

    Private Sub TxtGageName_Mw_KeyPress(sender As Object, e As KeyPressEventArgs)
        e.Handled = True
    End Sub

    Private Sub CmbCategory_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCategory.SelectedIndexChanged
        Try


            'Dim ds As DataSet = dicData("Structure")
            'dtstructure = ds.Tables(0)
            'Dim ds2 As DataSet = dicData("SheetMetal")
            'dtsheetmetal = ds2.Tables(0)

            CmbValidations()

            If cmbCategory.Text = "Structure" Then
                dt = dtstructure.Copy()

            ElseIf cmbCategory.Text = "SheetMetal" Then
                dt = dtsheetmetal.Copy()

            Else
                Exit Sub
            End If

            FillPartType(dt)

            FillMaterialUsed2(dt)


            SetPartTypeWiseDetails(dt)

            'TEMP25APR2023
            SetAutoSuggestForGageName()




            rbPartTypewise.Enabled = True
            rbMaterialWise.Enabled = True
            btnApply.Enabled = True

            rbMaterialWise.Checked = True
            isBrowse = True

            Validation()
            If Not cmbMaterialUsed2_Mw.Items.Contains(cmbMaterialUsed2_Mw.Text) Then
                Dim materialused = cmbMaterialUsed2_Mw.Items(0)
                cmbMaterialUsed2_Mw.Text = materialused
                FillMaterialUsedWiseDetails(dt, materialused)
            End If



        Catch ex As Exception

        End Try
    End Sub



    Public Sub SetAutoSuggestForGageName()




        'If cmbMaterialUsed2_Mw.Items.Contains(txtMaterialUsed_C.Text) Then

        '    cmbMaterialUsed2_Mw.Text = txtMaterialUsed_C.Text
        '    'txtGageName_C.Text += " (Trumpf Air punch)"
        '    For i = 0 To cmbBendTypeGageWise_Mw.Items.Count - 1
        '        If txtGageName_C.Text.Contains(cmbBendTypeGageWise_Mw.Items(i)) Then
        '            cmbBendTypeGageWise_Mw.SelectedItem = cmbBendTypeGageWise_Mw.Items(i)
        '            cmbBendTypeGageWise_Mw.Text = cmbBendTypeGageWise_Mw.Items(i)
        '            Exit For
        '        End If
        '    Next
        'End If
    End Sub
    Public Sub CmbValidations()
        If cmbCategory.SelectedItem = "Structure" Then
            cmbSize_Pw.Enabled = False
            cmbBendTypeGageWise_Pw.Enabled = False
            cmbBendRadius_Pw.Enabled = False
            cmbGageName_Pw.Enabled = False
            cmbThickness_Pw.Enabled = False


            txtSize2_Mw.Enabled = False
            cmbBendTypeGageWise_Mw.Enabled = False
            txtBendRadius_Mw.Enabled = False
            'temp2Feb2024
            'txtGageName_Mw.Enabled = False
            cmbGageName.Enabled = False
            txtThickness2_Mw.Enabled = False
        ElseIf cmbCategory.SelectedItem = "SheetMetal" Then
            cmbSize_Pw.Enabled = True
            cmbBendTypeGageWise_Pw.Enabled = True
            cmbBendRadius_Pw.Enabled = True
            cmbGageName_Pw.Enabled = True
            cmbThickness_Pw.Enabled = True

            txtSize2_Mw.Enabled = True
            cmbBendTypeGageWise_Mw.Enabled = True
            txtBendRadius_Mw.Enabled = True
            'temp2Feb2024
            'txtGageName_Mw.Enabled = True

            '16th october 2024
            'cmbGageName.Enabled = True

            txtThickness2_Mw.Enabled = True
        End If
    End Sub

    Private Sub PartAutomationForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'System.Windows.Forms.Application.Exit()
    End Sub

    Private Sub BtnShowGuideLines_Click(sender As Object, e As EventArgs) Handles btnShowGuideLines.Click

        If rbMaterialWise.Checked Then

            Dim imageName As String = $"{txtImageName.Text}.png"
            Dim frm As New InstructionForm(imageName)
            frm.Show()
        Else
            Dim imageName As String = $"{txtImageName.Text}.png"
            Dim frm As New InstructionForm(imageName)
            frm.Show()

        End If
    End Sub

    Private Sub CmbGageName_Pw_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbGageName_Pw.SelectedIndexChanged
        If cmbCategory.Text = "SheetMetal" And rbPartTypewise.Checked = True Then
            Dim partTypeCol As String = ExcelUtil.ExcelSheetColumns.Type.ToString()
            Dim gradeCol As String = ExcelUtil.ExcelSheetColumns.Grade.ToString()
            Dim thicknessCol As String = ExcelUtil.ExcelSheetColumns.Thickness.ToString()
            Dim MaterialUsedCol As String = ExcelUtil.ExcelSheetColumns.Material_Used.ToString()
            Dim MaterialSpecCol As String = ExcelUtil.ExcelSheetColumns.Material_Specification.ToString()
            Dim BendTypeCol As String = ExcelUtil.ExcelSheetColumns.Bend_Type.ToString()
            Dim gageCol As String = ExcelUtil.ExcelSheetColumns.Gage_Name.ToString()


            Dim dv As New DataView(dt)
            dv.RowFilter = $"{partTypeCol}='{cmbPartType_Pw.Text}' And {gageCol}='{cmbGageName_Pw.Text}' And {gradeCol}='{cmbGrade_Pw.Text}'"

            For Each drv As DataRowView In dv
                Dim thickness1 As String = drv(ExcelUtil.ExcelSheetColumns.Thickness.ToString).ToString()
                Dim materialUsed1 As String = drv(ExcelUtil.ExcelSheetColumns.Material_Used.ToString).ToString()
                Dim BendRadius1 As String = drv(ExcelUtil.ExcelSheetColumns.Bend_Radius.ToString).ToString()
                Dim BendType1 As String = drv(ExcelUtil.ExcelSheetColumns.Bend_Type.ToString).ToString()
                Dim MaterialSpec1 As String = drv(ExcelUtil.ExcelSheetColumns.Material_Specification.ToString).ToString()

                cmbThickness_Pw.Items.Clear()
                cmbThickness_Pw.Items.Add(thickness1)
                cmbThickness_Pw.SelectedItem = cmbThickness_Pw.Items(0)

                cmbMaterialUsed_Pw.Items.Clear()
                cmbMaterialUsed_Pw.Items.Add(materialUsed1)
                cmbMaterialUsed_Pw.SelectedItem = cmbMaterialUsed_Pw.Items(0)

                cmbMaterialSpec_Pw.Items.Clear()
                cmbMaterialSpec_Pw.Items.Add(MaterialSpec1)
                cmbMaterialSpec_Pw.SelectedItem = cmbMaterialSpec_Pw.Items(0)

                cmbBendRadius_Pw.Items.Clear()
                cmbBendRadius_Pw.Items.Add(BendRadius1)
                cmbBendRadius_Pw.SelectedItem = cmbBendRadius_Pw.Items(0)


                cmbBendTypeGageWise_Pw.Items.Clear()
                cmbBendTypeGageWise_Pw.Items.Add(BendType1)
                cmbBendTypeGageWise_Pw.SelectedItem = cmbBendTypeGageWise_Pw.Items(0)
            Next

        End If
        'If rbPartTypewise.Checked Then

        '    Try
        '        cmbThickness_Pw.Items.Clear()

        '        Dim dv As New DataView(dt)
        '        Dim partTypeCol As String = ExcelUtil.ExcelSheetColumns.Type.ToString()
        '        Dim sizeCol As String = ExcelUtil.ExcelSheetColumns.Size.ToString()
        '        Dim gradeCol As String = ExcelUtil.ExcelSheetColumns.Grade.ToString()
        '        Dim BendTypeCol As String = ExcelUtil.ExcelSheetColumns.Bend_Type.ToString()
        '        Dim gageCol As String = ExcelUtil.ExcelSheetColumns.Gage_Name.ToString()
        '        Dim thicknessCol As String = ExcelUtil.ExcelSheetColumns.Thickness.ToString()

        '        dv.RowFilter = $"{partTypeCol}='{cmbPartType_Pw.Text}' And {sizeCol}='{cmbSize_Pw.Text}' And {gradeCol}='{cmbGrade_Pw.Text}' And {BendTypeCol}='{cmbBendTypeGageWise_Pw.Text}' And {gageCol}='{cmbGageName_Pw.Text}'"
        '        Dim dtSize As DataTable = dv.ToTable()
        '        For Each drv As DataRowView In dv
        '            Dim thickness As String = drv(ExcelUtil.ExcelSheetColumns.Thickness.ToString)

        '            If Not cmbThickness_Pw.Items.Contains(thickness) Then
        '                cmbThickness_Pw.Items.Add(thickness)
        '            End If
        '            'txtGageTable.Text = drv(ExcelUtil.excelSheetColumns.Gage_Table.ToString)

        '        Next
        '        If cmbThickness_Pw.Items.Count > 0 Then
        '            cmbThickness_Pw.SelectedItem = cmbThickness_Pw.Items(0)
        '        End If
        '    Catch ex As Exception

        '    End Try
        'End If

    End Sub

    Private Sub FillBendTypeGageNameWise(ByVal dt As DataTable) 'TEMP13SETP2023 
        'Priority code
        cmbBendTypeGageWise_Mw.Items.Clear()

        Dim dv As New DataView(dt)
        Dim gageNameCol As String = ExcelUtil.ExcelSheetColumns.Gage_Name.ToString()

        Dim Priorityspecol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Priority.ToString()
        Dim Priority As Integer = 0
        Dim value As Integer = 0

        'temp2Feb2024
        dv.RowFilter = $"{gageNameCol}='{cmbGageName.Text}'"

        'cmbBendTypeGageWise_Mw.Items.Clear() 'new
        For Each drv As DataRowView In dv

            Dim bendTypeGageWise As String = drv(ExcelUtil.ExcelSheetColumns.Bend_Type.ToString)

            'If Not cmbBendTypeGageWise_Mw.Items.Contains(bendTypeGageWise) Then
            '    If Priority = 0 Then
            '        Priority = Convert.ToInt32(drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Priority.ToString).ToString())
            '    End If
            '    value = Convert.ToInt32(drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Priority.ToString).ToString())


            '    If (Priority > value) Then
            '        Priority = value
            '    End If
            '    Dim BendTypeGageWiseName As String = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Bend_Type.ToString).ToString()
            '    If Not cmbBendTypeGageWise_Mw.Items.Contains(BendTypeGageWiseName) Then
            cmbBendTypeGageWise_Mw.Items.Add(bendTypeGageWise) 'onlyold
            '    End If


            'End If

            'old
            If cmbBendTypeGageWise_Pw.Items.Count > 0 Then
                cmbBendTypeGageWise_Pw.SelectedItem = cmbBendTypeGageWise_Pw.Items(0)
            End If

        Next
        'If dv.Count = 1 And cmbBendTypeGageWise_Mw.Items.Count > 0 Then
        '    cmbBendTypeGageWise_Mw.SelectedItem = cmbBendTypeGageWise_Mw.Items(0)
        'Else
        '    dv.RowFilter = $"{gageNameCol}='{txtGageName_Mw.Text}' and {Priorityspecol}='{Priority}'"
        '    For Each drv As DataRowView In dv
        '        Dim GageName As String = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Gage_Name.ToString).ToString()

        '        cmbBendTypeGageWise_Mw.SelectedItem = GageName
        '    Next

        'End If
    End Sub

    Private Sub CmbBendTypeGageWise_Pw_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbBendTypeGageWise_Pw.SelectedIndexChanged

        'If rbPartTypewise.Checked Then

        '    Try
        '        'cmbThickness_Pw.Items.Clear()
        '        'cmbMaterialUsed_Pw.Items.Clear()
        '        'cmbMaterialSpec_Pw.Items.Clear()
        '        cmbGageName_Pw.Items.Clear()
        '        txtGageTable.Text = String.Empty

        '        Dim dv As New DataView(dt)
        '        Dim partTypeCol As String = ExcelUtil.ExcelSheetColumns.Type.ToString()
        '        Dim sizeCol As String = ExcelUtil.ExcelSheetColumns.Size.ToString()
        '        Dim gradeCol As String = ExcelUtil.ExcelSheetColumns.Grade.ToString()
        '        Dim BendTypeCol As String = ExcelUtil.ExcelSheetColumns.Bend_Type.ToString()
        '        Dim gageCol As String = ExcelUtil.ExcelSheetColumns.Gage_Name.ToString()
        '        Dim thicknessCol As String = ExcelUtil.ExcelSheetColumns.Thickness.ToString()
        '        Dim bendRadiusCol As String = ExcelUtil.ExcelSheetColumns.Bend_Radius.ToString()
        '        Dim gageTableCol As String = ExcelUtil.ExcelSheetColumns.Gage_Table.ToString()

        '        dv.RowFilter = $"{partTypeCol}='{cmbPartType_Pw.Text}' And {sizeCol}='{cmbSize_Pw.Text}' And {gradeCol}='{cmbGrade_Pw.Text}' And {BendTypeCol}='{cmbBendTypeGageWise_Pw.Text}'"
        '        Dim dtSize As DataTable = dv.ToTable()
        '        For Each drv As DataRowView In dv

        '            'Dim thickness1 As String = drv(ExcelUtil.excelSheetColumns.Thickness.ToString)

        '            'If Not cmbThickness_Pw.Items.Contains(thickness1) Then
        '            '    cmbThickness_Pw.Items.Add(thickness1)
        '            'End If

        '            'Dim bendRadius1 As String = drv(ExcelUtil.excelSheetColumns.Bend_Radius.ToString)

        '            'If Not cmbBendRadius_Pw.Items.Contains(bendRadius1) Then
        '            '    cmbBendRadius_Pw.Items.Add(bendRadius1)
        '            'End If

        '            Dim gageName As String = drv(ExcelUtil.ExcelSheetColumns.Gage_Name.ToString)

        '            If Not cmbGageName_Pw.Items.Contains(gageName) Then
        '                cmbGageName_Pw.Items.Add(gageName)
        '            End If

        '            txtGageTable.Text = drv(ExcelUtil.ExcelSheetColumns.Gage_Table.ToString)

        '        Next

        '        If cmbGageName_Pw.Items.Count > 0 Then
        '            cmbGageName_Pw.SelectedItem = cmbGageName_Pw.Items(0)
        '        End If

        '        'If cmbBendRadius_Pw.Items.Count > 0 Then
        '        '    cmbBendRadius_Pw.SelectedItem = cmbBendRadius_Pw.Items(0)
        '        'End If
        '    Catch ex As Exception

        '    End Try

        'End If

    End Sub

    Private Sub CmbBendTypeGageWise_Mw_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbBendTypeGageWise_Mw.SelectedIndexChanged


        FillBendRadiusDetails(dt)
        If cmbCategory.Text = "SheetMetal" Then
            'temp7APR2023
            FillGageName()
        End If
        Validation()
    End Sub

    Private Sub FillGageName()


        Try
            'cmbThickness_Pw.Items.Clear()

            Dim dv As New DataView(dt)
            Dim materialUsedCol As String = ExcelUtil.ExcelSheetColumns.Material_Used.ToString()
            Dim sizeCol As String = ExcelUtil.ExcelSheetColumns.Size.ToString()
            Dim gradeCol As String = ExcelUtil.ExcelSheetColumns.Grade.ToString()
            Dim BendTypeCol As String = ExcelUtil.ExcelSheetColumns.Bend_Type.ToString()
            Dim gageCol As String = ExcelUtil.ExcelSheetColumns.Gage_Name.ToString()
            Dim thicknessCol As String = ExcelUtil.ExcelSheetColumns.Thickness.ToString()



            'TEMP27APR20234
            'dv.RowFilter = $"{materialUsedCol}='{cmbMaterialUsed2_Mw.Text}' And {sizeCol}='{txtSize2_Mw.Text} ' And {gradeCol}='{txtGrade2_Mw.Text}' And {BendTypeCol}='{cmbBendTypeGageWise_Mw.Text}'"
            dv.RowFilter = $"{materialUsedCol}='{cmbMaterialUsed2_Mw.Text}'  And {gradeCol}='{txtGrade2_Mw.Text}' And {BendTypeCol}='{cmbBendTypeGageWise_Mw.Text}'"

            Dim dtSize As DataTable = dv.ToTable()

            For Each drv As DataRowView In dv

                'temp2Feb2024
                'txtGageName_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Gage_Name.ToString)
                cmbGageName.Text = drv(ExcelUtil.ExcelSheetColumns.Gage_Name.ToString)
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub FillBendRadiusDetails(ByVal dt As DataTable)

        Try

            txtBendRadius_Mw.Text = String.Empty

            Dim dv As New DataView(dt)

            Dim materialUsedCol As String = ExcelUtil.ExcelSheetColumns.Material_Used.ToString()
            Dim bendTypeCol As String = ExcelUtil.ExcelSheetColumns.Bend_Type.ToString()
            dv.RowFilter = $"{bendTypeCol}='{cmbBendTypeGageWise_Mw.Text}' And {materialUsedCol}='{cmbMaterialUsed2_Mw.Text}'"

            For Each drv As DataRowView In dv
                txtBendRadius_Mw.Text = drv(ExcelUtil.ExcelSheetColumns.Bend_Radius.ToString)
                Exit For

            Next
        Catch ex As Exception

        End Try
    End Sub

    Public Sub GetData()
        If IO.File.Exists(txtExcelPath.Text) Then
            dicData = ExcelUtil.ReadMaterials(txtExcelPath.Text)
            Dim ds As DataSet = dicData("Structure")
            dtstructure = ds.Tables(0)
            Dim ds2 As DataSet = dicData("SheetMetal")
            dtsheetmetal = ds2.Tables(0)

            FillCategory()
        Else
            MessageBox.Show("Please select valid excelpath for bec material")
        End If
    End Sub

    Private Sub BtnGetData_Click(sender As Object, e As EventArgs) Handles btnGetData.Click

        'If IO.File.Exists(txtExcelPath.Text) Then

        '    If GlobalEntity.dictRawMaterials.Count = 0 Then
        '        GlobalEntity.dictRawMaterials = ExcelUtil.ReadRawMaterials2(txtExcelPath.Text)
        '    End If

        '    dt = GlobalEntity.dictRawMaterials("BECMaterials").Tables(0)

        '    FillCategory(dt)

        'Else
        '    MsgBox("Please select material excel details.")
        'End If
        GetData()

    End Sub


    'temp2Feb2024
    Private Sub FillGageName1()
        Try

            Dim dv As DataView = New DataView(dt)
            Dim Becmaterialusedcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Used.ToString()

            dv.RowFilter = $" {Becmaterialusedcol}='{cmbMaterialUsed2_Mw.Text}' "
            cmbGageName.Items.Clear()

            'temp 6Feb2024
            txtBendRadius_Mw.Clear()

            For Each drv As DataRowView In dv
                Dim GageName As String = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Gage_Name.ToString).ToString()
                If Not cmbGageName.Items.Contains(GageName) Then
                    cmbGageName.Items.Add(GageName)
                End If
            Next

            If dv.Count = 1 And cmbGageName.Items.Count > 0 Then
                cmbGageName.SelectedItem = cmbGageName.Items(0)
            Else
                For Each drv As DataRowView In dv
                    Dim GageName As String = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Gage_Name.ToString).ToString()
                    cmbGageName.SelectedItem = GageName

                    'temp 6Feb2024
                    Dim BendRadius As String = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Bend_Radius.ToString).ToString()
                    txtBendRadius_Mw.Text = BendRadius
                Next
            End If
        Catch ex As Exception
            MessageBox.Show($"While fetching FillGageName", "Error")
            CustomLogUtil.Log("While fetching FillGageName", ex.Message, ex.StackTrace)
        End Try

        '16th october 2024
        cmbGageName.Enabled = False
    End Sub



    Private Sub cmbGageName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbGageName.SelectedIndexChanged
        'temp 6Feb2024
        FillBendRadius()
    End Sub

    'temp 6Feb2024
    Private Sub FillBendRadius()
        Try

            Dim dv As DataView = New DataView(dt)
            Dim GageName As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Gage_Name.ToString.ToString()
            Dim Becmaterialusedcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Used.ToString()

            dv.RowFilter = $"{GageName}='{cmbGageName.Text}' And {Becmaterialusedcol}='{cmbMaterialUsed2_Mw.Text}'"
            txtBendRadius_Mw.Clear()

            For Each drv As DataRowView In dv
                Dim BendRadius As String = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Bend_Radius.ToString).ToString()
                txtBendRadius_Mw.Text = BendRadius
            Next
        Catch ex As Exception
            MessageBox.Show($"While fetching FillBendRadius", "Error")
            CustomLogUtil.Log("While fetching FillBendRadius", ex.Message, ex.StackTrace)
        End Try
    End Sub
End Class