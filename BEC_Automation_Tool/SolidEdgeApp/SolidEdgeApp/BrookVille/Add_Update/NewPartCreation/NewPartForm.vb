Imports System.Runtime.InteropServices
'Imports SolidEdge.Framework.Interop
Imports WK.Libraries.BetterFolderBrowserNS
Public Class NewPartForm

    Dim objApp As SolidEdgeFramework.Application = Nothing

    Dim dtExcelTemplateData As DataTable = Nothing

    Private Sub OpenDocument()

        Try

            Dim docPath As String = IO.Path.Combine(txtTemplateDirectoryPath.Text, cmbOriginalTempalte.Text) '"C:\Program Files\Siemens\Solid Edge 2022\Template\ANSI Inch\ansi inch part.par"

            Dim tempDirPath As String = IO.Path.Combine(txtTemplateDirectoryPath.Text, "Temp")
            If Not IO.Directory.Exists(tempDirPath) Then
                IO.Directory.CreateDirectory(tempDirPath)
            End If

            Dim tempDocPath As String = IO.Path.Combine(tempDirPath, cmbOriginalTempalte.Text)
            If Not IO.File.Exists(tempDocPath) Then
                My.Computer.FileSystem.CopyFile(docPath, tempDocPath)
            End If

            If IO.File.Exists(tempDocPath) Then
                objApp.Documents.Open(tempDocPath)
            Else
                MsgBox($"File {docPath} is not exists.")
            End If


        Catch ex As Exception
            MsgBox($"Error in open document {ex.Message} {vbNewLine} {ex.StackTrace}")
        End Try
    End Sub

    Private Sub NewPartForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        SetSolidEdgeInstance()

        'OpenDocument()

    End Sub

    Private Sub SetSolidEdgeInstance()

        Try

            objApp = Marshal.GetActiveObject("SolidEdge.Application")

        Catch ex As Exception

            MsgBox($"Error in fetching the Solid-Edge instance {ex.Message} {vbNewLine} {ex.StackTrace}")

        End Try

    End Sub

    Private Sub btnBrowseExcel_Click(sender As Object, e As EventArgs) Handles btnBrowseExcel.Click
        Try
            Using dialog As New OpenFileDialog
                dialog.Filter = "Excel files (*.xlsx)|*.xlsx"
                If dialog.ShowDialog() <> DialogResult.OK Then Return
                txtExcelPath.Text = dialog.FileName
            End Using

            If IO.File.Exists(txtExcelPath.Text) Then

                If GlobalEntity.dictRawMaterials.Count = 0 Then
                    GlobalEntity.dictRawMaterials = ExcelUtil.ReadRawMaterials2(txtExcelPath.Text)
                End If

                dtExcelTemplateData = GlobalEntity.dictRawMaterials("BECMaterials").Tables(0)

                'FillCategory(dtExcelTemplateData, True)
                'dtExcelTemplateData = GlobalEntity.dictRawMaterials2("BECMaterials").Tables(0)

                FillCategory(dtExcelTemplateData, True)

            Else

                MsgBox("Please select material excel details.")

            End If

            'SetControlBrowseExcel()

            'Dim dtAssemblyData As DataTable = dgvDocumentDetails.DataSource
            'Dim dtExcelData As DataTable = GlobalEntity.dictRawMaterials("BECMaterials").Tables(0)
            'ValidateAssemblyData(dtAssemblyData, dtExcelData)

        Catch ex As Exception

        End Try

    End Sub

    Public Sub FillCategory(ByVal dt As DataTable, ByVal isSheetMetalPart As Boolean)


        Try
            cmbCategory.Items.Clear()


            Dim categoryCol As String = ExcelUtil.ExcelSheetColumns.Category.ToString()

            Dim categoryList As List(Of String) = dt.AsEnumerable() _
                                               .Select(Function(r) r.Field(Of String)(categoryCol)) _
                                               .Distinct() _
                                               .ToList()

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
            MsgBox($"Error in fill category {ex.Message} {vbNewLine} {ex.StackTrace}")
        End Try

    End Sub

    Private Sub cmbCategory_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCategory.SelectedIndexChanged
        Try

            FillType(dtExcelTemplateData)


        Catch ex As Exception
            MsgBox($"Error in category selection change {ex.Message} {vbNewLine} {ex.StackTrace}")
        End Try
    End Sub

    Public Sub FillType(ByVal dt As DataTable)

        Try

            cmbType.Items.Clear()


            Dim typeCol As String = ExcelUtil.ExcelSheetColumns.Type.ToString()

            Dim categoryCol As String = ExcelUtil.ExcelSheetColumns.Category.ToString()

            Dim dv As DataView = New DataView(dt)
            dv.RowFilter = $"{categoryCol}='{cmbCategory.Text}'"


            Dim typeList As List(Of String) = New List(Of String)()
            Dim dtSize As DataTable = dv.ToTable()
            For Each drv As DataRowView In dv

                Dim type1 As String = drv(ExcelUtil.ExcelSheetColumns.Type.ToString)

                If Not typeList.Contains(type1) Then
                    typeList.Add(drv(ExcelUtil.ExcelSheetColumns.Type.ToString))
                End If
            Next


            For Each typeName As String In typeList

                If Not cmbType.Items.Contains(typeName) Then
                    cmbType.Items.Add(typeName)
                End If

            Next

            If cmbType.Items.Count > 0 Then
                cmbType.SelectedItem = cmbType.Items(0)
            End If

            Dim mySource As AutoCompleteStringCollection = New AutoCompleteStringCollection()
            mySource.AddRange(typeList.ToArray)
            cmbType.AutoCompleteMode = AutoCompleteMode.SuggestAppend
            cmbType.AutoCompleteSource = AutoCompleteSource.CustomSource
            cmbType.AutoCompleteCustomSource = mySource


        Catch ex As Exception
            MsgBox($"Error in fill material used2 {ex.Message} {vbNewLine} {ex.StackTrace}")
        End Try

    End Sub

    Private Sub cmbType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbType.SelectedIndexChanged
        FillTemplateName(dtExcelTemplateData)
    End Sub

    Public Sub FillTemplateName(ByVal dt As DataTable)

        Try

            cmbTemplate.Items.Clear()


            Dim materialUsedCol As String = ExcelUtil.ExcelSheetColumns.Material_Used.ToString()

            Dim categoryCol As String = ExcelUtil.ExcelSheetColumns.Category.ToString()
            Dim typeCol As String = ExcelUtil.ExcelSheetColumns.Type.ToString()

            Dim dv As DataView = New DataView(dt)
            dv.RowFilter = $"{categoryCol}='{cmbCategory.Text}' And {typeCol}='{cmbType.Text}'"


            Dim templateList As List(Of String) = New List(Of String)()
            Dim dtTemplate As DataTable = dv.ToTable()
            For Each drv As DataRowView In dv
                Try

                    Dim type1 As String = drv(ExcelUtil.ExcelSheetColumns.Material_Used.ToString)

                    If Not templateList.Contains(type1) Then
                        templateList.Add(drv(ExcelUtil.ExcelSheetColumns.Material_Used.ToString))
                    End If

                Catch ex As Exception

                End Try
            Next


            For Each tempalteName As String In templateList

                If Not cmbTemplate.Items.Contains(tempalteName) Then
                    cmbTemplate.Items.Add(tempalteName)
                End If

            Next

            If cmbTemplate.Items.Count > 0 Then
                cmbTemplate.SelectedItem = cmbTemplate.Items(0)
            End If

            Dim mySource As AutoCompleteStringCollection = New AutoCompleteStringCollection()
            mySource.AddRange(templateList.ToArray)
            cmbTemplate.AutoCompleteMode = AutoCompleteMode.SuggestAppend
            cmbTemplate.AutoCompleteSource = AutoCompleteSource.CustomSource
            cmbTemplate.AutoCompleteCustomSource = mySource


        Catch ex As Exception
            MsgBox($"Error in fill material used2 {ex.Message} {vbNewLine} {ex.StackTrace}")
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
            MsgBox("Advanced Browse folder" + ex.Message + vbNewLine + ex.StackTrace)
        End Try


        Return folderpath
    End Function
    Private Sub btnTemplateLocation_Click(sender As Object, e As EventArgs) Handles btnTemplateLocation.Click
        Try
            Dim folderpath As String = ""
            folderpath = browseFolderAdvanced()
            If Not folderpath = String.Empty Then
                txtTemplateDirectoryPath.Text = folderpath
            End If
        Catch ex1 As Exception
            Try
                Dim MyFolderBrowser As Ookii.Dialogs.VistaFolderBrowserDialog = New Ookii.Dialogs.VistaFolderBrowserDialog
                Dim dlgResult As DialogResult = MyFolderBrowser.ShowDialog
                If dlgResult = DialogResult.OK Then
                    txtTemplateDirectoryPath.Text = MyFolderBrowser.SelectedPath
                End If
            Catch ex As Exception
                If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
                    Dim path As String = FolderBrowserDialog1.SelectedPath
                    txtTemplateDirectoryPath.Text = path
                End If
            End Try

        End Try
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnCreateDocument_Click(sender As Object, e As EventArgs) Handles btnCreateDocument.Click

        SetControlsEnability(False)
        OpenDocument()
        SetControlsEnability(True)
    End Sub

    Private Sub SetControlsEnability(ByVal flag As Boolean)

        btnBrowseExcel.Enabled = flag
        btnTemplateLocation.Enabled = flag
        btnCreateDocument.Enabled = flag
        btnClose.Enabled = flag
        cmbCategory.Enabled = flag
        cmbType.Enabled = flag
        cmbTemplate.Enabled = flag

    End Sub

    Private Sub cmbTemplate_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbTemplate.SelectedIndexChanged
        FillMaterialUsedWiseDetails(dtExcelTemplateData, cmbTemplate.Text)
    End Sub

    Private Sub FillMaterialUsedWiseDetails(ByVal dt As DataTable, ByVal materialUsed As String)

        Try

            cmbThickness.Items.Clear()
            cmbBECMaterial.Items.Clear()
            cmbOriginalTempalte.Items.Clear()

            Dim dv As DataView = New DataView(dt)

            Dim materialUsedCol As String = ExcelUtil.ExcelSheetColumns.Material_Used.ToString()
            dv.RowFilter = $"{materialUsedCol}='{cmbTemplate.Text}'"


            For Each drv As DataRowView In dv

                cmbThickness.Items.Add(drv(ExcelUtil.ExcelSheetColumns.Thickness.ToString))
                cmbBECMaterial.Items.Add(drv(ExcelUtil.ExcelSheetColumns.BEC_Material.ToString))
                cmbOriginalTempalte.Items.Add(drv(ExcelUtil.ExcelSheetColumns.Template.ToString))

                'txtSize2_Mw.Text = drv(ExcelUtil.excelSheetColumns.Size.ToString)
                'txtGrade2_Mw.Text = drv(ExcelUtil.excelSheetColumns.Grade.ToString)
                'txtGageName_Mw.Text = drv(ExcelUtil.excelSheetColumns.Gage_Name.ToString)
                'txtGageTable.Text = drv(ExcelUtil.excelSheetColumns.Gage_Table.ToString)
                'txtThickness2_Mw.Text = drv(ExcelUtil.excelSheetColumns.Thickness.ToString)
                'txtBendRadius_Mw.Text = drv(ExcelUtil.excelSheetColumns.Bend_Radius.ToString)
                'txtPartType_Mw.Text = drv(ExcelUtil.excelSheetColumns.Type.ToString)
                'txtMaterialSpec2_Mw.Text = drv(ExcelUtil.excelSheetColumns.Material_Specification.ToString)
                'txtBECMaterial2_Mw.Text = drv(ExcelUtil.excelSheetColumns.BEC_Material.ToString)
                Exit For

            Next

            If cmbThickness.Items.Count > 0 Then
                cmbThickness.SelectedItem = cmbThickness.Items(0)
            End If

            If cmbBECMaterial.Items.Count > 0 Then
                cmbBECMaterial.SelectedItem = cmbBECMaterial.Items(0)
            End If

            If cmbOriginalTempalte.Items.Count > 0 Then
                cmbOriginalTempalte.SelectedItem = cmbOriginalTempalte.Items(0)
            End If


        Catch ex As Exception
            MsgBox($"Error in fill material used wise details {ex.Message} {vbNewLine} {ex.StackTrace}")
        End Try
    End Sub
End Class

