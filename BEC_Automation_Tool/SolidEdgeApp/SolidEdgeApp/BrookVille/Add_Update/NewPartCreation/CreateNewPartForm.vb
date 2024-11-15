Imports System.Runtime.InteropServices
'Imports SolidEdge.Framework.Interop
Imports SolidEdgeFramework
Imports WK.Libraries.BetterFolderBrowserNS

Public Class CreateNewPartForm

    Dim dicData As Dictionary(Of String, DataSet) = New Dictionary(Of String, DataSet)()

    Dim dtstructure As DataTable = New DataTable("Structure")
    Dim dtsheetmetal As DataTable = New DataTable("SheetMetal")
    Dim dtData As DataTable = Nothing

    Dim Application As SolidEdgeFramework.Application = Nothing
    Dim Partdocument As SolidEdgePart.PartDocument = Nothing
    Dim objSheetMetalDocument As SolidEdgePart.SheetMetalDocument = Nothing
    Dim objDocument As SolidEdgePart.PartDocument = Nothing
    Dim SolidEdgePartPath As String = Nothing

    Private Sub btnGetData_Click(sender As Object, e As EventArgs) Handles btnGetData.Click
        Try

            'Dim aPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location
            'Dim dirPath As String = IO.Path.GetDirectoryName(aPath)
            Dim mtcExcelPath As String = txtExcelPath.Text 'Config.configObj.becMaterialExcelPath ' IO.Path.Combine(dirPath, "BEC_Material.xlsx")

            If IO.File.Exists(txtExcelPath.Text) Then
                dicData = ExcelUtil.ReadMaterials(mtcExcelPath)
                Dim ds As DataSet = dicData("Structure")
                dtstructure = ds.Tables(0)
                Dim ds2 As DataSet = dicData("SheetMetal")
                dtsheetmetal = ds2.Tables(0)

                FillCategory()
            Else
                MessageBox.Show($"Please select valid excelpath for bec material", "Message")
            End If

        Catch ex As Exception
            CustomLogUtil.Log("Error While fetching BEC Meterial Data From Excel", ex.Message, ex.StackTrace)
        End Try
    End Sub
    Private Sub CreateNewPartForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtExcelPath.Text = Config.configObj.becMaterialExcelPath

        TxtSolidEdgePartsTemplateDirectory.Text = Config.configObj.solidEdgePartTemplateDirectory

        CustomLogUtil.Heading("New Part Creation Form Open.....")

        'Dim aPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location
        'Dim dirPath As String = IO.Path.GetDirectoryName(aPath)
        'Dim mtcExcelPath As String = IO.Path.Combine(dirPath, "BEC_Material.xlsx")
        'dicData = ExcelUtil.ReadMaterials(mtcExcelPath)
        'Dim ds As DataSet = dicData("Structure")
        'dtstructure = ds.Tables(0)
        'Dim ds2 As DataSet = dicData("SheetMetal")
        'dtsheetmetal = ds2.Tables(0)

        'FillCategory()

        'Move above code in GetData button 

        Me.Text = Me.Text + $" ({GlobalEntity.Version})"

    End Sub

    Private Sub FillCategory()
        cmbCategory.Items.Add("Structure")
        cmbCategory.Items.Add("SheetMetal")
        cmbCategory.SelectedItem = cmbCategory.Items(0)
    End Sub

    Private Sub cmbCategory_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCategory.SelectedIndexChanged

        txtBendRadius.Text = String.Empty
        txtBendType.Text = String.Empty
        txtHeight.Text = String.Empty
        txtWidth.Text = String.Empty

        If cmbCategory.Text = "Structure" Then
            dtData = dtstructure.Copy()
            cmbGageTable.Visible = False
            lblGageTable.Visible = False
            cmbGageName.Visible = False
            lblGageName.Visible = False
            lblHeight.Visible = True
            txtHeight.Visible = True
            lblWidth.Visible = True
            txtWidth.Visible = True

            lblBendType.Visible = False
            txtBendType.Visible = False
            lblBendRadius.Visible = False
            txtBendRadius.Visible = False

            lblLinearLength.Visible = True
            txtLinearLength.Visible = True
        Else
            dtData = dtsheetmetal.Copy()
            cmbGageTable.Visible = True
            lblGageTable.Visible = True
            cmbGageName.Visible = True
            lblGageName.Visible = True
            lblHeight.Visible = False
            txtHeight.Visible = False
            lblWidth.Visible = False
            txtWidth.Visible = False
            txtTemplate.Text = "SheetMetal"

            lblBendType.Visible = True
            txtBendType.Visible = True
            lblBendRadius.Visible = True
            txtBendRadius.Visible = True

            lblLinearLength.Visible = False
            txtLinearLength.Visible = False
        End If
        FillType()
    End Sub

    Private Sub FillType()
        Try
            cmbType.Items.Clear()
            Dim typeCol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Type.ToString()

            Dim typeList As List(Of String) = dtData.AsEnumerable() _
                                               .Select(Function(r) r.Field(Of String)(typeCol)) _
                                               .Distinct() _
                                               .ToList()

            For Each type As String In typeList
                Try
                    If Not cmbType.Items.Contains(type) Then
                        cmbType.Items.Add(type)
                    End If
                Catch ex As Exception

                End Try


            Next

            If cmbType.Items.Count > 0 Then
                cmbType.SelectedItem = cmbType.Items(0)
            End If
        Catch ex As Exception
            MessageBox.Show($"Error While Filling Type", "Error")
            CustomLogUtil.Log("While Filling Type", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub cmbType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbType.SelectedIndexChanged
        FillMaterialUsed()
        LinearLength()

        If cmbType.Text.Contains("TUBING ROUND") Or cmbType.Text.Contains("ROUND BAR") Then
            lblDiameter.Visible = True
            txtDiameter.Visible = True

            txtHeight.Text = "0"
            txtWidth.Text = "0"
            txtHeight.Enabled = False
            lblHeight.Enabled = False

            txtWidth.Enabled = False
            lblWidth.Enabled = False
        Else
            lblDiameter.Visible = False
            txtDiameter.Visible = False


            txtHeight.Enabled = True
            lblHeight.Enabled = True
            txtWidth.Enabled = True
            lblWidth.Enabled = True
        End If

    End Sub

    Private Sub FillMaterialUsed()
        Try

            Dim dv As DataView = New DataView(dtData)
            Dim typeCol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Type.ToString()
            dv.RowFilter = $"Convert([{typeCol}], 'System.String') = '{cmbType.Text}'"

            cmbMaterialUsed.Items.Clear()
            cmbMaterialUsed.Text = String.Empty
            Dim cnt As Integer = 1
            For Each drv As DataRowView In dv

                If cnt = 190 Then
                    Debug.Print("aaaa")
                End If
                Dim materialUsed As String = String.Empty
                Try
                    If Not drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Used.ToString) = Nothing Then
                        materialUsed = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Used.ToString)
                    End If

                Catch ex As Exception

                End Try

                If Not cmbMaterialUsed.Items.Contains(materialUsed) Then
                    cmbMaterialUsed.Items.Add(materialUsed)
                End If
                Debug.Print($"#### {cnt.ToString()}")
                cnt = cnt + 1
            Next

            If cmbMaterialUsed.Items.Count > 0 Then
                cmbMaterialUsed.SelectedItem = cmbMaterialUsed.Items(0)
            End If
        Catch ex As Exception
            MessageBox.Show($"Error on FillMaterialUse", "Error")
            CustomLogUtil.Log("On Fill MaterialUse", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub cmbMaterialUsed_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbMaterialUsed.SelectedIndexChanged
        HeightWidth()
        ThicknessMaterialSpec()

        If cmbCategory.Text = "Structure" Then
            BECMaterial()
            FillTemplate()
            LinearLength()

        End If

        If cmbType.Text.Contains("TUBING ROUND") Or cmbType.Text.Contains("ROUND") Then
            SetDiameter()
        End If
        HighlightFields()
    End Sub
    Private Sub ThicknessMaterialSpecForStructure()
        Try
            If cmbCategory.Text = "Structure" Then
                Dim dv As DataView = New DataView(dtData)
                Dim typeCol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Type.ToString()
                Dim Becmaterialusedcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Used.ToString()
                Dim BecmaterialSpeccol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Specification.ToString()
                'Temp18APR2023
                Dim filter As String = $"Convert([{typeCol}], 'System.String') = '{cmbType.Text}' And Convert([{Becmaterialusedcol}], 'System.String') = '{cmbMaterialUsed.Text}' And Convert([{BecmaterialSpeccol}], 'System.String') = '{txtMaterialSpec.Text}'"
                dv.RowFilter = filter '$"{typeCol}='{cmbType.Text}'and {Becmaterialusedcol}='{cmbMaterialUsed.Text}'and {BecmaterialSpeccol}='{txtMaterialSpec.Text}'"

                cmbThickness.Items.Clear()
                cmbThickness.Text = String.Empty
                For Each drv As DataRowView In dv

                    Try
                        Dim Thickness As String = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Thickness.ToString).ToString()
                        If Not cmbThickness.Items.Contains(Thickness) And Not Thickness = String.Empty Then
                            cmbThickness.Items.Add(Thickness)
                        End If
                    Catch ex As Exception
                    End Try

                    'Try
                    '    txtMaterialSpec.Text = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Specification.ToString)
                    'Catch ex As Exception
                    'End Try


                Next

                If cmbThickness.Items.Count = 0 Then
                    cmbThickness.Items.Add("0")
                End If

                If cmbThickness.Items.Count > 0 Then
                    cmbThickness.SelectedItem = cmbThickness.Items(0)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show($"Error while Fetching ThicknessMaterialSpec", "Error")
            CustomLogUtil.Log("While Fetching ThicknessMaterialSpec", ex.Message, ex.StackTrace)
        End Try
    End Sub
    Private Sub ThicknessMaterialSpec()
        Try

            Dim dv As DataView = New DataView(dtData)
            Dim typeCol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Type.ToString()
            Dim Becmaterialusedcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Used.ToString()
            'Temp18APR2023
            Dim filter As String = $"Convert([{typeCol}], 'System.String') = '{cmbType.Text}' And Convert([{Becmaterialusedcol}], 'System.String') = '{cmbMaterialUsed.Text}'"
            dv.RowFilter = filter '$"{typeCol}='{cmbType.Text}'and {Becmaterialusedcol}='{cmbMaterialUsed.Text}'"

            cmbThickness.Items.Clear()
            cmbThickness.Text = String.Empty
            For Each drv As DataRowView In dv

                Try
                    Dim Thickness As String = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Thickness.ToString)
                    If Not cmbThickness.Items.Contains(Thickness) And Not Thickness = String.Empty Then
                        cmbThickness.Items.Add(Thickness)
                    End If
                Catch ex As Exception
                End Try

                Try
                    txtMaterialSpec.Text = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Specification.ToString)
                Catch ex As Exception
                End Try


            Next

            If cmbThickness.Items.Count = 0 Then
                cmbThickness.Items.Add("0")
            End If

            If cmbThickness.Items.Count > 0 Then
                cmbThickness.SelectedItem = cmbThickness.Items(0)
            End If
        Catch ex As Exception
            MessageBox.Show($"Error while Fetching ThicknessMaterialSpec", "Error")
            CustomLogUtil.Log("While Fetching ThicknessMaterialSpec", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub BECMaterial()
        Try


            If cmbCategory.Text = "Structure" Then


                Dim dv As DataView = New DataView(dtData)
                Dim typeCol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Type.ToString()
                Dim Becmaterialusedcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Used.ToString()
                Dim Thicknesscol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Thickness.ToString()
                'Temp18APR2023
                Dim filter As String = $"Convert([{typeCol}], 'System.String') = '{cmbType.Text}' And Convert([{Becmaterialusedcol}], 'System.String') = '{cmbMaterialUsed.Text}'"
                dv.RowFilter = filter '$"{typeCol}='{cmbType.Text}'and {Becmaterialusedcol}='{cmbMaterialUsed.Text}'" ' and {Thicknesscol}='{cmbThickness.Text}'"

                cmbBECMaterial.Items.Clear()
                cmbBECMaterial.Text = String.Empty

                For Each drv As DataRowView In dv

                    Dim BECmaterial As String = If(IsDBNull(drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.BEC_Material.ToString)), "", drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.BEC_Material.ToString))

                    If Not cmbBECMaterial.Items.Contains(BECmaterial) And Not BECmaterial = String.Empty Then

                        cmbBECMaterial.Items.Add(BECmaterial)

                    End If

                Next

                If cmbBECMaterial.Items.Count > 0 Then

                    cmbBECMaterial.SelectedItem = cmbBECMaterial.Items(0)

                End If

            Else

                Dim dv As DataView = New DataView(dtData)
                Dim typeCol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Type.ToString()
                Dim Becmaterialusedcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Used.ToString()
                Dim Thicknesscol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Thickness.ToString()
                dv.RowFilter = $"{typeCol}='{cmbType.Text}'and {Becmaterialusedcol}='{cmbMaterialUsed.Text}' and {Thicknesscol}='{cmbThickness.Text}'"

                cmbBECMaterial.Items.Clear()
                cmbBECMaterial.Text = String.Empty
                For Each drv As DataRowView In dv

                    Dim BECmaterial As String = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.BEC_Material.ToString)
                    If Not cmbBECMaterial.Items.Contains(BECmaterial) Then
                        cmbBECMaterial.Items.Add(BECmaterial)
                    End If

                Next

                If cmbBECMaterial.Items.Count > 0 Then
                    cmbBECMaterial.SelectedItem = cmbBECMaterial.Items(0)
                End If
            End If

        Catch ex As Exception
            MessageBox.Show($"Error While Fetching BEC Material", "Error")
            CustomLogUtil.Log("While Fetching BEC Material", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub cmbThickness_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbThickness.SelectedIndexChanged
        BECMaterial()
        If cmbCategory.Text = "Structure" Then
            LinearLength()
            HeightWidth()
        End If
        HighlightFields()
    End Sub
    Private Sub LinearLength()
        Try
            Dim dv As DataView = New DataView(dtData)
            Dim typeCol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Type.ToString()
            Dim Becmaterialusedcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Used.ToString()
            Dim Thicknesscol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Thickness.ToString()
            'Temp18APR2023
            Dim filter As String = $"Convert([{typeCol}], 'System.String') = '{cmbType.Text}' And Convert([{Becmaterialusedcol}], 'System.String') = '{cmbMaterialUsed.Text}'"
            dv.RowFilter = filter '$"{typeCol}='{cmbType.Text}'and {Becmaterialusedcol}='{cmbMaterialUsed.Text}'"

            For Each drv As DataRowView In dv
                Dim LinearLength As String = "0"
                Try
                    LinearLength = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Linear_Length.ToString)
                Catch ex As Exception
                End Try
                txtLinearLength.Text = LinearLength
            Next
        Catch ex As Exception

        End Try

    End Sub
    Private Sub HeightWidth()
        Try

            Dim dv As DataView = New DataView(dtData)
            Dim typeCol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Type.ToString()
            Dim Becmaterialusedcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Used.ToString()
            Dim Thicknesscol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Thickness.ToString()
            Dim BecMaterialcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.BEC_Material.ToString()
            'Temp18APR2023
            Dim filter As String = $"Convert([{typeCol}], 'System.String') = '{cmbType.Text}' And Convert([{Becmaterialusedcol}], 'System.String') = '{cmbMaterialUsed.Text}'"
            dv.RowFilter = filter '$"{typeCol}='{cmbType.Text}'and {Becmaterialusedcol}='{cmbMaterialUsed.Text}'" ' and {Thicknesscol}='{cmbThickness.Text}' and {BecMaterialcol}='{cmbBECMaterial.Text}'"

            txtHeight.Clear()
            txtWidth.Clear()
            For Each drv As DataRowView In dv

                Dim Height As String = "0"
                Try
                    Height = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Height.ToString)
                Catch ex As Exception
                End Try

                Dim width As String = "0"
                Try
                    width = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Width.ToString)
                Catch ex As Exception

                End Try


                txtHeight.Text = Height
                txtWidth.Text = width

            Next


        Catch ex As Exception

        End Try
    End Sub

    Private Sub getdata(ByVal tempDirPath As String, ByVal fileName As String)
        Try

            Dim dv As DataView = New DataView(dtData)
            Dim typeCol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Type.ToString()
            Dim Becmaterialusedcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Used.ToString()
            Dim Thicknesscol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Thickness.ToString()
            'Temp18APR2023
            Dim filter As String = $"Convert([{typeCol}], 'System.String') = '{cmbType.Text}' And Convert([{Becmaterialusedcol}], 'System.String') = '{cmbMaterialUsed.Text}'"
            dv.RowFilter = filter '$"{typeCol}='{cmbType.Text}'and {Becmaterialusedcol}='{cmbMaterialUsed.Text}'"


            For Each drv As DataRowView In dv

                Dim Height As String = "0"
                Dim width As String = "0"
                Dim Diameter As String = "0"
                Dim Gap As String = "0"
                Dim Thickness As String = "0"
                Dim LinearLength As String = "0"
                Try
                    Thickness = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Thickness.ToString)
                Catch ex As Exception

                End Try

                Dim Materialspe As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Specification.ToString()
                Dim BecMaterial As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.BEC_Material.ToString()
                Dim Description As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Description.ToString()

                Dim aPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location
                'Dim dirPath As String = IO.Path.GetDirectoryName(aPath)
                'Dim dirPath As String = IO.Path.GetDirectoryName(SolidEdgePartPath)' solidedgepartpath nothing problem
                Dim dirPath As String = TxtSolidEdgePartsTemplateDirectory.Text
                Dim tempalte As String = String.Empty

                Dim NewName = String.Empty '"C:\Users\milipatel\Desktop\Temp\TempPart.par"

                Dim documents As SolidEdgeFramework.Documents = Application.Documents

                If cmbCategory.Text = "Structure" Then

                    Try
                        If Not drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Height.ToString) Is Nothing Then
                            Height = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Height.ToString)
                        End If

                    Catch ex As Exception

                    End Try

                    Try
                        width = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Width.ToString)
                    Catch ex As Exception

                    End Try

                    Try
                        Diameter = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Diameter.ToString)
                    Catch ex As Exception

                    End Try

                    Try
                        Gap = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Gap.ToString)
                    Catch ex As Exception

                    End Try

                    Try
                        If Not drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Linear_Length.ToString) Is Nothing Then
                            LinearLength = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Linear_Length.ToString)
                        End If

                    Catch ex As Exception

                    End Try

                    Try
                        If Not drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Thickness.ToString) Is Nothing Then
                            Thickness = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Thickness.ToString)
                        End If

                    Catch ex As Exception

                    End Try

                    tempalte = IO.Path.Combine(dirPath, $"{txtTemplate.Text}.par") '"Angle.par")

                    NewName = IO.Path.Combine(tempDirPath, $"{fileName}.par")

                    'Application = SolidEdgeCommunity.SolidEdgeUtils.Connect(True, False)

                    Partdocument = DirectCast(documents.Open(tempalte), SolidEdgePart.PartDocument)

                    ReportVariables(Partdocument, Height, width, Thickness, Diameter, Gap, LinearLength, NewName, $"{txtTemplate.Text}.par")

                    'Partdocument.Close()
                    'Partdocument = Application.ActiveDocument

                Else
                    tempalte = IO.Path.Combine(dirPath, "SheetMetal.PSM")
                    NewName = IO.Path.Combine(tempDirPath, $"{fileName}.PSM")


                    'objSheetMetalDocument = DirectCast(documents.Open(tempalte), SolidEdgePart.SheetMetalDocument)

                    'objSheetMetalDocument.Close()

                    My.Computer.FileSystem.CopyFile(tempalte, NewName)

                    objSheetMetalDocument = DirectCast(documents.Open(NewName), SolidEdgePart.SheetMetalDocument)


                End If


                Exit For

            Next


        Catch ex As Exception

        End Try
    End Sub

    Private Sub ApplyMaterial_Gage()
        Try

            Dim materialName As String = cmbBECMaterial.Text

            If materialName = String.Empty Then
                Exit Sub
            End If

            Dim objMatTable As SolidEdgeFramework.MatTable = Application.GetMaterialTable()


            ' Set active document handle

            If (cmbCategory.Text = "Structure") Then
                objDocument = Application.ActiveDocument
                objMatTable.SetActiveDocument(objDocument)

                objMatTable.ApplyMaterialToDoc(objDocument, materialName, txtMaterialLibrary.Text)
            Else

                objSheetMetalDocument = Application.ActiveDocument
                objMatTable.SetActiveDocument(objSheetMetalDocument)

                objMatTable.ApplyMaterialToDoc(objSheetMetalDocument, materialName, txtMaterialLibrary.Text)

                objMatTable.SetDocumentToGageTableAssociation(objSheetMetalDocument, cmbGageName.Text, cmbGageTable.Text, True, True)

            End If

        Catch ex As Exception

        End Try
    End Sub

    'Dim objSheetMetalDocument As SolidEdgePart.SheetMetalDocument = Nothing
    'Dim objDocument As SolidEdgePart.PartDocument = Nothing
    Private Sub ApplyCustomProperties(ByVal custPropertiesObj As CustomProperties, ByRef objDocument As SolidEdgePart.PartDocument, ByRef objSheetMetalDocument As SolidEdgePart.SheetMetalDocument)
        Try

            If Not objDocument Is Nothing Then

                Dim propSets As SolidEdgeFramework.PropertySets = objDocument.Properties

                Dim custProps As Properties = propSets.Item("Custom")



                For Each prop1 As [Property] In custProps

                    Try


                        If prop1.Name = "Material Used" Then

                            prop1.Value = custPropertiesObj.materialUsed


                        ElseIf prop1.Name = "MATL SPEC" Then

                            prop1.Value = custPropertiesObj.materialSpec



                        End If

                    Catch ex As Exception
                    End Try

                Next
                propSets.Save()

            End If



            If Not objSheetMetalDocument Is Nothing Then

                Dim propSets As SolidEdgeFramework.PropertySets = objSheetMetalDocument.Properties

                Dim custProps As Properties = propSets.Item("Custom")

                For Each prop1 As [Property] In custProps

                    Try

                        If prop1.Name = "Material Used" Then

                            prop1.Value = custPropertiesObj.materialUsed


                        ElseIf prop1.Name = "MATL SPEC" Then

                            prop1.Value = custPropertiesObj.materialSpec

                        End If

                    Catch ex As Exception
                    End Try

                Next
                custProps.Add("Gage Name", cmbGageName.SelectedItem)
                propSets.Save()

            End If

            ' MsgBox("Property updation completed.")
        Catch ex As Exception

        End Try
    End Sub
    Public Shared Sub ReportVariables(ByVal document As SolidEdgeFramework.SolidEdgeDocument, ByVal Height As String, ByVal width As String, ByVal thickness As String, ByRef Diameter As String, ByRef Gap As String, ByRef LinearLength As String, ByRef NewName As String, ByRef templateName As String)

        Try

            Dim variables As SolidEdgeFramework.Variables = Nothing
            Dim variableList As SolidEdgeFramework.VariableList = Nothing
            Dim variable As SolidEdgeFramework.variable = Nothing
            Dim dimension As SolidEdgeFrameworkSupport.Dimension = Nothing


            Dim Height1 As String = Height
            Dim wth As String = width
            Dim thikns As String = thickness
            Dim diamtr As String = Diameter
            Dim gap1 As String = Gap
            Dim LinearLength1 As String = LinearLength

            Dim a As String = width
            variables = DirectCast(document.Variables, SolidEdgeFramework.Variables)


            variableList = DirectCast(variables.Query(pFindCriterium:="*", NamedBy:=SolidEdgeConstants.VariableNameBy.seVariableNameByBoth, VarType:=SolidEdgeConstants.VariableVarType.SeVariableVarTypeBoth), SolidEdgeFramework.VariableList)


            For Each variableListItem In variableList.OfType(Of Object)()
                Height1 = Height
                wth = width
                thikns = thickness
                diamtr = Diameter
                gap1 = Gap
                LinearLength1 = LinearLength
                Dim variableListItemType = SolidEdgeCommunity.Runtime.InteropServices.ComObject.GetType(variableListItem)


                Dim objectType = SolidEdgeCommunity.Runtime.InteropServices.ComObject.GetPropertyValue(Of SolidEdgeFramework.ObjectType)(variableListItem, "Type", CType(0, SolidEdgeFramework.ObjectType))


                Select Case objectType
                    Case SolidEdgeFramework.ObjectType.igDimension

                        dimension = DirectCast(variableListItem, SolidEdgeFrameworkSupport.Dimension)

                        ' MsgBox(dimension.Value)
                        'Console.WriteLine("Dimension: '{0}' = '{1}' ({2})", dimension.DisplayName, dimension.Value, objectType)

                        If templateName.ToUpper().Contains("INCH") Then
                            'INCH
                            Height1 = (Double.Parse(Height1) / 39.37).ToString()
                            wth = (Double.Parse(wth) / 39.37).ToString()
                            thikns = (Double.Parse(thickness) / 39.37).ToString()
                            diamtr = (Double.Parse(diamtr) / 39.37).ToString()
                            gap1 = (Double.Parse(gap1) / 39.37).ToString()
                            LinearLength1 = (Double.Parse(LinearLength1) / 39.37).ToString()
                        Else
                            'MM
                            Height1 = (Double.Parse(Height) / 1000).ToString()
                            wth = (Double.Parse(width) / 1000).ToString()
                            thikns = (Double.Parse(thickness) / 1000).ToString()
                            diamtr = (Double.Parse(diamtr) / 1000).ToString()
                            gap1 = (Double.Parse(gap1) / 1000).ToString()
                            LinearLength1 = (Double.Parse(LinearLength1) / 1000).ToString()
                        End If
                        If dimension.DisplayName = "Height" Then
                            dimension.Value = Height1
                        ElseIf dimension.DisplayName = "Width" Then
                            dimension.Value = wth
                        ElseIf dimension.DisplayName = "Thickness" Then
                            dimension.Value = thikns
                        ElseIf dimension.DisplayName = "Diameter" Then
                            dimension.Value = diamtr
                        ElseIf dimension.DisplayName = "Gap" Then
                            dimension.Value = gap1
                        ElseIf dimension.DisplayName = "LinearLength" Then
                            If Not LinearLength1 = "0" Then
                                dimension.Value = LinearLength1
                            End If

                        End If

                            Case SolidEdgeFramework.ObjectType.igVariable
                        variable = DirectCast(variableListItem, SolidEdgeFramework.variable)
                        'Console.WriteLine("Variable: '{0}' = '{1}' ({2})", variable.DisplayName, variable.Value, objectType)
                    Case Else

                End Select
            Next variableListItem

        Catch ex As Exception
            MessageBox.Show($"Error in set variables", "Error")
            CustomLogUtil.Log($"in set variables {ex.Message} {vbNewLine} {ex.StackTrace}")
        End Try
        ' partDocument.SaveAs(NewName)
        document.SaveAs(NewName)
        'document.Close(True)
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
    Public Sub HighlightFields()
        Dim Red As Int16 = 255 '229 
        Dim Green As Int16 = 102 '51 ' 255 
        Dim Blue As Int16 = 102 '51 '204 

        If txtBendType.Text = "Missing" Or txtBendType.Text = "" Or txtBendType.Text.Contains("Not") Or txtBendType.Text.Contains("Not".ToUpper) Or txtBendType.Text.Contains("Not".ToLower) Then
            txtBendType.BackColor = Color.FromArgb(Red, Green, Blue)
        Else
            txtBendType.BackColor = Color.White
        End If
        If txtBendRadius.Text = "Missing" Or txtBendRadius.Text = "" Or txtBendRadius.Text.Contains("Not") Or txtBendRadius.Text.Contains("Not".ToUpper) Or txtBendRadius.Text.Contains("Not".ToLower) Then
            txtBendRadius.BackColor = Color.FromArgb(Red, Green, Blue)
        Else
            txtBendRadius.BackColor = Color.White
        End If
        If cmbGageTable.Text = "Missing" Or cmbGageTable.Text = "" Or cmbGageTable.Text.Contains("Not") Or cmbGageTable.Text.Contains("Not".ToUpper) Or cmbGageTable.Text.Contains("Not".ToLower) Then
            cmbGageTable.BackColor = Color.FromArgb(Red, Green, Blue)
        Else
            cmbGageTable.BackColor = Color.White
        End If
        If cmbGageName.Text = "Missing" Or cmbGageName.Text = "" Or cmbGageName.Text.Contains("Not") Or cmbGageName.Text.Contains("Not".ToUpper) Or cmbGageName.Text.Contains("Not".ToLower) Then
            'cmbGageName.BackColor = Color.White
            cmbGageName.BackColor = Color.FromArgb(Red, Green, Blue)
        Else
            cmbGageName.BackColor = Color.White
        End If
        If cmbBECMaterial.Text = "Missing" Or cmbBECMaterial.Text = "" Or cmbBECMaterial.Text.Contains("Not") Or cmbBECMaterial.Text.Contains("Not".ToUpper) Or cmbBECMaterial.Text.Contains("Not".ToLower) Then
            cmbBECMaterial.BackColor = Color.FromArgb(Red, Green, Blue)
        Else
            cmbBECMaterial.BackColor = Color.White
        End If
        If cmbThickness.Text = "Missing" Or cmbThickness.Text = "" Or cmbThickness.Text.Contains("Not") Or cmbThickness.Text.Contains("Not".ToUpper) Or cmbThickness.Text.Contains("Not".ToLower) Then

            cmbThickness.BackColor = Color.FromArgb(Red, Green, Blue)
        Else
            cmbThickness.BackColor = Color.White
        End If

        If cmbCategory.Text = "Structure" Then
            If txtHeight.Text = "Missing" Or txtHeight.Text = "" Or txtHeight.Text.Contains("Not") Or txtHeight.Text.Contains("Not".ToUpper) Or txtHeight.Text.Contains("Not".ToLower) Then

                txtHeight.BackColor = Color.FromArgb(Red, Green, Blue)
            Else
                txtHeight.BackColor = Color.White
            End If

            If txtWidth.Text = "Missing" Or txtWidth.Text = "" Or txtWidth.Text.Contains("Not") Or txtWidth.Text.Contains("Not".ToUpper) Or txtWidth.Text.Contains("Not".ToLower) Then

                txtWidth.BackColor = Color.FromArgb(Red, Green, Blue)
            Else
                txtWidth.BackColor = Color.White
            End If
            If txtDiameter.Text = "Missing" Or txtDiameter.Text = "" Or txtDiameter.Text.Contains("Not") Or txtDiameter.Text.Contains("Not".ToUpper) Or txtDiameter.Text.Contains("Not".ToLower) Then

                txtDiameter.BackColor = Color.FromArgb(Red, Green, Blue)
            Else
                txtDiameter.BackColor = Color.White
            End If
            If txtLinearLength.Text = "Missing" Or txtLinearLength.Text = "" Or txtLinearLength.Text.Contains("Not") Or txtLinearLength.Text.Contains("Not".ToUpper) Or txtLinearLength.Text.Contains("Not".ToLower) Then

                txtLinearLength.BackColor = Color.FromArgb(Red, Green, Blue)
            Else
                txtLinearLength.BackColor = Color.White
            End If
        End If

    End Sub
    Private Sub btnCreatePart_Click(sender As Object, e As EventArgs) Handles btnCreatePart.Click
        btnCreatePart.Enabled = False
        PanelBody.Enabled = False
        Try

            Application = SolidEdgeCommunity.SolidEdgeUtils.Connect(True, False)

            Application.Visible = True

            Dim folderPath As String = OutputDirPath()

            getdata(folderPath, txtFileName.Text)

            If cmbCategory.Text = "Structure" Then

                Partdocument = Application.ActiveDocument

            Else

                objSheetMetalDocument = Application.ActiveDocument

            End If

            ApplyMaterial_Gage()

            Dim custPropObj As CustomProperties = New CustomProperties()

            custPropObj = SetCustomProp(custPropObj)

            ApplyCustomProperties(custPropObj, objDocument, objSheetMetalDocument)
            AppySummaryInfoProperties(objDocument, objSheetMetalDocument)

        Catch ex As Exception
            MessageBox.Show($"Error While Creating Part", "Error")
            CustomLogUtil.Log("While Creating Part", ex.Message, ex.StackTrace)
        End Try

        MessageBox.Show("Process done", "Message")
        CustomLogUtil.Heading("New Part Creation Form Process completed.....")
        'If IO.Directory.Exists(TxtSolidEdgePartsTemplateDirectory.Text) Then
        '    Process.Start(TxtSolidEdgePartsTemplateDirectory.Text)
        'End If

        btnCreatePart.Enabled = True
        PanelBody.Enabled = True

    End Sub
    'TEMP27-SEP-2023
    Public Sub AppySummaryInfoProperties(ByRef objDocument As SolidEdgePart.PartDocument, ByRef objSheetMetalDocument As SolidEdgePart.SheetMetalDocument)


        If Not objDocument Is Nothing Then

            Dim propSets As SolidEdgeFramework.PropertySets = objDocument.Properties
            For Each objProps In propSets

                If objProps.Name = "SummaryInformation" Then
                    For Each objProp In objProps
                        If objProp.Name = "Author" Then
                            Console.WriteLine(SystemInformation.UserName.ToString)
                            objProp.Value = SystemInformation.UserName.ToString
                            Exit For
                        End If
                    Next
                End If
            Next


            propSets.Save()
        End If
        If Not objSheetMetalDocument Is Nothing Then

            Dim propSets As SolidEdgeFramework.PropertySets = objSheetMetalDocument.Properties
            For Each objProps In propSets

                If objProps.Name = "SummaryInformation" Then
                    For Each objProp In objProps
                        If objProp.Name = "Author" Then
                            Console.WriteLine(SystemInformation.UserName.ToString)
                            objProp.Value = SystemInformation.UserName.ToString
                            Exit For
                        End If
                    Next
                End If
            Next


            propSets.Save()
        End If

    End Sub
    Private Function SetCustomProp(ByVal custPropObj As CustomProperties) As CustomProperties
        custPropObj = New CustomProperties()
        custPropObj.materialUsed = cmbMaterialUsed.Text
        custPropObj.materialSpec = txtMaterialSpec.Text
        custPropObj.gageName = cmbGageName.SelectedItem
        Return custPropObj
    End Function

    Private Sub cmbBECMaterial_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbBECMaterial.SelectedIndexChanged
        cmbGageTable.Items.Clear()

        If cmbCategory.Text = "Structure" Then
            'HeightWidth()
            'LinearLength()
        Else
        FillGageTable()
        End If
        HighlightFields()
    End Sub

    Private Sub FillGageTable()
        Try

            Dim dv As DataView = New DataView(dtData)
            Dim typeCol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Type.ToString()
            Dim Becmaterialusedcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Used.ToString()
            Dim Thicknesscol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Thickness.ToString()
            Dim BECMaterialcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.BEC_Material.ToString()

            dv.RowFilter = $"{typeCol}='{cmbType.Text}'and {Becmaterialusedcol}='{cmbMaterialUsed.Text}' and {Thicknesscol}='{cmbThickness.Text}' and {BECMaterialcol}='{cmbBECMaterial.Text}'"

            cmbGageTable.Items.Clear()

            For Each drv As DataRowView In dv

                Dim GageTable As String = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Gage_Table.ToString).ToString()
                If Not cmbGageTable.Items.Contains(GageTable) Then
                    cmbGageTable.Items.Add(GageTable)
                End If

            Next

            If cmbGageTable.Items.Count > 0 Then
                cmbGageTable.SelectedItem = cmbGageTable.Items(0)
            End If
        Catch ex As Exception
            MessageBox.Show($"Error While fetching FillGageTable", "Error")
            CustomLogUtil.Log("While fetching FillGageTable", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub cmbGageTable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbGageTable.SelectedIndexChanged
        FillGageName()
        FillBendDetails()
        HighlightFields()
    End Sub

    Private Sub FillGageName()
        Try

            Dim dv As DataView = New DataView(dtData)
            Dim typeCol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Type.ToString()
            Dim Becmaterialusedcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Used.ToString()
            Dim Thicknesscol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Thickness.ToString()
            Dim BECMaterialcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.BEC_Material.ToString()
            Dim GageTablecol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Gage_Table.ToString()
            Dim Becmaterialspecol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Specification.ToString()
            Dim Priorityspecol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Priority.ToString()
            Dim Priority As Integer = 0
            Dim value As Integer = 0

            dv.RowFilter = $"{typeCol}='{cmbType.Text}'and {Becmaterialusedcol}='{cmbMaterialUsed.Text}' and {Becmaterialspecol}='{txtMaterialSpec.Text}'and {BECMaterialcol}='{cmbBECMaterial.Text}' and {Thicknesscol}='{cmbThickness.Text}' "

            ' dv.RowFilter = $" {Becmaterialusedcol}='{cmbMaterialUsed.Text}' "


            cmbGageName.Items.Clear()

            For Each drv As DataRowView In dv
                If Priority = 0 Then
                    Priority = Convert.ToInt32(drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Priority.ToString).ToString())
                End If
                value = Convert.ToInt32(drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Priority.ToString).ToString())


                If (Priority > value) Then
                    Priority = value
                End If
                Dim GageName As String = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Gage_Name.ToString).ToString()
                If Not cmbGageName.Items.Contains(GageName) Then
                    cmbGageName.Items.Add(GageName)
                End If

            Next

            If dv.Count = 1 And cmbGageName.Items.Count > 0 Then
                cmbGageName.SelectedItem = cmbGageName.Items(0)
            Else
                dv.RowFilter = $"{typeCol}='{cmbType.Text}'and {Becmaterialusedcol}='{cmbMaterialUsed.Text}' and {Becmaterialspecol}='{txtMaterialSpec.Text}'and {BECMaterialcol}='{cmbBECMaterial.Text}' and {Thicknesscol}='{cmbThickness.Text}' and {Priorityspecol}='{Priority}'"

                'dv.RowFilter = $" {Becmaterialusedcol}='{cmbMaterialUsed.Text}' and {Priorityspecol}='{Priority}'"
                For Each drv As DataRowView In dv
                    Dim GageName As String = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Gage_Name.ToString).ToString()

                    cmbGageName.SelectedItem = GageName
                Next

            End If
        Catch ex As Exception
            MessageBox.Show($"While fetching FillGageName", "Error")
            CustomLogUtil.Log("While fetching FillGageName", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub cmbGageName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbGageName.SelectedIndexChanged

        FillBendDetails()
        HighlightFields()


    End Sub

    Private Sub FillBendDetails()
        Try

            Dim dv As DataView = New DataView(dtData)
            Dim typeCol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Type.ToString()
            Dim Becmaterialusedcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Used.ToString()
            Dim Thicknesscol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Thickness.ToString()
            Dim BECMaterialcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.BEC_Material.ToString()
            Dim GageTablecol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Gage_Table.ToString()
            Dim GageNamecol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Gage_Name.ToString()
            If cmbGageName.Text = "" Or cmbGageTable.Text = "" Then
                dv.RowFilter = $"{typeCol}='{cmbType.Text}'and {Becmaterialusedcol}='{cmbMaterialUsed.Text}' and {Thicknesscol}='{cmbThickness.Text}' and {BECMaterialcol}='{cmbBECMaterial.Text}' "
            Else
                dv.RowFilter = $"{typeCol}='{cmbType.Text}'and {Becmaterialusedcol}='{cmbMaterialUsed.Text}' and {Thicknesscol}='{cmbThickness.Text}' and {BECMaterialcol}='{cmbBECMaterial.Text}' and {GageTablecol}='{cmbGageTable.Text}'  and {GageNamecol}='{cmbGageName.Text}'"
            End If


            txtBendType.Text = String.Empty
            txtBendRadius.Text = String.Empty

            For Each drv As DataRowView In dv

                txtBendType.Text = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Bend_Type.ToString).ToString()
                txtBendRadius.Text = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Bend_Radius.ToString).ToString()

                Exit For
            Next

        Catch ex As Exception
            MessageBox.Show($"While Fetching FillBendDetails", "Error")
            CustomLogUtil.Log($"While Fetching FillBendDetails", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub FillTemplate()
        Try

            Dim dv As DataView = New DataView(dtData)
            Dim typeCol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Type.ToString()
            Dim Becmaterialusedcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Used.ToString()
            Dim Thicknesscol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Thickness.ToString()
            Dim BECMaterialcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.BEC_Material.ToString()
            Dim GageTablecol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Gage_Table.ToString()
            Dim GageNamecol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Gage_Name.ToString()
            'Temp18APR2023
            Dim filter As String = $"Convert([{typeCol}], 'System.String') = '{cmbType.Text}' And Convert([{Becmaterialusedcol}], 'System.String') = '{cmbMaterialUsed.Text}'"
            dv.RowFilter = filter '$"{typeCol}='{cmbType.Text}'and {Becmaterialusedcol}='{cmbMaterialUsed.Text}'"

            txtTemplate.Text = String.Empty

            For Each drv As DataRowView In dv

                txtTemplate.Text = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Template.ToString).ToString()

                Exit For
            Next

        Catch ex As Exception
            MessageBox.Show($"Error While Fetching FillTemplate", "Error")
            CustomLogUtil.Log($"While Filling FillTemplate", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub SetDiameter()
        Try
            txtDiameter.Text = "0"
            Dim dv As DataView = New DataView(dtData)
            Dim typeCol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Type.ToString()
            Dim Becmaterialusedcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Material_Used.ToString()
            Dim Thicknesscol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Thickness.ToString()
            Dim BECMaterialcol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.BEC_Material.ToString()
            Dim GageTablecol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Gage_Table.ToString()
            Dim GageNamecol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Gage_Name.ToString()
            Dim TemplateNamecol As String = ExcelUtil.ExcelSheetColumnsCreatePartStructure.Template.ToString()
            'Temp18APR2023
            Dim filter As String = $"Convert([{typeCol}], 'System.String') = '{cmbType.Text}' And Convert([{Becmaterialusedcol}], 'System.String') = '{cmbMaterialUsed.Text}' And Convert([{TemplateNamecol}], 'System.String') = '{txtTemplate.Text}'"
            dv.RowFilter = filter '$"{typeCol}='{cmbType.Text}'and {Becmaterialusedcol}='{cmbMaterialUsed.Text}' and {TemplateNamecol}='{txtTemplate.Text}'"

            For Each drv As DataRowView In dv

                txtDiameter.Text = drv(ExcelUtil.ExcelSheetColumnsCreatePartStructure.Diameter.ToString).ToString()

                Exit For
            Next
        Catch ex As Exception
            MessageBox.Show($"Error While fetching SetDiameter", "Error")
            CustomLogUtil.Log($"While fetching SetDiameter", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub txtTemplate_TextChanged(sender As Object, e As EventArgs) Handles txtTemplate.TextChanged



    End Sub

    Private Sub btnBrowseExcel_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnBrowseExcel_Click_1(sender As Object, e As EventArgs) Handles btnBrowseExcel.Click
        Using dialog As New OpenFileDialog
            dialog.Filter = "Excel files (*.xlsx)|*.xlsx"
            If dialog.ShowDialog() <> DialogResult.OK Then Return
            txtExcelPath.Text = dialog.FileName
        End Using
    End Sub

    Private Sub BtnGetSolidEdgeParts_Click(sender As Object, e As EventArgs) Handles BtnBrowseSolidEdgePartsTemplateDir.Click

        SolidEdgePartPath = OutputDirPath()
        SolidEdgePartPath += "\"
        TxtSolidEdgePartsTemplateDirectory.Text = SolidEdgePartPath
    End Sub

    Private Sub CreateNewPartForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'KillSolidEdgeProcess.Kill()
    End Sub


End Class