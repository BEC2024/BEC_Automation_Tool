Imports System.IO
Imports System.Text
Public Class CommonReportExcelUtil
    Dim dt As New DataTable()
    Dim dt2 As New DataTable()
    Dim Projectname As String = Nothing
    Dim er As String = "No"
    Dim notexist As String = "0"
    Dim No As String = "1"
    Dim Yes As String = "0"
    Dim Count_4_MTC_MTR = 0

    Public Sub CommonReport(comObj As CommonReport)

        Try
            dt.Clear()
            dt2.Clear()
            If Not dt.Columns.Count > 0 Then
                'dt.Columns.Add(New DataColumn("Rivision level"))'4
                'dt.Columns.Add(New DataColumn("Revision Number correct")) '23
                'dt.Columns.Add(New DataColumn("Component Name")) '22
                'heading
                dt.Columns.Add(New DataColumn("Type")) '0
                dt.Columns.Add(New DataColumn("Part Name")) '1
                dt.Columns.Add(New DataColumn("Category")) '2
                dt.Columns.Add(New DataColumn("Autor Name")) '3
                dt.Columns.Add(New DataColumn("C_Project Name")) '4
                dt.Columns.Add(New DataColumn("C_Part number match")) '5

                dt.Columns.Add(New DataColumn("C_Document number")) '6
                dt.Columns.Add(New DataColumn("C_Autor name")) '7
                dt.Columns.Add(New DataColumn("Dash In Unused Field")) '8
                dt.Columns.Add(New DataColumn("C_M2M Description")) '9
                dt.Columns.Add(New DataColumn("C_UOMs match")) '10
                dt.Columns.Add(New DataColumn("C_Interfenrce")) '11
                dt.Columns.Add(New DataColumn("inter part copies")) '12
                dt.Columns.Add(New DataColumn("part copies")) '13
                dt.Columns.Add(New DataColumn("C_broken file")) '14
                dt.Columns.Add(New DataColumn("C_Adjustable")) '15 till assembly
                dt.Columns.Add(New DataColumn("C_Mat'l spec")) '16
                dt.Columns.Add(New DataColumn("C_Material used")) '17
                dt.Columns.Add(New DataColumn("C_Removed Unused Features")) '18
                dt.Columns.Add(New DataColumn("C_ASTM minimum")) '19
                dt.Columns.Add(New DataColumn("C_Flat Pattern")) '20
                dt.Columns.Add(New DataColumn("C_Hole Tool Use")) '21

                dt.Columns.Add(New DataColumn("C_Virtual Thread")) '22
                dt.Columns.Add(New DataColumn("C_Defined Feature")) '23
                dt.Columns.Add(New DataColumn("C_Supprresed Feature")) '24
                dt.Columns.Add(New DataColumn("C_Vendor Part Number")) '25
                dt.Columns.Add(New DataColumn("C_Hardware Parts")) '26
                dt.Columns.Add(New DataColumn("C_M2M Sourced Marked")) '27
                dt.Columns.Add(New DataColumn("C_SE StatusBaseline")) '28
                dt.Columns.Add(New DataColumn("Exist Features")) '29
                dt.Columns.Add(New DataColumn("Mating Parts")) '30
                dt.Columns.Add(New DataColumn("Environment")) '31


                dt.Columns.Add(New DataColumn("R_Broken Geometry")) '32
                'dt.Columns.Add(New DataColumn("UpdateOnFileSave_R")) '33
                dt.Columns.Add(New DataColumn("R_ConstrainedFeatures")) '34
                dt.Columns.Add(New DataColumn("R_ALL Features Removed")) '35

                dt.Columns.Add(New DataColumn("R_HardwarePartBox")) '36

                '---------------------------------------------------------------------------------------------------------
                'MTC Assembly Manuals 
                dt.Columns.Add(New DataColumn("C_RefPart Occurrence Prop")) '37
                dt.Columns.Add(New DataColumn("C_Assembly_Features")) '38
                dt.Columns.Add(New DataColumn("C_Model_Preview")) '39
                dt.Columns.Add(New DataColumn("C_PartConstraint")) '40
                'MTC part mannuals
                dt.Columns.Add(New DataColumn("C_Material_Spec_Field")) '41
                dt.Columns.Add(New DataColumn("C_Mass&Density_Update")) '42

                'MTC Sheetmetal mannuals
                ' included C_MassAndDensity_Update
                dt.Columns.Add(New DataColumn("C_BendRadius_Update")) '43
                dt.Columns.Add(New DataColumn("C_FlatPattern_Update")) '44

                'MTC baseline manual
                dt.Columns.Add(New DataColumn("C_Hardware_Instances_&_Stackups")) '45
                dt.Columns.Add(New DataColumn("C_DimensionalGeometry_Match_For_ALL")) '46
                dt.Columns.Add(New DataColumn("C_Dimensional_Geometry_Match")) '47
                dt.Columns.Add(New DataColumn("C_Suppressed(Unused)Features")) '48
                dt.Columns.Add(New DataColumn("C_Child_Components_FullyConstrained_?")) '49
                dt.Columns.Add(New DataColumn("C_Simplified_AssemblyModel")) '50
                dt.Columns.Add(New DataColumn("C_ChildPart_OccurrenceProperties ")) '51
                dt.Columns.Add(New DataColumn("C_Vendor_MaterialData")) '52
                dt.Columns.Add(New DataColumn("C_Baseline_Mass_&_Density_Update")) '53
                dt.Columns.Add(New DataColumn("C_PMI_Instruction")) '54
                dt.Columns.Add(New DataColumn("C_BinderData")) '55
                dt.Columns.Add(New DataColumn("C_TerminalAssigned")) '56
                dt.Columns.Add(New DataColumn("C_ReferencePlanes")) '57

                '-------------------------------------------------------------------------------
                'MTR 
                'a+p+s
                dt.Columns.Add(New DataColumn("R_ConstrainedRefPlane")) '58
                dt.Columns.Add(New DataColumn("R_FullyConstrained")) '59
                dt.Columns.Add(New DataColumn("R_FullyConstrainedAndGround")) '60
                'a+p+s
                dt.Columns.Add(New DataColumn("R_ProperConstrained")) '61
                'a+p+s
                dt.Columns.Add(New DataColumn("R_NeedSimplified")) '62

                dt.Columns.Add(New DataColumn("R_PreviewModelGEE")) '63

                dt.Columns.Add(New DataColumn("R_WeightMass")) '64

                dt.Columns.Add(New DataColumn("R_OccurrenceProperties ")) '65
                dt.Columns.Add(New DataColumn("R_PatternsWithinPatterns")) '66
                dt.Columns.Add(New DataColumn("R_ExternallySuppliedModels")) '67
                dt.Columns.Add(New DataColumn("R_holePatterns")) '68
                dt.Columns.Add(New DataColumn("R_HoleDiameters")) '69
                dt.Columns.Add(New DataColumn("R_HoleClearances")) '70
                dt.Columns.Add(New DataColumn("R_HoleDistance")) '71
                dt.Columns.Add(New DataColumn("R_PartLineupWithMatingParts")) '72
                dt.Columns.Add(New DataColumn("R_FITS_In_MatingParts")) '73

                'MTR parts
                dt.Columns.Add(New DataColumn("R_Density")) '74
                'p+s
                dt.Columns.Add(New DataColumn("R_Material")) '75 
                'MTR Sheetmetal
                dt.Columns.Add(New DataColumn("R_UnusedSketches")) '76
                dt.Columns.Add(New DataColumn("R_linkSketches")) '77
                dt.Columns.Add(New DataColumn("R_BinderProcess")) '78
                dt.Columns.Add(New DataColumn("R_BendRadius")) '79
                '---------------------------------------------------------------------------------------
                dt.Columns.Add(New DataColumn("Total")) '37
                dt.Columns.Add(New DataColumn("Report Date")) '38
                dt.Columns.Add(New DataColumn("Modified Date")) '38



            End If

            If Not dt2.Columns.Count > 0 Then
                dt2.Columns.Add(New DataColumn("MTC FileName")) '
                dt2.Columns.Add(New DataColumn("MTR FileName")) '
            End If

            'function call
            checkExistFiles(comObj)
            createCommonReport(comObj)
            releaseObject(comObj)


        Catch ex As Exception
            MessageBox.Show($"Error While creating Common Reoprt for KPI", "Error")
            CustomLogUtil.Log($"While creating Common Reoprt for KPI", ex.Message, ex.StackTrace)
        End Try
    End Sub
    Public Sub checkExistFiles(ByVal comObj As CommonReport)
        Try

            Dim strFileName As String = comObj.dir + "\Generated Files(Do not delete).xlsx"
            If System.IO.File.Exists(strFileName) Then
                Dim _excel As New Microsoft.Office.Interop.Excel.Application
                Dim wBook As Microsoft.Office.Interop.Excel.Workbook
                Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet
                Dim misValue As Object = System.Reflection.Missing.Value
                wBook = _excel.Workbooks.Open(strFileName)
                wSheet = wBook.Sheets("Generated files")
                Dim str2 = wSheet.UsedRange.Rows.Count
                Dim ALLFiles As New ArrayList()
                ALLFiles = comObj.files
                Dim CheckedMTCFiles As New ArrayList()
                Dim CheckedMTRFiles As New ArrayList()
                For i = 2 To str2
                    If wSheet.Cells(i, 1).value = Nothing And wSheet.Cells(i, 2).value = Nothing Then

                        Exit For
                    End If
                    If Not wSheet.Cells(i, 1).value.ToString = "" Then
                        CheckedMTCFiles.Add(wSheet.Cells(i, 1).value.ToString)

                    End If
                    If Not wSheet.Cells(i, 2).value.ToString = "" Then
                        CheckedMTRFiles.Add(wSheet.Cells(i, 2).value.ToString)

                    End If
                Next

                For i = 0 To CheckedMTCFiles.Count - 1
                    comObj.MTCfiles.Remove(CheckedMTCFiles(i))
                Next
                For i = 0 To CheckedMTRFiles.Count - 1
                    comObj.MTRfiles.Remove(CheckedMTRFiles(i))
                Next
                wBook.Close()
                _excel.Quit()
                releaseObject(_excel)
                releaseObject(wBook)
                releaseObject(wSheet)
                killProcess()

                Count_4_MTC_MTR = 1
                MTC_main(comObj)
                Count_4_MTC_MTR = 2
                MTR_main(comObj)
                Count_4_MTC_MTR = Nothing
            Else
                Count_4_MTC_MTR = 1
                MTC_main(comObj)
                Count_4_MTC_MTR = 2
                MTR_main(comObj)
                Count_4_MTC_MTR = Nothing
            End If

        Catch ex As Exception
            MessageBox.Show($"Error While Checking Exsiting Files", "Error")
            CustomLogUtil.Log("Error While Checking Exsiting Files", ex.Message, ex.StackTrace)

        End Try
    End Sub
    Public Sub createCommonReport(ByVal comObj As CommonReport)
        Try
            Dim strFileName As String = comObj.dir + "\Generated Files(Do not delete).xlsx"
            If Not System.IO.File.Exists(strFileName) Then
                newKPI_Report(comObj)
            Else
                existKPI_Report(comObj)

            End If

        Catch ex As Exception

            MessageBox.Show($"Error While Creating Report", "Error")
            CustomLogUtil.Log($"Error While Creating Report", ex.Message, ex.StackTrace)
        End Try


    End Sub

    Public Sub NewCSV_Report(ByVal comObj As CommonReport)
        Try
            Dim writer As StreamWriter
            Dim sepChar As String = ","
            writer = New StreamWriter(comObj.dir + "\KPI_Report.csv")
            Dim sep As String = ""
            Dim builder As New Text.StringBuilder
            For Each col As DataColumn In dt.Columns
                builder.Append(col.ColumnName).Append(",")
                'sep = sepChar
            Next

            writer.WriteLine(builder.ToString())
            builder.Clear()

            For Each row As DataRow In dt.Rows
                For Each col As DataColumn In dt.Columns

                    builder.Append(row(col.ColumnName)).Append(",")
                    'sep = sepChar
                Next
                writer.WriteLine(builder.ToString())
                builder.Clear()
            Next
            writer.Dispose()
        Catch ex As Exception
            'MsgBox($"while Creating New KPI Report in CSV {vbNewLine}{ex.Message}{vbNewLine} {ex.StackTrace}")
        End Try
    End Sub
    Public Sub newKPI_Report(ByVal comObj As CommonReport)
        Try
            Dim _excel As New Microsoft.Office.Interop.Excel.Application
            Dim wBook As Microsoft.Office.Interop.Excel.Workbook
            ' Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet
            Dim wSheet2 As Microsoft.Office.Interop.Excel.Worksheet
            Dim formatRange As Microsoft.Office.Interop.Excel.Range

            Dim dc As System.Data.DataColumn
            Dim dr As System.Data.DataRow
            Dim colIndex As Integer = 0
            Dim rowIndex As Integer = 0

            wBook = _excel.Workbooks.Add()
            wSheet2 = wBook.ActiveSheet()


            wSheet2.Name = "Generated files"


            If comObj.MTCfiles.Count > comObj.MTRfiles.Count Then
                For comObj.i = 0 To comObj.MTCfiles.Count - 1
                    row2Data(comObj)
                Next
            ElseIf comObj.MTCfiles.Count < comObj.MTRfiles.Count Then
                For comObj.i = 0 To comObj.MTRfiles.Count - 1
                    row2Data(comObj)
                Next
            ElseIf comObj.MTCfiles.Count = comObj.MTRfiles.Count Then
                For comObj.i = 0 To comObj.MTRfiles.Count - 1
                    row2Data(comObj)
                Next
            End If

            colIndex = 0
            rowIndex = 0

            For Each dc In dt2.Columns
                colIndex += 1
                _excel.Cells(1, colIndex) = dc.ColumnName
            Next
            For Each dr In dt2.Rows
                rowIndex += 1
                colIndex = 0
                For Each dc In dt2.Columns
                    colIndex += 1
                    _excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
                Next
            Next


            formatRange = wSheet2.Range("a1", "b1")

            formatRange.EntireRow.Font.Bold = True
            formatRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            formatRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
            formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SandyBrown)

            formatRange = wSheet2.Range("a1", "b1")
            formatRange.RowHeight = 20
            wSheet2.Range("A1:B1").Columns.EntireColumn.AutoFit()
            formatRange.Font.Name = "Arial"
            formatRange.Font.Size = 11
            Dim str3 = wSheet2.UsedRange.Rows.Count
            For i = 2 To str3
                formatRange = wSheet2.Range("a" & i & ":b" & i & "")
                formatRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                formatRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
                formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow)
                wSheet2.Columns.AutoFit()

            Next
            'wBook.Sheets.Add()
            'wSheet = wBook.ActiveSheet()
            'wSheet.Name = "KPI Report"

            'colIndex = 0
            'rowIndex = 0

            'For Each dc In dt.Columns
            '    colIndex = colIndex + 1
            '    _excel.Cells(1, colIndex) = dc.ColumnName
            'Next
            'For Each dr In dt.Rows
            '    rowIndex = rowIndex + 1
            '    colIndex = 0
            '    For Each dc In dt.Columns
            '        colIndex = colIndex + 1
            '        _excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)

            '    Next
            'Next




            'formatRange = wSheet.Range("a1", "an1")

            'formatRange.EntireRow.Font.Bold = True
            'formatRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            'formatRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
            'formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SandyBrown)

            'formatRange = wSheet.Range("a1", "an1")
            'formatRange.RowHeight = 20
            'wSheet.Range("A1:AN1").Columns.EntireColumn.AutoFit()
            'formatRange.Font.Name = "Arial"
            'formatRange.Font.Size = 11
            'Dim str2 = wSheet.UsedRange.Rows.Count
            'For i = 2 To str2
            '    formatRange = wSheet.Range("a" & i & ":an" & i & "")
            '    formatRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            '    formatRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
            '    formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow)
            '    wSheet.Columns.AutoFit()

            'Next

            NewCSV_Report(comObj)

            Dim strFileName As String = comObj.dir + "\Generated Files(Do not delete).xlsx"
            _excel.DisplayAlerts = False

            wBook.SaveAs(strFileName)
            CustomLogUtil.Log("Report Successfully created")
            wBook.Close()
            _excel.Quit()
            releaseObject(_excel)
            releaseObject(wBook)
            ' releaseObject(wSheet)
            releaseObject(wSheet2)
            releaseObject(dt)
            releaseObject(dt2)
            killProcess()


        Catch ex As Exception
            ' MsgBox($"while Creating New excel: Generated Files(Do not delete) {vbNewLine}{ex.Message}{vbNewLine} {ex.StackTrace}")
        End Try
    End Sub
    Public Sub existKPI_Report(ByVal comObj As CommonReport)
        Try

            existCSV_report(comObj)
            Dim _excel As New Microsoft.Office.Interop.Excel.Application
            Dim wBook As Microsoft.Office.Interop.Excel.Workbook
            Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet
            ' Dim wSheet2 As Microsoft.Office.Interop.Excel.Worksheet
            Dim misValue As Object = System.Reflection.Missing.Value
            Dim strFileName As String = comObj.dir + "\Generated Files(Do not delete).xlsx"
            wBook = _excel.Workbooks.Open(strFileName)




            'wSheet2 = wBook.Worksheets("KPI Report")
            'wSheet2.Activate()


            'Dim colIndex As Integer = 0
            'Dim rowIndex As Integer = wSheet2.UsedRange.Rows.Count - 1
            'For Each dc In dt.Columns
            '    colIndex = colIndex + 1

            'Next
            'For Each dr In dt.Rows
            '    rowIndex = rowIndex + 1
            '    colIndex = 0
            '    For Each dc In dt.Columns
            '        colIndex = colIndex + 1
            '        _excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)

            '    Next
            'Next

            Dim formatRange As Microsoft.Office.Interop.Excel.Range

            'formatRange = wSheet2.Range("a1", "an1")

            'formatRange.EntireRow.Font.Bold = True
            'formatRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            'formatRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
            'formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SandyBrown)

            'formatRange = wSheet2.Range("a1", "an1")
            'formatRange.RowHeight = 20
            'wSheet2.Range("A1:AN1").Columns.EntireColumn.AutoFit()
            'formatRange.Font.Name = "Arial"
            'formatRange.Font.Size = 11
            'Dim str2 = wSheet2.UsedRange.Rows.Count
            'For i = 2 To str2
            '    formatRange = wSheet2.Range("a" & i & ":an" & i & "")
            '    formatRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            '    formatRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
            '    formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow)
            '    wSheet2.Columns.AutoFit()

            'Next


            wSheet = wBook.Sheets("Generated files")
            wSheet.Activate()
            Dim str2 = wSheet.UsedRange.Rows.Count

            If comObj.MTCfiles.Count > comObj.MTRfiles.Count Then
                For comObj.i = 0 To comObj.MTCfiles.Count - 1
                    row2Data(comObj)
                Next
            ElseIf comObj.MTCfiles.Count < comObj.MTRfiles.Count Then
                For comObj.i = 0 To comObj.MTRfiles.Count - 1
                    row2Data(comObj)
                Next
            ElseIf comObj.MTCfiles.Count = comObj.MTRfiles.Count Then
                For comObj.i = 0 To comObj.MTRfiles.Count - 1
                    row2Data(comObj)
                Next
            End If

            Dim colIndex = 0
            Dim rowIndex = wSheet.UsedRange.Rows.Count - 1

            For Each dc In dt2.Columns
                colIndex += 1

            Next
            For Each dr In dt2.Rows
                rowIndex += 1
                colIndex = 0
                For Each dc In dt2.Columns
                    colIndex += 1
                    _excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)

                Next
            Next


            formatRange = wSheet.Range("a1", "b1")

            formatRange.EntireRow.Font.Bold = True
            formatRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            formatRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
            formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SandyBrown)

            formatRange = wSheet.Range("a1", "b1")
            formatRange.RowHeight = 20
            wSheet.Range("A1:b1").Columns.EntireColumn.AutoFit()
            formatRange.Font.Name = "Arial"
            formatRange.Font.Size = 11
            Dim str3 As Integer = rowIndex + 1
            For i = 2 To str3
                formatRange = wSheet.Range("a" & i & ":b" & i & "")
                formatRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                formatRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
                formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow)
                wSheet.Columns.AutoFit()

            Next




            _excel.DisplayAlerts = False
            wBook.SaveAs(strFileName)
            CustomLogUtil.Log("Report Successfully created")
            wBook.Close()
            _excel.Quit()
            releaseObject(_excel)
            releaseObject(wBook)
            releaseObject(wSheet)
            'releaseObject(wSheet2)
            killProcess()

        Catch ex As Exception
            'MsgBox($"while Creating exist KPI report {vbNewLine}{ex.Message}{vbNewLine} {ex.StackTrace}")
        End Try
    End Sub
    Public Sub existCSV_report(ByVal comObj As CommonReport)
        Try


            Dim writer As StreamWriter
            Dim builder As New Text.StringBuilder
            writer = New StreamWriter(comObj.dir + "\KPI_Report.csv", append:=True)

            For Each row As DataRow In dt.Rows
                For Each col As DataColumn In dt.Columns

                    builder.Append(row(col.ColumnName)).Append(",")
                    'sep = sepChar
                Next
                writer.WriteLine(builder.ToString())
                builder.Clear()
            Next
            writer.Dispose()
        Catch ex As Exception
            'MsgBox($"while Creating exist KPI report {vbNewLine}{ex.Message}{vbNewLine} {ex.StackTrace}")
        End Try
    End Sub
    Public Sub MTC_main(ByVal comObj As CommonReport)
        For comObj.i = 0 To comObj.MTCfiles.Count - 1
            comObj.excelFilePath = comObj.MTCfiles(comObj.i)
            MTC_Assembly(comObj)
            MTC_Part(comObj)
            MTC_Sheetmetal(comObj)
            MTC_baseline(comObj)
            MTC_Electrical(comObj)
        Next

    End Sub

    Public Sub MTR_main(ByVal comObj As CommonReport)
        For comObj.i = 0 To comObj.MTRfiles.Count - 1
            comObj.excelFilePath = comObj.MTRfiles(comObj.i)

            MTR_Assembly(comObj)
            MTR_Part(comObj)
            MTR_Sheetmetal(comObj)
        Next
    End Sub

    Public Sub MTC_Assembly(ByVal comObj As CommonReport)
        Try
            Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
            xlsWB = xlsApp.Workbooks.Open(comObj.excelFilePath, Nothing, False)

            Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing



            xlsSheet1 = xlsWB.Sheets("Assembly")
            xlsSheet1.Activate()
            Dim RowCnt = xlsSheet1.UsedRange.Rows.Count

            Dim ColCnt = xlsSheet1.UsedRange.Columns.Count


            xlsCell1 = xlsSheet1.UsedRange()


            Dim count = Nothing



            For i = 4 To ColCnt
                Dim row1 As DataRow = dt.NewRow()
                'initialization

                '0file
                comObj.fileName0 = "MTC"

                '1PartName Value(Error)
                comObj.partName1 = If(xlsSheet1.Rows(1).Cells(i).value = Nothing, er, xlsSheet1.Rows(1).Cells(i).value.ToString())


                '2Category
                comObj.category2 = "Assembly"




                '3author name/Who is the Author of the file?
                comObj.AuthorName3 = If(xlsSheet1.Rows(5).Cells(i).value = Nothing, er, xlsSheet1.Rows(5).Cells(i).value.ToString())


                '4projectname value/'  Is the correct "Project" used in the "Project Name" field? (Ensure consistent project titleIs used within a project) *
                comObj.projectName4 = If(xlsSheet1.Rows(1).Cells(4).value = Nothing, er, xlsSheet1.Rows(1).Cells(4).value.ToString())
                Projectname = comObj.projectName4


                '5part number match value/Is the part number match with M2M? *
                comObj.partNumberMatch5 = If(xlsSheet1.Rows(3).Cells(i).value = Nothing, er, xlsSheet1.Rows(3).Cells(i).value.ToString())





                '6document number/Is the Document Number correct? (BEC designed parts must have proper doc number.Hardware And other exceptions are N/A) *
                comObj.documentNumber6 = If(xlsSheet1.Rows(8).Cells(i).value = Nothing, er, xlsSheet1.Rows(8).Cells(i).value.ToString())



                '7authorNameIs the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) 
                comObj.AuthorName7 = If(xlsSheet1.Rows(9).Cells(i).value = Nothing, er, xlsSheet1.Rows(9).Cells(i).value.ToString())


                '8dashinunusedfield/Do all technically unused properties have a "dash" populated? 
                comObj.DashInUnusedFiled8 = If(xlsSheet1.Rows(10).Cells(i).value = Nothing, er, xlsSheet1.Rows(10).Cells(i).value.ToString())

                '9 m2mdescription/Does the Model Title MATCH the M2M Item Master (INV) Description Field?
                comObj.m2mDiscriptionMatch9 = If(xlsSheet1.Rows(11).Cells(i).value = Nothing, er, xlsSheet1.Rows(11).Cells(i).value.ToString())


                '10uomsmatch/Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
                comObj.uomsMatch10 = If(xlsSheet1.Rows(12).Cells(i).value = Nothing, er, xlsSheet1.Rows(12).Cells(i).value.ToString())



                '11interface/Is interfernces found in assembly?
                comObj.interface11 = If(xlsSheet1.Rows(14).Cells(i).value = Nothing, er, xlsSheet1.Rows(14).Cells(i).value.ToString())


                '12interpartcopiesIs interfernces found in assembly?
                comObj.interpartcopies12 = If(xlsSheet1.Rows(15).Cells(i).value = Nothing, er, xlsSheet1.Rows(15).Cells(i).value.ToString())


                '13partcopiespart copies detected
                comObj.partcopies13 = If(xlsSheet1.Rows(16).Cells(i).value = Nothing, er, xlsSheet1.Rows(16).Cells(i).value.ToString())


                '14brokenfilebroken  file Path detected
                comObj.brokefile14 = If(xlsSheet1.Rows(17).Cells(i).value = Nothing, er, xlsSheet1.Rows(17).Cells(i).value.ToString())


                '15Adjustable/Is the part "Adjustable"? (part should NOT be adjustable) *
                comObj.adjustable15 = If(xlsSheet1.Rows(19).Cells(i).value = Nothing, er, xlsSheet1.Rows(17).Cells(i).value.ToString())


                '16Mat'l spec/Is the mat'l spec field populated? (indicated by RED in metadata audit file)
                comObj.Matl_spec16 = notexist

                '17MaterialUsed/Is the material used field populated? *
                comObj.MaterialUsed17 = notexist

                '18RemovedUnusedFeatures/Have any suppressed (unused) features been removed from the model Pathfinder? *
                comObj.RemovedUnusedFeatures18 = notexist


                '19ASTMminimum/Is the bend radius of the part equal to or above the ASTM minimum? *
                comObj.ASTMminimum19 = notexist

                '20FlatPattern/. Is the flat pattern turned ON? (Applicable only to parts needing dxf or cut length) *
                comObj.FlatPattern20 = notexist

                '21HoleToolUse/Did I use the hole tool for ALL holes requiring fasteners? (include clearance holes forhardware, tapped holes, And slots) *
                comObj.HoleToolUse21 = notexist




                '22virtualThread/Virtual thread applied for Fasteners?
                comObj.VirtualThread22 = notexist

                '23defined feature
                comObj.defineFeatures23 = notexist

                '24Supperessed Feature
                comObj.suppressedFeature24 = notexist

                '25Vendorpartnumber
                comObj.vendorPartNumber25 = notexist

                '26Hardwareparts
                comObj.HardwareParts26 = notexist

                '27m2mSourceMarked
                comObj.m2mSourceMarked27 = notexist

                '28SEStatus
                comObj.SEstatus28 = notexist

                '29exist features
                comObj.ExistFeatures29 = notexist

                '30MatingParts
                comObj.MatingParts30 = notexist

                '31 environment
                comObj.Environment31 = notexist

                '32Geometry
                comObj.Geometry32 = notexist

                '33UpdateOnFileSave33
                ' comObj.UpdateOnFileSave33 = notexist

                '34ConstrainedFeatures34
                comObj.ConstrainedFeatures34 = notexist

                '35AllFeaturesRemoved35
                comObj.AllFeaturesRemoved35 = notexist

                '36HardwarePartBox
                comObj.HardwarePartBox36 = notexist

                '-------------------------------------------------------------------------------
                'mannual Questions
                '37 RefPartOccurrenceProp/reference parts?(higher level, report To PL, physical properties, all Set To "NO") * (Use occurrence property macro to recorrect setting)
                comObj.C_RefPartOccurrenceProp40 = If(xlsSheet1.Rows(20).Cells(i).value = Nothing, er, xlsSheet1.Rows(20).Cells(i).value.ToString())

                '38 AssemblyFeatures/Are the assembly features minimized? *
                comObj.C_AssemblyFeatures41 = If(xlsSheet1.Rows(22).Cells(i).value = Nothing, er, xlsSheet1.Rows(22).Cells(i).value.ToString())

                '39ModelPreview/Is the Preview of the model saved in conformance to the GEE? (On button select,
                comObj.C_ModelPreview42 = If(xlsSheet1.Rows(23).Cells(i).value = Nothing, er, xlsSheet1.Rows(23).Cells(i).value.ToString())

                '40PartConstraint/Record number of parts that are NOT fully constrained. **Record "0" if all parts are constrained correctly. (All parts locked and constraints convey design intent - i.e incorrect to "ground" parts when actual intent is for mating surfaces) *
                comObj.C_PartConstraint43 = If(xlsSheet1.Rows(24).Cells(i).value = Nothing, er, xlsSheet1.Rows(24).Cells(i).value.ToString())

                '---------------------------------------------------------------------------------------------------------
                'mtc part

                '41Material_Spec_Field/Is the material spec field filled out with the proper designation? (Modification,Purchased, Assembly, etc.) *
                comObj.C_Material_Spec_Field44 = notexist

                '--------------------------------------------------------------------------------------------------------------------------------
                'mtc sheetmetal +part mass and density

                '42MassAndDensity_Update/Is the density defined and the mass up to date? * (macro will auto-update if file not read only)
                comObj.C_MassAndDensity_Update45 = notexist

                '43Bend_Radius_Update/Did I ensure the bend radius was producible? (Evaluate based on NPM0001C1 and/or consult ProdE/ManE) *
                comObj.C_Bend_Radius_Update46 = notexist

                '44 FlatPattern_Update/Is the flat pattern up to date? *
                comObj.C_FlatPattern_Update47 = notexist

                '--------------------------------------------------------------------------------------------------------------------------------
                'MTC Baseline
                '45 HardwareInstancesAndStackups/Are all instances of hardware consistent and have logical stackups? (Examples: 1.a SS nut applied to a zinc coated bolt having a copper washer; 2. Nine(9) instances of grade 8 and one(1) instance of grade 5 within a mounting configuration) *
                comObj.C_HardwareInstancesAndStackups48 = notexist

                '46 DimensionalGeometry_Match_For_ALL/Has the dimensional geometry of the model been verified to match the vendor supplied data? (This is to include use of the hole tool for ALL features requiring fasteners) *
                comObj.C_DimensionalGeometry_Match_For_ALL49 = notexist

                '47 DimensionalGeometry_Match/Has the dimensional geometry of the model been verified to match the vendor supplied data? *
                comObj.C_DimensionalGeometry_Match50 = notexist

                '48 suppressedUnusedFeatures/ Have any suppressed (unused) features been removed from the model Pathfinder? *
                comObj.C_suppressedUnusedFeatures51 = notexist

                '49 ChildComponentsFullyConstrained/Are all child components fully constrained? (Ground constraints should be minimized for BEC generated models, but are acceptable for Vendor supplied assemblies) *
                comObj.C_ChildComponentsFullyConstrained52 = notexist

                '50 SimplifiedAssemblyModel/Is the assembly model simplified in design? (unnecessary details or "internal" parts should be removed if possible) *
                comObj.C_SimplifiedAssemblyModel53 = notexist

                '51 ChildPartOccurrenceProperties/Have the required occurrence properties for ALL child components been set to NO? (Reports/Parts List, Physical Properties, and Interference Analysis should all be set to NO) *
                comObj.C_ChildPartOccurrenceProperties54 = notexist

                '52 VendorMaterialData/Is the appropriate material applied to the model per Vendor supplied data? (If the component contains multiple material PURCHASED should be utilized) *
                comObj.C_VendorMaterialData55 = notexist

                '53 BaselineMassAndDensityUpdate/Is the density defined and the mass up to date? ("Update on file save" should be unchecked. When a Vendor weight is known "User-defined properties" should be checked and the mass applied) * (Manual check -if read-only)
                comObj.C_BaselineMassAndDensityUpdate56 = notexist

                '54 PMI_Instruction/Is PMI included for any other further or special situations? (kit identification, missing vendor information, etc.) *
                comObj.C_PMI_Instruction57 = notexist

                '55 BinderData/Does the model have a Binder attached and is properly filled out? (A hyperlink to the Vendor supplied data sheet located on the VAULT should be included) *
                comObj.C_BinderData58 = notexist

                '56 TerminalAssigned/Have the required terminals been assigned to the part model? *
                comObj.C_TerminalAssigned59 = notexist

                '57 ReferencePlanes/ Has the model been constructed on the correct reference planes per GEE0008? *
                comObj.C_ReferencePlanes60 = notexist

                '------------------------------------------------------------------------------------------------------------------------------------------
                'MTR assembly + sheetmetal+part 

                'assembly+part+sheetmetal
                '58 ConstrainedRefPlane/Verify the part model has been created and constrained to the correct reference plane.
                comObj.R_ConstrainedRefPlane61 = notexist

                '59 R_FullyConstrained/Verify that ALL parts/sub-assemblies have been fully constrained
                comObj.R_FullyConstrained62 = notexist

                '60 R_FullyConstrainedAndGround/Verify that ALL parts/sub-assemblies have been properly constrained***The “Ground” constraint should only be applied to the first component***
                comObj.R_FullyConstrainedAndGround63 = notexist

                'a+p
                '61 R_ProperConstrained/Verify proper constraints by utilizing “Assembly Relationship Manager”
                comObj.R_ProperConstrained64 = notexist

                'a+p+s
                '62 R_NeedSimplified/Verify if the assembly can be or needs to be simplified
                comObj.R_NeedSimplified65 = notexist

                '63 R_PreviewModelGEE/ matching part, background,to turn off shadows, sketches, coordinate systems And PMI reference planes off)
                comObj.R_PreviewModelGEE66 = notexist

                '64 R_WeightMass/Verify that the weight and mass are correct
                comObj.R_WeightMass67 = notexist

                '65 R_OccurrenceProperties/Verify the occurrence properties are set properly for the assembly model,
                comObj.R_OccurrenceProperties68 = notexist

                '66 R_PatternsWithinPatterns/Verify that “patterns” do not exist within other patterns. If present, they should be removed
                comObj.R_PatternsWithinPatterns69 = notexist

                '67 R_ExternallySuppliedModels/Do externally supplied models exist within the assembly? Are the models a surface model? Are the models a solid model?
                comObj.R_ExternallySuppliedModels70 = notexist

                '68 R_holePatterns/Verify that ALL hole patterns align properly with mating parts/components
                comObj.R_holePatterns71 = notexist

                '69 R_HoleDiameters/Verify that hole diameters are correct for size application
                comObj.R_HoleDiameters72 = notexist

                '70 R_HoleClearances/Verify that hole clearances are correct for fasteners used
                comObj.R_HoleClearances73 = notexist

                '71 R_HoleDistance/Verify that hole distance from edge of material is appropriate
                comObj.R_HoleDistance74 = notexist

                '72 R_PartLineupWithMatingParts/Verify that the part will line up correctly with mating parts
                comObj.R_PartLineupWithMatingParts75 = notexist

                '73 R_FITS_In_MatingParts/Verify that “FITS” are correct to mating parts Use Engineering Document MEA0001
                comObj.R_FITS_In_MatingParts76 = notexist

                'part
                '74 R_Density/Verify that the correct density has been applied to the model
                comObj.R_Density77 = notexist

                'p+s
                '75 R_Material/Verify the correct material is assigned to the part for the application
                comObj.R_Material = notexist

                'sheetmetal
                '76 R_UnusedSketches/Verify that any unused sketches have been removed
                comObj.R_UnusedSketches78 = notexist

                '77 R_linkSketches/Verify that any links within a sketch have been removed
                comObj.R_linkSketches79 = notexist

                '78 R_BinderProcess/Verify the Binder Document Process (GEE0001) was followed
                comObj.R_BinderProcess80 = notexist

                '79 R_BendRadius/Verify the bend radius meets requirements of the material selected
                comObj.R_BendRadius81 = notexist



                '-------------------------------------------------------------------------------------
                '38 date
                comObj.date38 = Date.Now()

                '39 lastModifiedDate
                comObj.lastModifiedDate39 = If(xlsSheet1.Rows(18).Cells(i).value = Nothing, er, xlsSheet1.Rows(18).Cells(i).value.ToString())

                '------------------------------------------------------------------

                row1Data(comObj)



                '3Rivisionlevel/What is the revision level? *
                'If xlsSheet1.Rows(4).Cells(i).value = Nothing Then
                '    mtcObj.revisonlevel3 = "Null"
                'Else
                '    mtcObj.revisonlevel3 = xlsSheet1.Rows(4).Cells(i).value.ToString()
                'End If

                '6Is the Revision number correct? *
                'If xlsSheet1.Rows(7).Cells(i).value = Nothing Then
                '    mtcObj.revisoncorrect6 = "Null"
                'Else
                '    mtcObj.revisoncorrect6 = xlsSheet1.Rows(7).Cells(i).value.ToString()
                'End If


                '23component Name
                'mtcObj.componentName23 = er
            Next


            xlsWB.Close()
            xlsApp.Quit()
            releaseObject(xlsApp)
            releaseObject(xlsWB)
            releaseObject(xlsSheet1)
            killProcess()


        Catch ex As Exception
            'MsgBox($"Error while reading MTC_Assembly sheet {vbNewLine}{ex.Message}{vbNewLine} {ex.StackTrace}")
        End Try

    End Sub
    Public Sub MTC_Part(ByVal comObj As CommonReport)
        Try

            Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
            xlsWB = xlsApp.Workbooks.Open(comObj.excelFilePath, Nothing, False)

            Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing



            xlsSheet1 = xlsWB.Sheets("Part")
            xlsSheet1.Activate()
            Dim RowCnt = xlsSheet1.UsedRange.Rows.Count

            Dim ColCnt = xlsSheet1.UsedRange.Columns.Count


            xlsCell1 = xlsSheet1.UsedRange()

            Dim count = Nothing
            For i = 4 To ColCnt
                Dim row1 As DataRow = dt.NewRow()
                '0file
                comObj.fileName0 = "MTC"

                '1PartName Value(Error)
                comObj.partName1 = If(xlsSheet1.Rows(1).Cells(i).value = Nothing, er, xlsSheet1.Rows(1).Cells(i).value.ToString())



                '2Category
                comObj.category2 = xlsSheet1.Name

                '3author name/Who is the Author of the file?
                comObj.AuthorName3 = If(xlsSheet1.Rows(5).Cells(i).value = Nothing, er, xlsSheet1.Rows(5).Cells(i).value.ToString())

                '4projectname value/'  Is the correct "Project" used in the "Project Name" field? (Ensure consistent project titleIs used within a project) *
                'from assembly sheet first cell name(variable ProjectName)

                '5part number match value/Is the part number match with M2M? *
                comObj.partNumberMatch5 = If(xlsSheet1.Rows(3).Cells(i).value = Nothing, er, xlsSheet1.Rows(3).Cells(i).value.ToString())




                '6document number/Is the Document Number correct? (BEC designed parts must have proper doc number.Hardware And other exceptions are N/A) *
                comObj.documentNumber6 = If(xlsSheet1.Rows(9).Cells(i).value = Nothing, er, xlsSheet1.Rows(9).Cells(i).value.ToString())


                '7authorNameIs the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) 
                comObj.AuthorName7 = If(xlsSheet1.Rows(6).Cells(i).value = Nothing, er, xlsSheet1.Rows(6).Cells(i).value.ToString())

                '8dashinunusedfield/Do all technically unused properties have a "dash" populated? 
                comObj.DashInUnusedFiled8 = If(xlsSheet1.Rows(10).Cells(i).value = Nothing, er, xlsSheet1.Rows(10).Cells(i).value.ToString())


                '9 m2mdescription/Does the Model Title MATCH the M2M Item Master (INV) Description Field?
                comObj.m2mDiscriptionMatch9 = If(xlsSheet1.Rows(11).Cells(i).value = Nothing, er, xlsSheet1.Rows(11).Cells(i).value.ToString())


                '10uomsmatch/Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
                comObj.uomsMatch10 = If(xlsSheet1.Rows(12).Cells(i).value = Nothing, er, xlsSheet1.Rows(12).Cells(i).value.ToString)



                '11interface/Is interfernces found in assembly?
                comObj.interface11 = notexist


                '12interpartcopiesIs interfernces found in assembly?
                comObj.interpartcopies12 = notexist


                '13partcopiespart copies detected
                comObj.partcopies13 = notexist


                '14brokenfilebroken  file Path detected
                comObj.brokefile14 = notexist

                '15Adjustable/Is the part "Adjustable"? (part should NOT be adjustable) *
                comObj.adjustable15 = If(xlsSheet1.Rows(18).Cells(i).value = Nothing, er, xlsSheet1.Rows(17).Cells(i).value.ToString())



                '16Mat'l spec/Is the mat'l spec field populated? (indicated by RED in metadata audit file)
                comObj.Matl_spec16 = If(xlsSheet1.Rows(13).Cells(i).value = Nothing, er, xlsSheet1.Rows(13).Cells(i).value.ToString)


                '17MaterialUsed/Is the material used field populated? *
                comObj.MaterialUsed17 = If(xlsSheet1.Rows(14).Cells(i).value = Nothing, er, xlsSheet1.Rows(14).Cells(i).value.ToString)

                '18RemovedUnusedFeatures/Have any suppressed (unused) features been removed from the model Pathfinder? *
                comObj.RemovedUnusedFeatures18 = If(xlsSheet1.Rows(17).Cells(i).value = Nothing, er, xlsSheet1.Rows(17).Cells(i).value.ToString)

                '19ASTMminimum/Is the bend radius of the part equal to or above the ASTM minimum? *
                comObj.ASTMminimum19 = notexist

                '20FlatPattern/. Is the flat pattern turned ON? (Applicable only to parts needing dxf or cut length) *
                comObj.FlatPattern20 = notexist

                '21HoleToolUse/Did I use the hole tool for ALL holes requiring fasteners? (include clearance holes forhardware, tapped holes, And slots) *
                comObj.HoleToolUse21 = notexist



                '22component Name
                'mtcObj.componentName23 = er

                '22virtualThread/Virtual thread applied for Fasteners?
                comObj.VirtualThread22 = notexist

                '23defined feature
                comObj.defineFeatures23 = notexist

                '24Supperessed Feature
                comObj.suppressedFeature24 = notexist

                '25Vendorpartnumber
                comObj.vendorPartNumber25 = notexist

                '26Hardwareparts
                comObj.HardwareParts26 = notexist

                '27m2mSourceMarked
                comObj.m2mSourceMarked27 = notexist

                '28SEStatus
                comObj.SEstatus28 = notexist

                '29exist features
                comObj.ExistFeatures29 = notexist

                '30MatingParts
                comObj.MatingParts30 = notexist

                '31 environment
                comObj.Environment31 = notexist

                '32Geometry
                comObj.Geometry32 = notexist

                '33UpdateOnFileSave33
                'comObj.UpdateOnFileSave33 = notexist

                '34ConstrainedFeatures34
                comObj.ConstrainedFeatures34 = notexist

                '35AllFeaturesRemoved35
                comObj.AllFeaturesRemoved35 = notexist

                '36HardwarePartBox
                comObj.HardwarePartBox36 = notexist

                '-------------------------------------------------------------------------------
                'mannual Questions
                '37 RefPartOccurrenceProp/reference parts?(higher level, report To PL, physical properties, all Set To "NO") * (Use occurrence property macro to recorrect setting)
                comObj.C_RefPartOccurrenceProp40 = notexist

                '38 AssemblyFeatures/Are the assembly features minimized? *
                comObj.C_AssemblyFeatures41 = notexist

                'a+p
                '39ModelPreview/Is the Preview of the model saved in conformance to the GEE? (On button select,
                comObj.C_ModelPreview42 = If(xlsSheet1.Rows(23).Cells(i).value = Nothing, er, xlsSheet1.Rows(23).Cells(i).value.ToString)

                '40PartConstraint/Record number of parts that are NOT fully constrained. **Record "0" if all parts are constrained correctly. (All parts locked and constraints convey design intent - i.e incorrect to "ground" parts when actual intent is for mating surfaces) *
                comObj.C_PartConstraint43 = notexist

                '---------------------------------------------------------------------------------------------------------
                'mtc part

                '41Material_Spec_Field/Is the material spec field filled out with the proper designation? (Modification,Purchased, Assembly, etc.) *
                comObj.C_Material_Spec_Field44 = If(xlsSheet1.Rows(21).Cells(i).value = Nothing, er, xlsSheet1.Rows(21).Cells(i).value.ToString)

                '--------------------------------------------------------------------------------------------------------------------------------
                'mtc sheetmetal +part mass and density

                '42MassAndDensity_Update/Is the density defined and the mass up to date? * (macro will auto-update if file not read only)
                comObj.C_MassAndDensity_Update45 = If(xlsSheet1.Rows(22).Cells(i).value = Nothing, er, xlsSheet1.Rows(22).Cells(i).value.ToString)

                '43Bend_Radius_Update/Did I ensure the bend radius was producible? (Evaluate based on NPM0001C1 and/or consult ProdE/ManE) *
                comObj.C_Bend_Radius_Update46 = notexist

                '44 FlatPattern_Update/Is the flat pattern up to date? *
                comObj.C_FlatPattern_Update47 = notexist

                '--------------------------------------------------------------------------------------------------------------------------------
                'MTC Baseline
                '45 HardwareInstancesAndStackups/Are all instances of hardware consistent and have logical stackups? (Examples: 1.a SS nut applied to a zinc coated bolt having a copper washer; 2. Nine(9) instances of grade 8 and one(1) instance of grade 5 within a mounting configuration) *
                comObj.C_HardwareInstancesAndStackups48 = notexist

                '46 DimensionalGeometry_Match_For_ALL/Has the dimensional geometry of the model been verified to match the vendor supplied data? (This is to include use of the hole tool for ALL features requiring fasteners) *
                comObj.C_DimensionalGeometry_Match_For_ALL49 = notexist

                '47 DimensionalGeometry_Match/Has the dimensional geometry of the model been verified to match the vendor supplied data? *
                comObj.C_DimensionalGeometry_Match50 = notexist

                '48 suppressedUnusedFeatures/ Have any suppressed (unused) features been removed from the model Pathfinder? *
                comObj.C_suppressedUnusedFeatures51 = notexist

                '49 ChildComponentsFullyConstrained/Are all child components fully constrained? (Ground constraints should be minimized for BEC generated models, but are acceptable for Vendor supplied assemblies) *
                comObj.C_ChildComponentsFullyConstrained52 = notexist

                '50 SimplifiedAssemblyModel/Is the assembly model simplified in design? (unnecessary details or "internal" parts should be removed if possible) *
                comObj.C_SimplifiedAssemblyModel53 = notexist

                '51 ChildPartOccurrenceProperties/Have the required occurrence properties for ALL child components been set to NO? (Reports/Parts List, Physical Properties, and Interference Analysis should all be set to NO) *
                comObj.C_ChildPartOccurrenceProperties54 = notexist

                '52 VendorMaterialData/Is the appropriate material applied to the model per Vendor supplied data? (If the component contains multiple material PURCHASED should be utilized) *
                comObj.C_VendorMaterialData55 = notexist

                '53 BaselineMassAndDensityUpdate/Is the density defined and the mass up to date? ("Update on file save" should be unchecked. When a Vendor weight is known "User-defined properties" should be checked and the mass applied) * (Manual check -if read-only)
                comObj.C_BaselineMassAndDensityUpdate56 = notexist

                '54 PMI_Instruction/Is PMI included for any other further or special situations? (kit identification, missing vendor information, etc.) *
                comObj.C_PMI_Instruction57 = notexist

                '55 BinderData/Does the model have a Binder attached and is properly filled out? (A hyperlink to the Vendor supplied data sheet located on the VAULT should be included) *
                comObj.C_BinderData58 = notexist

                '56 TerminalAssigned/Have the required terminals been assigned to the part model? *
                comObj.C_TerminalAssigned59 = notexist

                '57 ReferencePlanes/ Has the model been constructed on the correct reference planes per GEE0008? *
                comObj.C_ReferencePlanes60 = notexist

                '------------------------------------------------------------------------------------------------------------------------------------------
                'MTR assembly + sheetmetal+part 

                'assembly+part+sheetmetal
                '58 ConstrainedRefPlane/Verify the part model has been created and constrained to the correct reference plane.
                comObj.R_ConstrainedRefPlane61 = notexist

                '59 R_FullyConstrained/Verify that ALL parts/sub-assemblies have been fully constrained
                comObj.R_FullyConstrained62 = notexist

                '60 R_FullyConstrainedAndGround/Verify that ALL parts/sub-assemblies have been properly constrained***The “Ground” constraint should only be applied to the first component***
                comObj.R_FullyConstrainedAndGround63 = notexist

                'a+p
                '61 R_ProperConstrained/Verify proper constraints by utilizing “Assembly Relationship Manager”
                comObj.R_ProperConstrained64 = notexist

                'a+p+s
                '62 R_NeedSimplified/Verify if the assembly can be or needs to be simplified
                comObj.R_NeedSimplified65 = notexist

                '63 R_PreviewModelGEE/ matching part, background,to turn off shadows, sketches, coordinate systems And PMI reference planes off)
                comObj.R_PreviewModelGEE66 = notexist

                '64 R_WeightMass/Verify that the weight and mass are correct
                comObj.R_WeightMass67 = notexist

                '65 R_OccurrenceProperties/Verify the occurrence properties are set properly for the assembly model,
                comObj.R_OccurrenceProperties68 = notexist

                '66 R_PatternsWithinPatterns/Verify that “patterns” do not exist within other patterns. If present, they should be removed
                comObj.R_PatternsWithinPatterns69 = notexist

                '67 R_ExternallySuppliedModels/Do externally supplied models exist within the assembly? Are the models a surface model? Are the models a solid model?
                comObj.R_ExternallySuppliedModels70 = notexist

                '68 R_holePatterns/Verify that ALL hole patterns align properly with mating parts/components
                comObj.R_holePatterns71 = notexist

                '69 R_HoleDiameters/Verify that hole diameters are correct for size application
                comObj.R_HoleDiameters72 = notexist

                '70 R_HoleClearances/Verify that hole clearances are correct for fasteners used
                comObj.R_HoleClearances73 = notexist

                '71 R_HoleDistance/Verify that hole distance from edge of material is appropriate
                comObj.R_HoleDistance74 = notexist

                '72 R_PartLineupWithMatingParts/Verify that the part will line up correctly with mating parts
                comObj.R_PartLineupWithMatingParts75 = notexist

                '73 R_FITS_In_MatingParts/Verify that “FITS” are correct to mating parts Use Engineering Document MEA0001
                comObj.R_FITS_In_MatingParts76 = notexist

                'part
                '74 R_Density/Verify that the correct density has been applied to the model
                comObj.R_Density77 = notexist

                'p+s
                '75 R_Material/Verify the correct material is assigned to the part for the application
                comObj.R_Material = notexist

                'sheetmetal
                '76 R_UnusedSketches/Verify that any unused sketches have been removed
                comObj.R_UnusedSketches78 = notexist

                '77 R_linkSketches/Verify that any links within a sketch have been removed
                comObj.R_linkSketches79 = notexist

                '78 R_BinderProcess/Verify the Binder Document Process (GEE0001) was followed
                comObj.R_BinderProcess80 = notexist

                '79 R_BendRadius/Verify the bend radius meets requirements of the material selected
                comObj.R_BendRadius81 = notexist



                '-------------------------------------------------------------------------------------

                '38 date
                comObj.date38 = Date.Now()


                '39 lastModifiedDate
                comObj.lastModifiedDate39 = If(xlsSheet1.Rows(20).Cells(i).value = Nothing, er, xlsSheet1.Rows(20).Cells(i).value.ToString())
                '------------------------------------------------------------------
                row1Data(comObj)

            Next



            xlsWB.Close()
            xlsApp.Quit()
            releaseObject(xlsApp)
            releaseObject(xlsWB)
            releaseObject(xlsSheet1)
            killProcess()


        Catch ex As Exception
            'MsgBox($"Error while reading MTC_Part sheet {vbNewLine}{ex.Message}{vbNewLine} {ex.StackTrace}")
        End Try


    End Sub

    Public Sub MTC_Sheetmetal(ByVal comObj As CommonReport)
        Try

            Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
            xlsWB = xlsApp.Workbooks.Open(comObj.excelFilePath, Nothing, False)

            Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing



            xlsSheet1 = xlsWB.Sheets("Sheetmetal")
            xlsSheet1.Activate()
            Dim RowCnt = xlsSheet1.UsedRange.Rows.Count

            Dim ColCnt = xlsSheet1.UsedRange.Columns.Count


            xlsCell1 = xlsSheet1.UsedRange()

            Dim count = Nothing
            For i = 4 To ColCnt
                '0file
                comObj.fileName0 = "MTC"

                '1PartName Value(Error)
                comObj.partName1 = If(xlsSheet1.Rows(1).Cells(i).value = Nothing, er, xlsSheet1.Rows(1).Cells(i).value.ToString())



                '2Category
                comObj.category2 = xlsSheet1.Name

                '3author name/Who is the Author of the file?
                comObj.AuthorName3 = If(xlsSheet1.Rows(5).Cells(i).value = Nothing, er, xlsSheet1.Rows(5).Cells(i).value.ToString())


                '4projectname value/'  Is the correct "Project" used in the "Project Name" field? (Ensure consistent project titleIs used within a project) *
                'from assembly sheet first cell name(variable ProjectName)



                '5part number match value/Is the part number match with M2M? *
                comObj.partNumberMatch5 = If(xlsSheet1.Rows(3).Cells(i).value = Nothing, er, xlsSheet1.Rows(3).Cells(i).value.ToString())


                '6document number/Is the Document Number correct? (BEC designed parts must have proper doc number.Hardware And other exceptions are N/A) *
                comObj.documentNumber6 = If(xlsSheet1.Rows(8).Cells(i).value = Nothing, er, xlsSheet1.Rows(9).Cells(i).value.ToString())


                '7authorNameIs the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) 
                comObj.AuthorName7 = If(xlsSheet1.Rows(9).Cells(i).value = Nothing, er, xlsSheet1.Rows(6).Cells(i).value.ToString())

                '8dashinunusedfield/Do all technically unused properties have a "dash" populated? 
                comObj.DashInUnusedFiled8 = If(xlsSheet1.Rows(10).Cells(i).value = Nothing, er, xlsSheet1.Rows(10).Cells(i).value.ToString())


                '9 m2mdescription/Does the Model Title MATCH the M2M Item Master (INV) Description Field?
                comObj.m2mDiscriptionMatch9 = If(xlsSheet1.Rows(11).Cells(i).value = Nothing, er, xlsSheet1.Rows(11).Cells(i).value.ToString())


                '10uomsmatch/Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
                comObj.uomsMatch10 = If(xlsSheet1.Rows(12).Cells(i).value = Nothing, er, xlsSheet1.Rows(12).Cells(i).value.ToString)



                '11interface/Is interfernces found in assembly?
                comObj.interface11 = notexist


                '12interpartcopiesIs interfernces found in assembly?
                comObj.interpartcopies12 = notexist


                '13partcopiespart copies detected
                comObj.partcopies13 = notexist


                '14brokenfilebroken  file Path detected
                comObj.brokefile14 = notexist

                '15Adjustable/Is the part "Adjustable"? (part should NOT be adjustable) *
                ' mtcObj.adjustable16 = "No"
                comObj.adjustable15 = If(xlsSheet1.Rows(18).Cells(i).value = Nothing, er, xlsSheet1.Rows(18).Cells(i).value.ToString())



                '16Mat'l spec/Is the mat'l spec field populated? (indicated by RED in metadata audit file)
                comObj.Matl_spec16 = If(xlsSheet1.Rows(13).Cells(i).value = Nothing, er, xlsSheet1.Rows(13).Cells(i).value.ToString)


                '17MaterialUsed/Is the material used field populated? *
                comObj.MaterialUsed17 = If(xlsSheet1.Rows(14).Cells(i).value = Nothing, er, xlsSheet1.Rows(14).Cells(i).value.ToString)

                '18RemovedUnusedFeatures/Have any suppressed (unused) features been removed from the model Pathfinder? *
                comObj.RemovedUnusedFeatures18 = notexist

                '19ASTMminimum/Is the bend radius of the part equal to or above the ASTM minimum? *
                comObj.ASTMminimum19 = If(xlsSheet1.Rows(15).Cells(i).value = Nothing, er, xlsSheet1.Rows(15).Cells(i).value.ToString)

                '20FlatPattern/. Is the flat pattern turned ON? (Applicable only to parts needing dxf or cut length) *
                comObj.FlatPattern20 = If(xlsSheet1.Rows(16).Cells(i).value = Nothing, er, xlsSheet1.Rows(16).Cells(i).value.ToString)

                '21HoleToolUse/Did I use the hole tool for ALL holes requiring fasteners? (include clearance holes forhardware, tapped holes, And slots) *
                comObj.HoleToolUse21 = If(xlsSheet1.Rows(17).Cells(i).value = Nothing, er, xlsSheet1.Rows(17).Cells(i).value.ToString)


                '21component Name
                'mtcObj.componentName23 = er

                '22virtualThread/Virtual thread applied for Fasteners?
                comObj.VirtualThread22 = notexist

                '23defined feature
                comObj.defineFeatures23 = notexist

                '24Supperessed Feature
                comObj.suppressedFeature24 = notexist

                '25Vendorpartnumber
                comObj.vendorPartNumber25 = notexist

                '26Hardwareparts
                comObj.HardwareParts26 = notexist

                '27m2mSourceMarked
                comObj.m2mSourceMarked27 = notexist

                '28SEStatus
                comObj.SEstatus28 = notexist

                '29exist features
                comObj.ExistFeatures29 = notexist

                '30MatingParts
                comObj.MatingParts30 = notexist

                '31 environment
                comObj.Environment31 = notexist

                '32Geometry
                comObj.Geometry32 = notexist

                '33UpdateOnFileSave33
                'comObj.UpdateOnFileSave33 = notexist

                '34ConstrainedFeatures34
                comObj.ConstrainedFeatures34 = notexist

                '35AllFeaturesRemoved35
                comObj.AllFeaturesRemoved35 = notexist

                '36HardwarePartBox
                comObj.HardwarePartBox36 = notexist

                '-------------------------------------------------------------------------------
                'mannual Questions
                '37 RefPartOccurrenceProp/reference parts?(higher level, report To PL, physical properties, all Set To "NO") * (Use occurrence property macro to recorrect setting)
                comObj.C_RefPartOccurrenceProp40 = notexist

                '38 AssemblyFeatures/Are the assembly features minimized? *
                comObj.C_AssemblyFeatures41 = notexist

                'a+p
                '39ModelPreview/Is the Preview of the model saved in conformance to the GEE? (On button select,
                comObj.C_ModelPreview42 = If(xlsSheet1.Rows(25).Cells(i).value = Nothing, er, xlsSheet1.Rows(25).Cells(i).value.ToString)

                '40PartConstraint/Record number of parts that are NOT fully constrained. **Record "0" if all parts are constrained correctly. (All parts locked and constraints convey design intent - i.e incorrect to "ground" parts when actual intent is for mating surfaces) *
                comObj.C_PartConstraint43 = notexist

                '---------------------------------------------------------------------------------------------------------
                'mtc part

                'p+sheetmetal
                '41Material_Spec_Field/Is the material spec field filled out with the proper designation? (Modification,Purchased, Assembly, etc.) *
                comObj.C_Material_Spec_Field44 = If(xlsSheet1.Rows(22).Cells(i).value = Nothing, er, xlsSheet1.Rows(22).Cells(i).value.ToString)

                '--------------------------------------------------------------------------------------------------------------------------------
                'mtc sheetmetal +part mass and density

                '42MassAndDensity_Update/Is the density defined and the mass up to date? * (macro will auto-update if file not read only)
                comObj.C_MassAndDensity_Update45 = If(xlsSheet1.Rows(21).Cells(i).value = Nothing, er, xlsSheet1.Rows(21).Cells(i).value.ToString)

                '43Bend_Radius_Update/Did I ensure the bend radius was producible? (Evaluate based on NPM0001C1 and/or consult ProdE/ManE) *
                comObj.C_Bend_Radius_Update46 = If(xlsSheet1.Rows(23).Cells(i).value = Nothing, er, xlsSheet1.Rows(23).Cells(i).value.ToString)

                '44 FlatPattern_Update/Is the flat pattern up to date? *
                comObj.C_FlatPattern_Update47 = If(xlsSheet1.Rows(24).Cells(i).value = Nothing, er, xlsSheet1.Rows(24).Cells(i).value.ToString)

                '--------------------------------------------------------------------------------------------------------------------------------
                'MTC Baseline
                '45 HardwareInstancesAndStackups/Are all instances of hardware consistent and have logical stackups? (Examples: 1.a SS nut applied to a zinc coated bolt having a copper washer; 2. Nine(9) instances of grade 8 and one(1) instance of grade 5 within a mounting configuration) *
                comObj.C_HardwareInstancesAndStackups48 = notexist

                '46 DimensionalGeometry_Match_For_ALL/Has the dimensional geometry of the model been verified to match the vendor supplied data? (This is to include use of the hole tool for ALL features requiring fasteners) *
                comObj.C_DimensionalGeometry_Match_For_ALL49 = notexist

                '47 DimensionalGeometry_Match/Has the dimensional geometry of the model been verified to match the vendor supplied data? *
                comObj.C_DimensionalGeometry_Match50 = notexist

                '48 suppressedUnusedFeatures/ Have any suppressed (unused) features been removed from the model Pathfinder? *
                comObj.C_suppressedUnusedFeatures51 = notexist

                '49 ChildComponentsFullyConstrained/Are all child components fully constrained? (Ground constraints should be minimized for BEC generated models, but are acceptable for Vendor supplied assemblies) *
                comObj.C_ChildComponentsFullyConstrained52 = notexist

                '50 SimplifiedAssemblyModel/Is the assembly model simplified in design? (unnecessary details or "internal" parts should be removed if possible) *
                comObj.C_SimplifiedAssemblyModel53 = notexist

                '51 ChildPartOccurrenceProperties/Have the required occurrence properties for ALL child components been set to NO? (Reports/Parts List, Physical Properties, and Interference Analysis should all be set to NO) *
                comObj.C_ChildPartOccurrenceProperties54 = notexist

                '52 VendorMaterialData/Is the appropriate material applied to the model per Vendor supplied data? (If the component contains multiple material PURCHASED should be utilized) *
                comObj.C_VendorMaterialData55 = notexist

                '53 BaselineMassAndDensityUpdate/Is the density defined and the mass up to date? ("Update on file save" should be unchecked. When a Vendor weight is known "User-defined properties" should be checked and the mass applied) * (Manual check -if read-only)
                comObj.C_BaselineMassAndDensityUpdate56 = notexist

                '54 PMI_Instruction/Is PMI included for any other further or special situations? (kit identification, missing vendor information, etc.) *
                comObj.C_PMI_Instruction57 = notexist

                '55 BinderData/Does the model have a Binder attached and is properly filled out? (A hyperlink to the Vendor supplied data sheet located on the VAULT should be included) *
                comObj.C_BinderData58 = notexist

                '56 TerminalAssigned/Have the required terminals been assigned to the part model? *
                comObj.C_TerminalAssigned59 = notexist

                '57 ReferencePlanes/ Has the model been constructed on the correct reference planes per GEE0008? *
                comObj.C_ReferencePlanes60 = notexist

                '------------------------------------------------------------------------------------------------------------------------------------------
                'MTR assembly + sheetmetal+part 

                'assembly+part+sheetmetal
                '58 ConstrainedRefPlane/Verify the part model has been created and constrained to the correct reference plane.
                comObj.R_ConstrainedRefPlane61 = notexist

                '59 R_FullyConstrained/Verify that ALL parts/sub-assemblies have been fully constrained
                comObj.R_FullyConstrained62 = notexist

                '60 R_FullyConstrainedAndGround/Verify that ALL parts/sub-assemblies have been properly constrained***The “Ground” constraint should only be applied to the first component***
                comObj.R_FullyConstrainedAndGround63 = notexist

                'a+p
                '61 R_ProperConstrained/Verify proper constraints by utilizing “Assembly Relationship Manager”
                comObj.R_ProperConstrained64 = notexist

                'a+p+s
                '62 R_NeedSimplified/Verify if the assembly can be or needs to be simplified
                comObj.R_NeedSimplified65 = notexist

                '63 R_PreviewModelGEE/ matching part, background,to turn off shadows, sketches, coordinate systems And PMI reference planes off)
                comObj.R_PreviewModelGEE66 = notexist

                '64 R_WeightMass/Verify that the weight and mass are correct
                comObj.R_WeightMass67 = notexist

                '65 R_OccurrenceProperties/Verify the occurrence properties are set properly for the assembly model,
                comObj.R_OccurrenceProperties68 = notexist

                '66 R_PatternsWithinPatterns/Verify that “patterns” do not exist within other patterns. If present, they should be removed
                comObj.R_PatternsWithinPatterns69 = notexist

                '67 R_ExternallySuppliedModels/Do externally supplied models exist within the assembly? Are the models a surface model? Are the models a solid model?
                comObj.R_ExternallySuppliedModels70 = notexist

                '68 R_holePatterns/Verify that ALL hole patterns align properly with mating parts/components
                comObj.R_holePatterns71 = notexist

                '69 R_HoleDiameters/Verify that hole diameters are correct for size application
                comObj.R_HoleDiameters72 = notexist

                '70 R_HoleClearances/Verify that hole clearances are correct for fasteners used
                comObj.R_HoleClearances73 = notexist

                '71 R_HoleDistance/Verify that hole distance from edge of material is appropriate
                comObj.R_HoleDistance74 = notexist

                '72 R_PartLineupWithMatingParts/Verify that the part will line up correctly with mating parts
                comObj.R_PartLineupWithMatingParts75 = notexist

                '73 R_FITS_In_MatingParts/Verify that “FITS” are correct to mating parts Use Engineering Document MEA0001
                comObj.R_FITS_In_MatingParts76 = notexist

                'part
                '74 R_Density/Verify that the correct density has been applied to the model
                comObj.R_Density77 = notexist

                'p+s
                '75 R_Material/Verify the correct material is assigned to the part for the application
                comObj.R_Material = notexist

                'sheetmetal
                '76 R_UnusedSketches/Verify that any unused sketches have been removed
                comObj.R_UnusedSketches78 = notexist

                '77 R_linkSketches/Verify that any links within a sketch have been removed
                comObj.R_linkSketches79 = notexist

                '78 R_BinderProcess/Verify the Binder Document Process (GEE0001) was followed
                comObj.R_BinderProcess80 = notexist

                '79 R_BendRadius/Verify the bend radius meets requirements of the material selected
                comObj.R_BendRadius81 = notexist



                '-------------------------------------------------------------------------------------


                '38 date
                comObj.date38 = Date.Now()
                '39 lastModifiedDate
                comObj.lastModifiedDate39 = If(xlsSheet1.Rows(20).Cells(i).value = Nothing, er, xlsSheet1.Rows(20).Cells(i).value.ToString())
                '----------------------------------------------------------------------------------------------------------

                row1Data(comObj)



            Next


            xlsWB.Close()
            xlsApp.Quit()
            releaseObject(xlsApp)
            releaseObject(xlsWB)
            releaseObject(xlsSheet1)
            killProcess()

        Catch ex As Exception
            'MsgBox($"Error while reading MTC_SheetMetal sheet {vbNewLine}{ex.Message}{vbNewLine} {ex.StackTrace}")
        End Try


    End Sub

    Public Sub MTC_baseline(ByVal comObj As CommonReport)

        Try

            Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
            xlsWB = xlsApp.Workbooks.Open(comObj.excelFilePath, Nothing, False)

            Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing



            xlsSheet1 = xlsWB.Sheets("Baseline")
            xlsSheet1.Activate()
            Dim RowCnt = xlsSheet1.UsedRange.Rows.Count

            Dim ColCnt = xlsSheet1.UsedRange.Columns.Count


            xlsCell1 = xlsSheet1.UsedRange()

            Dim count = Nothing
            For i = 4 To ColCnt

                '0filename
                comObj.fileName0 = "MTC"

                '1PartName Value(Error)
                comObj.partName1 = If(xlsSheet1.Rows(1).Cells(i).value = Nothing, er, xlsSheet1.Rows(1).Cells(i).value.ToString())


                '2Category
                comObj.category2 = xlsSheet1.Name


                '3author name/Who is the Author of the file?
                comObj.AuthorName3 = If(xlsSheet1.Rows(4).Cells(i).value = Nothing, er, xlsSheet1.Rows(4).Cells(i).value.ToString())



                '4projectname value/'  Is the correct "Project" used in the "Project Name" field? (Ensure consistent project titleIs used within a project) *
                'from assembly sheet first cell name(variable ProjectName)

                '5part number match value/Is the part number match with M2M? *
                comObj.partNumberMatch5 = If(xlsSheet1.Rows(2).Cells(i).value = Nothing, er, xlsSheet1.Rows(2).Cells(i).value.ToString())



                '6document number/Is the "Document Number" field populated with the correct part number? (This should MATCH the M2M Item Master Part Number field) *
                comObj.documentNumber6 = If(xlsSheet1.Rows(15).Cells(i).value = Nothing, er, xlsSheet1.Rows(15).Cells(i).value.ToString())


                '7authorNameIs the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) 
                comObj.AuthorName7 = If(xlsSheet1.Rows(12).Cells(i).value = Nothing, er, xlsSheet1.Rows(12).Cells(i).value.ToString())


                '8dashinunusedfield/Do all technically unused properties have a "dash" populated? 
                comObj.DashInUnusedFiled8 = If(xlsSheet1.Rows(19).Cells(i).value = Nothing, er, xlsSheet1.Rows(19).Cells(i).value.ToString())


                '9 m2mdescription/Does the Model Title MATCH the M2M Item Master (INV) Description Field?
                comObj.m2mDiscriptionMatch9 = If(xlsSheet1.Rows(11).Cells(i).value = Nothing, er, xlsSheet1.Rows(11).Cells(i).value.ToString())


                '10uomsmatch/Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
                comObj.uomsMatch10 = If(xlsSheet1.Rows(20).Cells(i).value = Nothing, er, xlsSheet1.Rows(20).Cells(i).value.ToString)



                '11interface/Is interfernces found in assembly?
                comObj.interface11 = notexist


                '12interpartcopiesIs interfernces found in assembly?
                comObj.interpartcopies12 = If(xlsSheet1.Rows(24).Cells(i).value = Nothing, er, xlsSheet1.Rows(24).Cells(i).value.ToString)


                '13partcopiespart copies detected
                comObj.partcopies13 = If(xlsSheet1.Rows(25).Cells(i).value = Nothing, er, xlsSheet1.Rows(25).Cells(i).value.ToString)


                '14brokenfilebroken  file Path detected
                comObj.brokefile14 = If(xlsSheet1.Rows(26).Cells(i).value = Nothing, er, xlsSheet1.Rows(26).Cells(i).value.ToString)


                '15Adjustable/Is the part "Adjustable"? (part should NOT be adjustable) *
                comObj.adjustable15 = If(xlsSheet1.Rows(27).Cells(i).value = Nothing, er, xlsSheet1.Rows(27).Cells(i).value.ToString())



                '16Mat'l spec/Is the "Mat'l Spec" field populated? (PURCHASED for library components) *
                comObj.Matl_spec16 = If(xlsSheet1.Rows(9).Cells(i).value = Nothing, er, xlsSheet1.Rows(9).Cells(i).value.ToString)


                '17MaterialUsed/Is the "Material Used" field populated? (PURCHASED for library components) *
                comObj.MaterialUsed17 = If(xlsSheet1.Rows(10).Cells(i).value = Nothing, er, xlsSheet1.Rows(10).Cells(i).value.ToString)

                '18RemovedUnusedFeatures/Have any suppressed (unused) features been removed from the model Pathfinder? *
                comObj.RemovedUnusedFeatures18 = If(xlsSheet1.Rows(35).Cells(i).value = Nothing, er, xlsSheet1.Rows(35).Cells(i).value.ToString)

                '19ASTMminimum/Is the bend radius of the part equal to or above the ASTM minimum? *
                comObj.ASTMminimum19 = notexist

                '20FlatPattern/. Is the flat pattern turned ON? (Applicable only to parts needing dxf or cut length) *
                comObj.FlatPattern20 = notexist

                '21HoleToolUse/Did I use the hole tool for ALL holes requiring fasteners? (include clearance holes forhardware, tapped holes, And slots) *
                comObj.HoleToolUse21 = If(xlsSheet1.Rows(22).Cells(i).value = Nothing, er, xlsSheet1.Rows(22).Cells(i).value.ToString)


                '22virtualThread/Virtual thread applied for Fasteners?
                comObj.VirtualThread22 = If(xlsSheet1.Rows(6).Cells(i).value = Nothing, er, xlsSheet1.Rows(6).Cells(i).value.ToString)

                '23defined feature/Are ALL features fully defined in the model Pathfinder? * (Sketch fully constraint?)
                comObj.defineFeatures23 = If(xlsSheet1.Rows(7).Cells(i).value = Nothing, er, xlsSheet1.Rows(7).Cells(i).value.ToString)

                '24Supperessed Feature
                comObj.suppressedFeature24 = If(xlsSheet1.Rows(8).Cells(i).value = Nothing, er, xlsSheet1.Rows(8).Cells(i).value.ToString)

                '25Vendorpartnumber
                comObj.vendorPartNumber25 = If(xlsSheet1.Rows(14).Cells(i).value = Nothing, er, xlsSheet1.Rows(14).Cells(i).value.ToString)

                '26Hardwareparts
                comObj.HardwareParts26 = If(xlsSheet1.Rows(18).Cells(i).value = Nothing, er, xlsSheet1.Rows(18).Cells(i).value.ToString)

                '27m2mSourceMarked
                comObj.m2mSourceMarked27 = If(xlsSheet1.Rows(21).Cells(i).value = Nothing, er, xlsSheet1.Rows(21).Cells(i).value.ToString)

                '28SEStatus
                comObj.SEstatus28 = If(xlsSheet1.Rows(29).Cells(i).value = Nothing, er, xlsSheet1.Rows(29).Cells(i).value.ToString)

                '29exist features
                comObj.ExistFeatures29 = notexist

                '30MatingParts
                comObj.MatingParts30 = notexist

                '31 environment
                comObj.Environment31 = notexist

                '32Geometry
                comObj.Geometry32 = notexist

                '33UpdateOnFileSave33
                ' comObj.UpdateOnFileSave33 = notexist

                '34ConstrainedFeatures34
                comObj.ConstrainedFeatures34 = notexist

                '35AllFeaturesRemoved35
                comObj.AllFeaturesRemoved35 = notexist

                '36HardwarePartBox
                comObj.HardwarePartBox36 = notexist

                '-------------------------------------------------------------------------------
                'mannual Questions
                '37 RefPartOccurrenceProp/reference parts?(higher level, report To PL, physical properties, all Set To "NO") * (Use occurrence property macro to recorrect setting)
                comObj.C_RefPartOccurrenceProp40 = notexist

                '38 AssemblyFeatures/Are the assembly features minimized? *
                comObj.C_AssemblyFeatures41 = notexist

                'a+p
                '39ModelPreview/Is the Preview of the model saved in conformance to the GEE? (On button select,
                comObj.C_ModelPreview42 = If(xlsSheet1.Rows(42).Cells(i).value = Nothing, er, xlsSheet1.Rows(42).Cells(i).value.ToString)

                '40PartConstraint/Record number of parts that are NOT fully constrained. **Record "0" if all parts are constrained correctly. (All parts locked and constraints convey design intent - i.e incorrect to "ground" parts when actual intent is for mating surfaces) *
                comObj.C_PartConstraint43 = notexist

                '---------------------------------------------------------------------------------------------------------
                'mtc part

                'p+sheetmetal
                '41Material_Spec_Field/Is the material spec field filled out with the proper designation? (Modification,Purchased, Assembly, etc.) *
                comObj.C_Material_Spec_Field44 = notexist

                '--------------------------------------------------------------------------------------------------------------------------------
                'mtc sheetmetal +part mass and density

                '42MassAndDensity_Update/Is the density defined and the mass up to date? * (macro will auto-update if file not read only)
                comObj.C_MassAndDensity_Update45 = notexist

                '43Bend_Radius_Update/Did I ensure the bend radius was producible? (Evaluate based on NPM0001C1 and/or consult ProdE/ManE) *
                comObj.C_Bend_Radius_Update46 = notexist

                '44 FlatPattern_Update/Is the flat pattern up to date? *
                comObj.C_FlatPattern_Update47 = notexist

                '--------------------------------------------------------------------------------------------------------------------------------
                'MTC Baseline
                '45 HardwareInstancesAndStackups/Are all instances of hardware consistent and have logical stackups? (Examples: 1.a SS nut applied to a zinc coated bolt having a copper washer; 2. Nine(9) instances of grade 8 and one(1) instance of grade 5 within a mounting configuration) *
                comObj.C_HardwareInstancesAndStackups48 = If(xlsSheet1.Rows(31).Cells(i).value = Nothing, er, xlsSheet1.Rows(31).Cells(i).value.ToString)

                '46 DimensionalGeometry_Match_For_ALL/Has the dimensional geometry of the model been verified to match the vendor supplied data? (This is to include use of the hole tool for ALL features requiring fasteners) *
                comObj.C_DimensionalGeometry_Match_For_ALL49 = If(xlsSheet1.Rows(34).Cells(i).value = Nothing, er, xlsSheet1.Rows(34).Cells(i).value.ToString)

                '47 DimensionalGeometry_Match/Has the dimensional geometry of the model been verified to match the vendor supplied data? *
                comObj.C_DimensionalGeometry_Match50 = If(xlsSheet1.Rows(35).Cells(i).value = Nothing, er, xlsSheet1.Rows(35).Cells(i).value.ToString)

                '48 suppressedUnusedFeatures/ Have any suppressed (unused) features been removed from the model Pathfinder? *
                comObj.C_suppressedUnusedFeatures51 = If(xlsSheet1.Rows(36).Cells(i).value = Nothing, er, xlsSheet1.Rows(36).Cells(i).value.ToString)

                '49 ChildComponentsFullyConstrained/Are all child components fully constrained? (Ground constraints should be minimized for BEC generated models, but are acceptable for Vendor supplied assemblies) *
                comObj.C_ChildComponentsFullyConstrained52 = If(xlsSheet1.Rows(37).Cells(i).value = Nothing, er, xlsSheet1.Rows(37).Cells(i).value.ToString)

                '50 SimplifiedAssemblyModel/Is the assembly model simplified in design? (unnecessary details or "internal" parts should be removed if possible) *
                comObj.C_SimplifiedAssemblyModel53 = If(xlsSheet1.Rows(38).Cells(i).value = Nothing, er, xlsSheet1.Rows(38).Cells(i).value.ToString)

                '51 ChildPartOccurrenceProperties/Have the required occurrence properties for ALL child components been set to NO? (Reports/Parts List, Physical Properties, and Interference Analysis should all be set to NO) *
                comObj.C_ChildPartOccurrenceProperties54 = If(xlsSheet1.Rows(39).Cells(i).value = Nothing, er, xlsSheet1.Rows(39).Cells(i).value.ToString)

                '52 VendorMaterialData/Is the appropriate material applied to the model per Vendor supplied data? (If the component contains multiple material PURCHASED should be utilized) *
                comObj.C_VendorMaterialData55 = If(xlsSheet1.Rows(40).Cells(i).value = Nothing, er, xlsSheet1.Rows(40).Cells(i).value.ToString)

                '53 BaselineMassAndDensityUpdate/Is the density defined and the mass up to date? ("Update on file save" should be unchecked. When a Vendor weight is known "User-defined properties" should be checked and the mass applied) * (Manual check -if read-only)
                comObj.C_BaselineMassAndDensityUpdate56 = If(xlsSheet1.Rows(41).Cells(i).value = Nothing, er, xlsSheet1.Rows(41).Cells(i).value.ToString)

                '54 PMI_Instruction/Is PMI included for any other further or special situations? (kit identification, missing vendor information, etc.) *
                comObj.C_PMI_Instruction57 = If(xlsSheet1.Rows(43).Cells(i).value = Nothing, er, xlsSheet1.Rows(43).Cells(i).value.ToString)

                '55 BinderData/Does the model have a Binder attached and is properly filled out? (A hyperlink to the Vendor supplied data sheet located on the VAULT should be included) *
                comObj.C_BinderData58 = If(xlsSheet1.Rows(44).Cells(i).value = Nothing, er, xlsSheet1.Rows(44).Cells(i).value.ToString)

                '56 TerminalAssigned/Have the required terminals been assigned to the part model? *
                comObj.C_TerminalAssigned59 = If(xlsSheet1.Rows(45).Cells(i).value = Nothing, er, xlsSheet1.Rows(45).Cells(i).value.ToString)

                '57 ReferencePlanes/ Has the model been constructed on the correct reference planes per GEE0008? *
                comObj.C_ReferencePlanes60 = If(xlsSheet1.Rows(46).Cells(i).value = Nothing, er, xlsSheet1.Rows(46).Cells(i).value.ToString)

                '------------------------------------------------------------------------------------------------------------------------------------------
                'MTR assembly + sheetmetal+part 

                'assembly+part+sheetmetal
                '58 ConstrainedRefPlane/Verify the part model has been created and constrained to the correct reference plane.
                comObj.R_ConstrainedRefPlane61 = notexist

                '59 R_FullyConstrained/Verify that ALL parts/sub-assemblies have been fully constrained
                comObj.R_FullyConstrained62 = notexist

                '60 R_FullyConstrainedAndGround/Verify that ALL parts/sub-assemblies have been properly constrained***The “Ground” constraint should only be applied to the first component***
                comObj.R_FullyConstrainedAndGround63 = notexist

                'a+p
                '61 R_ProperConstrained/Verify proper constraints by utilizing “Assembly Relationship Manager”
                comObj.R_ProperConstrained64 = notexist

                'a+p+s
                '62 R_NeedSimplified/Verify if the assembly can be or needs to be simplified
                comObj.R_NeedSimplified65 = notexist

                '63 R_PreviewModelGEE/ matching part, background,to turn off shadows, sketches, coordinate systems And PMI reference planes off)
                comObj.R_PreviewModelGEE66 = notexist

                '64 R_WeightMass/Verify that the weight and mass are correct
                comObj.R_WeightMass67 = notexist

                '65 R_OccurrenceProperties/Verify the occurrence properties are set properly for the assembly model,
                comObj.R_OccurrenceProperties68 = notexist

                '66 R_PatternsWithinPatterns/Verify that “patterns” do not exist within other patterns. If present, they should be removed
                comObj.R_PatternsWithinPatterns69 = notexist

                '67 R_ExternallySuppliedModels/Do externally supplied models exist within the assembly? Are the models a surface model? Are the models a solid model?
                comObj.R_ExternallySuppliedModels70 = notexist

                '68 R_holePatterns/Verify that ALL hole patterns align properly with mating parts/components
                comObj.R_holePatterns71 = notexist

                '69 R_HoleDiameters/Verify that hole diameters are correct for size application
                comObj.R_HoleDiameters72 = notexist

                '70 R_HoleClearances/Verify that hole clearances are correct for fasteners used
                comObj.R_HoleClearances73 = notexist

                '71 R_HoleDistance/Verify that hole distance from edge of material is appropriate
                comObj.R_HoleDistance74 = notexist

                '72 R_PartLineupWithMatingParts/Verify that the part will line up correctly with mating parts
                comObj.R_PartLineupWithMatingParts75 = notexist

                '73 R_FITS_In_MatingParts/Verify that “FITS” are correct to mating parts Use Engineering Document MEA0001
                comObj.R_FITS_In_MatingParts76 = notexist

                'part
                '74 R_Density/Verify that the correct density has been applied to the model
                comObj.R_Density77 = notexist

                'p+s
                '75 R_Material/Verify the correct material is assigned to the part for the application
                comObj.R_Material = notexist

                'sheetmetal
                '76 R_UnusedSketches/Verify that any unused sketches have been removed
                comObj.R_UnusedSketches78 = notexist

                '77 R_linkSketches/Verify that any links within a sketch have been removed
                comObj.R_linkSketches79 = notexist

                '78 R_BinderProcess/Verify the Binder Document Process (GEE0001) was followed
                comObj.R_BinderProcess80 = notexist

                '79 R_BendRadius/Verify the bend radius meets requirements of the material selected
                comObj.R_BendRadius81 = notexist



                '-------------------------------------------------------------------------------------


                '38 date
                comObj.date38 = Date.Now()
                '39 lastModifiedDate
                comObj.lastModifiedDate39 = If(xlsSheet1.Rows(30).Cells(i).value = Nothing, er, xlsSheet1.Rows(30).Cells(i).value.ToString())
                '----------------------------------------------------------------------------------------------------------

                row1Data(comObj)



            Next
            '23component Name/What type of component?
            'mtcObj.componentName23 = If(xlsSheet1.Rows(5).Cells(i).value = Nothing, er, xlsSheet1.Rows(5).Cells(i).value.ToString)

            xlsWB.Close()
            xlsApp.Quit()
            releaseObject(xlsApp)
            releaseObject(xlsWB)
            releaseObject(xlsSheet1)
            killProcess()

        Catch ex As Exception
            'MsgBox($"Error while reading MTC_Baseline sheet{vbNewLine} {ex.Message}{vbNewLine} {ex.StackTrace}")
        End Try

    End Sub

    Public Sub MTC_Electrical(ByVal comObj As CommonReport)
        Try
            Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
            xlsWB = xlsApp.Workbooks.Open(comObj.excelFilePath, Nothing, False)

            Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing



            xlsSheet1 = xlsWB.Sheets("Electrical")
            xlsSheet1.Activate()
            Dim RowCnt = xlsSheet1.UsedRange.Rows.Count

            Dim ColCnt = xlsSheet1.UsedRange.Columns.Count


            xlsCell1 = xlsSheet1.UsedRange()

            Dim count = Nothing
            For i = 4 To ColCnt

                '0filename
                comObj.fileName0 = "MTC"

                '1PartName Value(Error)
                comObj.partName1 = If(xlsSheet1.Rows(1).Cells(i).value = Nothing, er, xlsSheet1.Rows(1).Cells(i).value.ToString())

                '2Category
                comObj.category2 = xlsSheet1.Name


                '3author name/Who is the Author of the file?
                comObj.AuthorName3 = If(xlsSheet1.Rows(5).Cells(i).value = Nothing, er, xlsSheet1.Rows(5).Cells(i).value.ToString())

                '4projectname value/'  Is the correct "Project" used in the "Project Name" field? (Ensure consistent project titleIs used within a project) *
                'from assembly sheet first cell name(variable ProjectName)

                '5part number match value/Is the part number match with M2M? *
                comObj.partNumberMatch5 = If(xlsSheet1.Rows(3).Cells(i).value = Nothing, er, xlsSheet1.Rows(3).Cells(i).value.ToString())



                '6document number/Is the "Document Number" field populated with the correct part number? (This should MATCH the M2M Item Master Part Number field) *
                comObj.documentNumber6 = If(xlsSheet1.Rows(8).Cells(i).value = Nothing, er, xlsSheet1.Rows(8).Cells(i).value.ToString())


                '7authorNameIs the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) 
                comObj.AuthorName7 = If(xlsSheet1.Rows(9).Cells(i).value = Nothing, er, xlsSheet1.Rows(9).Cells(i).value.ToString())


                '8dashinunusedfield/Do all technically unused properties have a "dash" populated? 
                comObj.DashInUnusedFiled8 = If(xlsSheet1.Rows(10).Cells(i).value = Nothing, er, xlsSheet1.Rows(10).Cells(i).value.ToString())


                '9 m2mdescription/Does the Model Title MATCH the M2M Item Master (INV) Description Field?
                comObj.m2mDiscriptionMatch9 = If(xlsSheet1.Rows(11).Cells(i).value = Nothing, er, xlsSheet1.Rows(11).Cells(i).value.ToString())


                '10uomsmatch/Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
                comObj.uomsMatch10 = notexist


                '11interface/Is interfernces found in assembly?
                comObj.interface11 = notexist


                '12interpartcopiesIs interfernces found in assembly?
                comObj.interpartcopies12 = notexist


                '13partcopiespart copies detected
                comObj.partcopies13 = notexist


                '14brokenfilebroken  file Path detected
                comObj.brokefile14 = notexist

                '15Adjustable/Is the part "Adjustable"? (part should NOT be adjustable) *
                comObj.adjustable15 = notexist

                '16Mat'l spec/Is the "Mat'l Spec" field populated? (PURCHASED for library components) *
                comObj.Matl_spec16 = notexist


                '17MaterialUsed/Is the "Material Used" field populated? (PURCHASED for library components) *
                comObj.MaterialUsed17 = notexist

                '18RemovedUnusedFeatures/Have any suppressed (unused) features been removed from the model Pathfinder? *
                comObj.RemovedUnusedFeatures18 = notexist

                '19ASTMminimum/Is the bend radius of the part equal to or above the ASTM minimum? *
                comObj.ASTMminimum19 = notexist

                '20FlatPattern/. Is the flat pattern turned ON? (Applicable only to parts needing dxf or cut length) *
                comObj.FlatPattern20 = notexist

                '21HoleToolUse/Did I use the hole tool for ALL holes requiring fasteners? (include clearance holes forhardware, tapped holes, And slots) *
                comObj.HoleToolUse21 = notexist


                '22virtualThread/Virtual thread applied for Fasteners?
                comObj.VirtualThread22 = notexist

                '23defined feature/Are ALL features fully defined in the model Pathfinder? * (Sketch fully constraint?)
                comObj.defineFeatures23 = notexist

                '24Supperessed Feature
                comObj.suppressedFeature24 = notexist

                '25Vendorpartnumber
                comObj.vendorPartNumber25 = notexist

                '26Hardwareparts
                comObj.HardwareParts26 = notexist

                '27m2mSourceMarked
                comObj.m2mSourceMarked27 = notexist

                '28SEStatus
                comObj.SEstatus28 = notexist

                '29exist features
                comObj.ExistFeatures29 = notexist

                '30MatingParts
                comObj.MatingParts30 = notexist

                '31 environment
                comObj.Environment31 = notexist

                '32Geometry
                comObj.Geometry32 = notexist

                '33UpdateOnFileSave33
                'comObj.UpdateOnFileSave33 = notexist

                '34ConstrainedFeatures34
                comObj.ConstrainedFeatures34 = notexist

                '35AllFeaturesRemoved35
                comObj.AllFeaturesRemoved35 = notexist

                '36HardwarePartBox
                comObj.HardwarePartBox36 = notexist


                '-------------------------------------------------------------------------------
                'mannual Questions
                '37 RefPartOccurrenceProp/reference parts?(higher level, report To PL, physical properties, all Set To "NO") * (Use occurrence property macro to recorrect setting)
                comObj.C_RefPartOccurrenceProp40 = notexist

                '38 AssemblyFeatures/Are the assembly features minimized? *
                comObj.C_AssemblyFeatures41 = notexist

                'a+p
                '39ModelPreview/Is the Preview of the model saved in conformance to the GEE? (On button select,
                comObj.C_ModelPreview42 = notexist

                '40PartConstraint/Record number of parts that are NOT fully constrained. **Record "0" if all parts are constrained correctly. (All parts locked and constraints convey design intent - i.e incorrect to "ground" parts when actual intent is for mating surfaces) *
                comObj.C_PartConstraint43 = notexist

                '---------------------------------------------------------------------------------------------------------
                'mtc part

                'p+sheetmetal
                '41Material_Spec_Field/Is the material spec field filled out with the proper designation? (Modification,Purchased, Assembly, etc.) *
                comObj.C_Material_Spec_Field44 = notexist

                '--------------------------------------------------------------------------------------------------------------------------------
                'mtc sheetmetal +part mass and density

                '42MassAndDensity_Update/Is the density defined and the mass up to date? * (macro will auto-update if file not read only)
                comObj.C_MassAndDensity_Update45 = notexist

                '43Bend_Radius_Update/Did I ensure the bend radius was producible? (Evaluate based on NPM0001C1 and/or consult ProdE/ManE) *
                comObj.C_Bend_Radius_Update46 = notexist

                '44 FlatPattern_Update/Is the flat pattern up to date? *
                comObj.C_FlatPattern_Update47 = notexist

                '--------------------------------------------------------------------------------------------------------------------------------
                'MTC Baseline
                '45 HardwareInstancesAndStackups/Are all instances of hardware consistent and have logical stackups? (Examples: 1.a SS nut applied to a zinc coated bolt having a copper washer; 2. Nine(9) instances of grade 8 and one(1) instance of grade 5 within a mounting configuration) *
                comObj.C_HardwareInstancesAndStackups48 = notexist

                '46 DimensionalGeometry_Match_For_ALL/Has the dimensional geometry of the model been verified to match the vendor supplied data? (This is to include use of the hole tool for ALL features requiring fasteners) *
                comObj.C_DimensionalGeometry_Match_For_ALL49 = notexist

                '47 DimensionalGeometry_Match/Has the dimensional geometry of the model been verified to match the vendor supplied data? *
                comObj.C_DimensionalGeometry_Match50 = notexist

                '48 suppressedUnusedFeatures/ Have any suppressed (unused) features been removed from the model Pathfinder? *
                comObj.C_suppressedUnusedFeatures51 = notexist

                '49 ChildComponentsFullyConstrained/Are all child components fully constrained? (Ground constraints should be minimized for BEC generated models, but are acceptable for Vendor supplied assemblies) *
                comObj.C_ChildComponentsFullyConstrained52 = notexist

                '50 SimplifiedAssemblyModel/Is the assembly model simplified in design? (unnecessary details or "internal" parts should be removed if possible) *
                comObj.C_SimplifiedAssemblyModel53 = notexist

                '51 ChildPartOccurrenceProperties/Have the required occurrence properties for ALL child components been set to NO? (Reports/Parts List, Physical Properties, and Interference Analysis should all be set to NO) *
                comObj.C_ChildPartOccurrenceProperties54 = notexist

                '52 VendorMaterialData/Is the appropriate material applied to the model per Vendor supplied data? (If the component contains multiple material PURCHASED should be utilized) *
                comObj.C_VendorMaterialData55 = notexist

                '53 BaselineMassAndDensityUpdate/Is the density defined and the mass up to date? ("Update on file save" should be unchecked. When a Vendor weight is known "User-defined properties" should be checked and the mass applied) * (Manual check -if read-only)
                comObj.C_BaselineMassAndDensityUpdate56 = notexist

                '54 PMI_Instruction/Is PMI included for any other further or special situations? (kit identification, missing vendor information, etc.) *
                comObj.C_PMI_Instruction57 = notexist

                '55 BinderData/Does the model have a Binder attached and is properly filled out? (A hyperlink to the Vendor supplied data sheet located on the VAULT should be included) *
                comObj.C_BinderData58 = notexist

                '56 TerminalAssigned/Have the required terminals been assigned to the part model? *
                comObj.C_TerminalAssigned59 = notexist

                '57 ReferencePlanes/ Has the model been constructed on the correct reference planes per GEE0008? *
                comObj.C_ReferencePlanes60 = notexist

                '------------------------------------------------------------------------------------------------------------------------------------------
                'MTR assembly + sheetmetal+part 

                'assembly+part+sheetmetal
                '58 ConstrainedRefPlane/Verify the part model has been created and constrained to the correct reference plane.
                comObj.R_ConstrainedRefPlane61 = notexist

                '59 R_FullyConstrained/Verify that ALL parts/sub-assemblies have been fully constrained
                comObj.R_FullyConstrained62 = notexist

                '60 R_FullyConstrainedAndGround/Verify that ALL parts/sub-assemblies have been properly constrained***The “Ground” constraint should only be applied to the first component***
                comObj.R_FullyConstrainedAndGround63 = notexist

                'a+p
                '61 R_ProperConstrained/Verify proper constraints by utilizing “Assembly Relationship Manager”
                comObj.R_ProperConstrained64 = notexist

                'a+p+s
                '62 R_NeedSimplified/Verify if the assembly can be or needs to be simplified
                comObj.R_NeedSimplified65 = notexist

                '63 R_PreviewModelGEE/ matching part, background,to turn off shadows, sketches, coordinate systems And PMI reference planes off)
                comObj.R_PreviewModelGEE66 = notexist

                '64 R_WeightMass/Verify that the weight and mass are correct
                comObj.R_WeightMass67 = notexist

                '65 R_OccurrenceProperties/Verify the occurrence properties are set properly for the assembly model,
                comObj.R_OccurrenceProperties68 = notexist

                '66 R_PatternsWithinPatterns/Verify that “patterns” do not exist within other patterns. If present, they should be removed
                comObj.R_PatternsWithinPatterns69 = notexist

                '67 R_ExternallySuppliedModels/Do externally supplied models exist within the assembly? Are the models a surface model? Are the models a solid model?
                comObj.R_ExternallySuppliedModels70 = notexist

                '68 R_holePatterns/Verify that ALL hole patterns align properly with mating parts/components
                comObj.R_holePatterns71 = notexist

                '69 R_HoleDiameters/Verify that hole diameters are correct for size application
                comObj.R_HoleDiameters72 = notexist

                '70 R_HoleClearances/Verify that hole clearances are correct for fasteners used
                comObj.R_HoleClearances73 = notexist

                '71 R_HoleDistance/Verify that hole distance from edge of material is appropriate
                comObj.R_HoleDistance74 = notexist

                '72 R_PartLineupWithMatingParts/Verify that the part will line up correctly with mating parts
                comObj.R_PartLineupWithMatingParts75 = notexist

                '73 R_FITS_In_MatingParts/Verify that “FITS” are correct to mating parts Use Engineering Document MEA0001
                comObj.R_FITS_In_MatingParts76 = notexist

                'part
                '74 R_Density/Verify that the correct density has been applied to the model
                comObj.R_Density77 = notexist

                'p+s
                '75 R_Material/Verify the correct material is assigned to the part for the application
                comObj.R_Material = notexist

                'sheetmetal
                '76 R_UnusedSketches/Verify that any unused sketches have been removed
                comObj.R_UnusedSketches78 = notexist

                '77 R_linkSketches/Verify that any links within a sketch have been removed
                comObj.R_linkSketches79 = notexist

                '78 R_BinderProcess/Verify the Binder Document Process (GEE0001) was followed
                comObj.R_BinderProcess80 = notexist

                '79 R_BendRadius/Verify the bend radius meets requirements of the material selected
                comObj.R_BendRadius81 = notexist



                '-------------------------------------------------------------------------------------

                '38 date
                comObj.date38 = Date.Now()
                '39 lastModifiedDate
                comObj.lastModifiedDate39 = If(xlsSheet1.Rows(12).Cells(i).value = Nothing, er, xlsSheet1.Rows(12).Cells(i).value.ToString())
                '----------------------------------------------------------------------------------------------------------

                row1Data(comObj)



            Next
            '23component Name/What type of component?
            'mtcObj.componentName23 = er


            xlsWB.Close()
            xlsApp.Quit()
            releaseObject(xlsApp)
            releaseObject(xlsWB)
            releaseObject(xlsSheet1)
            killProcess()
        Catch ex As Exception
            'MsgBox($"Error while reading MTC_Electrical sheet {vbNewLine}{ex.Message}{vbNewLine} {ex.StackTrace}")
        End Try

    End Sub

    Public Sub MTR_Assembly(ByVal comObj As CommonReport)
        Try
            Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
            xlsWB = xlsApp.Workbooks.Open(comObj.excelFilePath, Nothing, False)

            Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing



            xlsSheet1 = xlsWB.Sheets("Assembly")
            xlsSheet1.Activate()
            Dim RowCnt = xlsSheet1.UsedRange.Rows.Count

            Dim ColCnt = xlsSheet1.UsedRange.Columns.Count


            xlsCell1 = xlsSheet1.UsedRange()


            Dim count = Nothing


            Projectname = If(xlsSheet1.Rows(1).Cells(4).value = Nothing, er, xlsSheet1.Rows(1).Cells(4).value.ToString())
            For i = 4 To ColCnt
                Dim row1 As DataRow = dt.NewRow()
                'initialization

                '0file
                comObj.fileName0 = "MTR"

                '1PartName Value(Error)
                comObj.partName1 = If(xlsSheet1.Rows(1).Cells(i).value = Nothing, er, xlsSheet1.Rows(1).Cells(i).value.ToString())


                '2Category
                comObj.category2 = "Assembly"




                '3author name/Who is the Author of the file?
                comObj.AuthorName3 = If(xlsSheet1.Rows(11).Cells(i).value = Nothing, er, xlsSheet1.Rows(11).Cells(i).value.ToString())


                '4projectname value/'  Is the correct "Project" used in the "Project Name" field? (Ensure consistent project titleIs used within a project) *
                comObj.projectName4 = Projectname


                '5part number match value/Is the part number match with M2M? *
                comObj.partNumberMatch5 = notexist





                '6document number/Is the Document Number correct? (BEC designed parts must have proper doc number.Hardware And other exceptions are N/A) *
                comObj.documentNumber6 = notexist



                '7authorNameIs the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) 
                comObj.AuthorName7 = notexist


                '8Dash/Verify the all categories below have been populated with relative Information Or, at a minimum, a dash (Summary, project, custom)
                comObj.DashInUnusedFiled8 = If(xlsSheet1.Rows(5).Cells(i).value = Nothing, er, xlsSheet1.Rows(5).Cells(i).value.ToString())

                '9 m2mdescription/Does the Model Title MATCH the M2M Item Master (INV) Description Field?
                comObj.m2mDiscriptionMatch9 = notexist


                '10uomsmatch/Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
                comObj.uomsMatch10 = notexist



                '11interface/Is interfernces found in assembly?
                comObj.interface11 = notexist


                '12broken interpartcopies/Verify that the inter-part copies are broken when released
                comObj.interpartcopies12 = If(xlsSheet1.Rows(3).Cells(i).value = Nothing, er, xlsSheet1.Rows(3).Cells(i).value.ToString())


                '13broken partcopies/Verify that the part copies are broken when released
                comObj.partcopies13 = If(xlsSheet1.Rows(4).Cells(i).value = Nothing, er, xlsSheet1.Rows(4).Cells(i).value.ToString())


                '14brokenfilebroken  file Path detected
                comObj.brokefile14 = notexist

                '15Adjustable/Is the part "Adjustable"? (part should NOT be adjustable) *
                comObj.adjustable15 = notexist


                '16Mat'l spec/Is the mat'l spec field populated? (indicated by RED in metadata audit file)
                comObj.Matl_spec16 = notexist

                '17MaterialUsed/Is the material used field populated? *
                comObj.MaterialUsed17 = notexist

                '18RemovedUnusedFeatures/Have any suppressed (unused) features been removed from the model Pathfinder? *
                comObj.RemovedUnusedFeatures18 = notexist


                '19ASTMminimum/Is the bend radius of the part equal to or above the ASTM minimum? *
                comObj.ASTMminimum19 = notexist

                '20FlatPattern/. Is the flat pattern turned ON? (Applicable only to parts needing dxf or cut length) *
                comObj.FlatPattern20 = notexist

                '21HoleToolUse/Did I use the hole tool for ALL holes requiring fasteners? (include clearance holes forhardware, tapped holes, And slots) *
                comObj.HoleToolUse21 = notexist




                '22virtualThread/Virtual thread applied for Fasteners?
                comObj.VirtualThread22 = notexist

                '23defined feature
                comObj.defineFeatures23 = notexist

                '24Supperessed Feature
                comObj.suppressedFeature24 = notexist

                '25Vendorpartnumber
                comObj.vendorPartNumber25 = notexist

                '26Hardwareparts
                comObj.HardwareParts26 = notexist

                '27m2mSourceMarked
                comObj.m2mSourceMarked27 = notexist

                '28SEStatus
                comObj.SEstatus28 = notexist

                '29exist features
                comObj.ExistFeatures29 = If(xlsSheet1.Rows(6).Cells(i).value = Nothing, er, xlsSheet1.Rows(6).Cells(i).value.ToString())

                '30MatingParts
                comObj.MatingParts30 = If(xlsSheet1.Rows(7).Cells(i).value = Nothing, er, xlsSheet1.Rows(7).Cells(i).value.ToString())

                '31 environment
                comObj.Environment31 = If(xlsSheet1.Rows(8).Cells(i).value = Nothing, er, xlsSheet1.Rows(8).Cells(i).value.ToString())

                '32Geometry
                comObj.Geometry32 = If(xlsSheet1.Rows(10).Cells(i).value = Nothing, er, xlsSheet1.Rows(10).Cells(i).value.ToString())

                '33UpdateOnFileSave33
                'comObj.UpdateOnFileSave33 = If(xlsSheet1.Rows(9).Cells(i).value = Nothing, er, xlsSheet1.Rows(9).Cells(i).value.ToString())

                '34ConstrainedFeatures34
                comObj.ConstrainedFeatures34 = notexist

                '35AllFeaturesRemoved35
                comObj.AllFeaturesRemoved35 = notexist

                '36HardwarePartBox
                comObj.HardwarePartBox36 = notexist

                '-------------------------------------------------------------------------------
                'mannual Questions
                '37 RefPartOccurrenceProp/reference parts?(higher level, report To PL, physical properties, all Set To "NO") * (Use occurrence property macro to recorrect setting)
                comObj.C_RefPartOccurrenceProp40 = notexist

                '38 AssemblyFeatures/Are the assembly features minimized? *
                comObj.C_AssemblyFeatures41 = notexist

                'a+p
                '39ModelPreview/Is the Preview of the model saved in conformance to the GEE? (On button select,
                comObj.C_ModelPreview42 = notexist

                '40PartConstraint/Record number of parts that are NOT fully constrained. **Record "0" if all parts are constrained correctly. (All parts locked and constraints convey design intent - i.e incorrect to "ground" parts when actual intent is for mating surfaces) *
                comObj.C_PartConstraint43 = notexist

                '---------------------------------------------------------------------------------------------------------
                'mtc part

                'p+sheetmetal
                '41Material_Spec_Field/Is the material spec field filled out with the proper designation? (Modification,Purchased, Assembly, etc.) *
                comObj.C_Material_Spec_Field44 = notexist

                '--------------------------------------------------------------------------------------------------------------------------------
                'mtc sheetmetal +part mass and density

                '42MassAndDensity_Update/Is the density defined and the mass up to date? * (macro will auto-update if file not read only)
                comObj.C_MassAndDensity_Update45 = notexist

                '43Bend_Radius_Update/Did I ensure the bend radius was producible? (Evaluate based on NPM0001C1 and/or consult ProdE/ManE) *
                comObj.C_Bend_Radius_Update46 = notexist

                '44 FlatPattern_Update/Is the flat pattern up to date? *
                comObj.C_FlatPattern_Update47 = notexist

                '--------------------------------------------------------------------------------------------------------------------------------
                'MTC Baseline
                '45 HardwareInstancesAndStackups/Are all instances of hardware consistent and have logical stackups? (Examples: 1.a SS nut applied to a zinc coated bolt having a copper washer; 2. Nine(9) instances of grade 8 and one(1) instance of grade 5 within a mounting configuration) *
                comObj.C_HardwareInstancesAndStackups48 = notexist

                '46 DimensionalGeometry_Match_For_ALL/Has the dimensional geometry of the model been verified to match the vendor supplied data? (This is to include use of the hole tool for ALL features requiring fasteners) *
                comObj.C_DimensionalGeometry_Match_For_ALL49 = notexist

                '47 DimensionalGeometry_Match/Has the dimensional geometry of the model been verified to match the vendor supplied data? *
                comObj.C_DimensionalGeometry_Match50 = notexist

                '48 suppressedUnusedFeatures/ Have any suppressed (unused) features been removed from the model Pathfinder? *
                comObj.C_suppressedUnusedFeatures51 = notexist

                '49 ChildComponentsFullyConstrained/Are all child components fully constrained? (Ground constraints should be minimized for BEC generated models, but are acceptable for Vendor supplied assemblies) *
                comObj.C_ChildComponentsFullyConstrained52 = notexist

                '50 SimplifiedAssemblyModel/Is the assembly model simplified in design? (unnecessary details or "internal" parts should be removed if possible) *
                comObj.C_SimplifiedAssemblyModel53 = notexist

                '51 ChildPartOccurrenceProperties/Have the required occurrence properties for ALL child components been set to NO? (Reports/Parts List, Physical Properties, and Interference Analysis should all be set to NO) *
                comObj.C_ChildPartOccurrenceProperties54 = notexist

                '52 VendorMaterialData/Is the appropriate material applied to the model per Vendor supplied data? (If the component contains multiple material PURCHASED should be utilized) *
                comObj.C_VendorMaterialData55 = notexist

                '53 BaselineMassAndDensityUpdate/Is the density defined and the mass up to date? ("Update on file save" should be unchecked. When a Vendor weight is known "User-defined properties" should be checked and the mass applied) * (Manual check -if read-only)
                comObj.C_BaselineMassAndDensityUpdate56 = notexist

                '54 PMI_Instruction/Is PMI included for any other further or special situations? (kit identification, missing vendor information, etc.) *
                comObj.C_PMI_Instruction57 = notexist

                '55 BinderData/Does the model have a Binder attached and is properly filled out? (A hyperlink to the Vendor supplied data sheet located on the VAULT should be included) *
                comObj.C_BinderData58 = notexist

                '56 TerminalAssigned/Have the required terminals been assigned to the part model? *
                comObj.C_TerminalAssigned59 = notexist

                '57 ReferencePlanes/ Has the model been constructed on the correct reference planes per GEE0008? *
                comObj.C_ReferencePlanes60 = notexist

                '------------------------------------------------------------------------------------------------------------------------------------------
                'MTR assembly + sheetmetal+part 

                'assembly+part+sheetmetal
                '58 ConstrainedRefPlane/Verify the part model has been created and constrained to the correct reference plane.
                comObj.R_ConstrainedRefPlane61 = If(xlsSheet1.Rows(13).Cells(i).value = Nothing, er, xlsSheet1.Rows(13).Cells(i).value.ToString())

                '59 R_FullyConstrained/Verify that ALL parts/sub-assemblies have been fully constrained
                comObj.R_FullyConstrained62 = If(xlsSheet1.Rows(14).Cells(i).value = Nothing, er, xlsSheet1.Rows(14).Cells(i).value.ToString())

                '60 R_FullyConstrainedAndGround/Verify that ALL parts/sub-assemblies have been properly constrained***The “Ground” constraint should only be applied to the first component***
                comObj.R_FullyConstrainedAndGround63 = If(xlsSheet1.Rows(15).Cells(i).value = Nothing, er, xlsSheet1.Rows(15).Cells(i).value.ToString())

                'a+p
                '61 R_ProperConstrained/Verify proper constraints by utilizing “Assembly Relationship Manager”
                comObj.R_ProperConstrained64 = If(xlsSheet1.Rows(16).Cells(i).value = Nothing, er, xlsSheet1.Rows(16).Cells(i).value.ToString())

                'a+p+s
                '62 R_NeedSimplified/Verify if the assembly can be or needs to be simplified
                comObj.R_NeedSimplified65 = If(xlsSheet1.Rows(17).Cells(i).value = Nothing, er, xlsSheet1.Rows(17).Cells(i).value.ToString())

                '63 R_PreviewModelGEE/ matching part, background,to turn off shadows, sketches, coordinate systems And PMI reference planes off)
                comObj.R_PreviewModelGEE66 = If(xlsSheet1.Rows(18).Cells(i).value = Nothing, er, xlsSheet1.Rows(18).Cells(i).value.ToString())

                '64 R_WeightMass/Verify that the weight and mass are correct
                comObj.R_WeightMass67 = If(xlsSheet1.Rows(19).Cells(i).value = Nothing, er, xlsSheet1.Rows(19).Cells(i).value.ToString())

                '65 R_OccurrenceProperties/Verify the occurrence properties are set properly for the assembly model,
                comObj.R_OccurrenceProperties68 = If(xlsSheet1.Rows(20).Cells(i).value = Nothing, er, xlsSheet1.Rows(20).Cells(i).value.ToString())

                '66 R_PatternsWithinPatterns/Verify that “patterns” do not exist within other patterns. If present, they should be removed
                comObj.R_PatternsWithinPatterns69 = If(xlsSheet1.Rows(21).Cells(i).value = Nothing, er, xlsSheet1.Rows(21).Cells(i).value.ToString())

                '67 R_ExternallySuppliedModels/Do externally supplied models exist within the assembly? Are the models a surface model? Are the models a solid model?
                comObj.R_ExternallySuppliedModels70 = If(xlsSheet1.Rows(22).Cells(i).value = Nothing, er, xlsSheet1.Rows(22).Cells(i).value.ToString())

                '68 R_holePatterns/Verify that ALL hole patterns align properly with mating parts/components
                comObj.R_holePatterns71 = If(xlsSheet1.Rows(23).Cells(i).value = Nothing, er, xlsSheet1.Rows(23).Cells(i).value.ToString())

                '69 R_HoleDiameters/Verify that hole diameters are correct for size application
                comObj.R_HoleDiameters72 = If(xlsSheet1.Rows(24).Cells(i).value = Nothing, er, xlsSheet1.Rows(24).Cells(i).value.ToString())

                '70 R_HoleClearances/Verify that hole clearances are correct for fasteners used
                comObj.R_HoleClearances73 = If(xlsSheet1.Rows(25).Cells(i).value = Nothing, er, xlsSheet1.Rows(25).Cells(i).value.ToString())

                '71 R_HoleDistance/Verify that hole distance from edge of material is appropriate
                comObj.R_HoleDistance74 = If(xlsSheet1.Rows(26).Cells(i).value = Nothing, er, xlsSheet1.Rows(26).Cells(i).value.ToString())

                '72 R_PartLineupWithMatingParts/Verify that the part will line up correctly with mating parts
                comObj.R_PartLineupWithMatingParts75 = If(xlsSheet1.Rows(27).Cells(i).value = Nothing, er, xlsSheet1.Rows(27).Cells(i).value.ToString())

                '73 R_FITS_In_MatingParts/Verify that “FITS” are correct to mating parts Use Engineering Document MEA0001
                comObj.R_FITS_In_MatingParts76 = If(xlsSheet1.Rows(28).Cells(i).value = Nothing, er, xlsSheet1.Rows(28).Cells(i).value.ToString())

                'part
                '74 R_Density/Verify that the correct density has been applied to the model
                comObj.R_Density77 = notexist

                'p+s
                '75 R_Material/Verify the correct material is assigned to the part for the application
                comObj.R_Material = notexist

                'sheetmetal
                '76 R_UnusedSketches/Verify that any unused sketches have been removed
                comObj.R_UnusedSketches78 = notexist

                '77 R_linkSketches/Verify that any links within a sketch have been removed
                comObj.R_linkSketches79 = notexist

                '78 R_BinderProcess/Verify the Binder Document Process (GEE0001) was followed
                comObj.R_BinderProcess80 = notexist

                '79 R_BendRadius/Verify the bend radius meets requirements of the material selected
                comObj.R_BendRadius81 = notexist



                '-------------------------------------------------------------------------------------

                '38 date
                comObj.date38 = Date.Now()

                '39 lastModifiedDate
                comObj.lastModifiedDate39 = If(xlsSheet1.Rows(12).Cells(i).value = Nothing, er, xlsSheet1.Rows(12).Cells(i).value.ToString())
                '------------------------------------------------------------------
                row1Data(comObj)



                '3Rivisionlevel/What is the revision level? *
                'If xlsSheet1.Rows(4).Cells(i).value = Nothing Then
                '    mtcObj.revisonlevel3 = "Null"
                'Else
                '    mtcObj.revisonlevel3 = xlsSheet1.Rows(4).Cells(i).value.ToString()
                'End If

                '6Is the Revision number correct? *
                'If xlsSheet1.Rows(7).Cells(i).value = Nothing Then
                '    mtcObj.revisoncorrect6 = "Null"
                'Else
                '    mtcObj.revisoncorrect6 = xlsSheet1.Rows(7).Cells(i).value.ToString()
                'End If


                '23component Name
                'mtcObj.componentName23 = er
            Next


            xlsWB.Close()
            xlsApp.Quit()
            releaseObject(xlsApp)
            releaseObject(xlsWB)
            releaseObject(xlsSheet1)
            killProcess()


        Catch ex As Exception
            'MsgBox($"Error while reading MTR_Assembly sheet {vbNewLine}{ex.Message}{vbNewLine} {ex.StackTrace}")
        End Try
    End Sub

    Public Sub MTR_Part(ByVal comObj As CommonReport)
        Try

            Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
            xlsWB = xlsApp.Workbooks.Open(comObj.excelFilePath, Nothing, False)

            Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing



            xlsSheet1 = xlsWB.Sheets("Part")
            xlsSheet1.Activate()
            Dim RowCnt = xlsSheet1.UsedRange.Rows.Count

            Dim ColCnt = xlsSheet1.UsedRange.Columns.Count


            xlsCell1 = xlsSheet1.UsedRange()

            Dim count = Nothing
            For i = 4 To ColCnt
                Dim row1 As DataRow = dt.NewRow()
                'initialization

                '0file
                comObj.fileName0 = "MTR"


                '1PartName Value(Error)
                comObj.partName1 = If(xlsSheet1.Rows(1).Cells(i).value = Nothing, er, xlsSheet1.Rows(1).Cells(i).value.ToString())


                '2Category
                comObj.category2 = "Part"




                '3author name/Who is the Author of the file?
                comObj.AuthorName3 = If(xlsSheet1.Rows(9).Cells(i).value = Nothing, er, xlsSheet1.Rows(9).Cells(i).value.ToString())


                '4projectname value/'  Is the correct "Project" used in the "Project Name" field? (Ensure consistent project titleIs used within a project) *
                comObj.projectName4 = Projectname

                '5part number match value/Is the part number match with M2M? *
                comObj.partNumberMatch5 = notexist





                '6document number/Is the Document Number correct? (BEC designed parts must have proper doc number.Hardware And other exceptions are N/A) *
                comObj.documentNumber6 = notexist



                '7authorNameIs the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) 
                comObj.AuthorName7 = notexist


                '8Dash/Verify the all categories below have been populated with relative Information Or, at a minimum, a dash (Summary, project, custom)
                comObj.DashInUnusedFiled8 = If(xlsSheet1.Rows(5).Cells(i).value = Nothing, er, xlsSheet1.Rows(5).Cells(i).value.ToString())

                '9 m2mdescription/Does the Model Title MATCH the M2M Item Master (INV) Description Field?
                comObj.m2mDiscriptionMatch9 = notexist


                '10uomsmatch/Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
                comObj.uomsMatch10 = notexist



                '11interface/Is interfernces found in assembly?
                comObj.interface11 = notexist


                '12broken interpartcopies/Verify that the inter-part copies are broken when released
                comObj.interpartcopies12 = notexist


                '13broken partcopies/Verify that the part copies are broken when released
                comObj.partcopies13 = notexist


                '14brokenfilebroken  file Path detected
                comObj.brokefile14 = notexist

                '15 Adjustable/Verify that the part model is NOT adjustable
                comObj.adjustable15 = If(xlsSheet1.Rows(4).Cells(i).value = Nothing, er, xlsSheet1.Rows(4).Cells(i).value.ToString())


                '16Mat'l spec/Is the mat'l spec field populated? (indicated by RED in metadata audit file)
                comObj.Matl_spec16 = notexist

                '17MaterialUsed/Is the material used field populated? *
                comObj.MaterialUsed17 = notexist

                '18RemovedUnusedFeatures/Have any suppressed (unused) features been removed from the model Pathfinder? *
                comObj.RemovedUnusedFeatures18 = notexist


                '19ASTMminimum/Is the bend radius of the part equal to or above the ASTM minimum? *
                comObj.ASTMminimum19 = notexist

                '20FlatPattern/. Is the flat pattern turned ON? (Applicable only to parts needing dxf or cut length) *
                comObj.FlatPattern20 = notexist

                '21HoleToolUse/Did I use the hole tool for ALL holes requiring fasteners? (include clearance holes forhardware, tapped holes, And slots) *
                comObj.HoleToolUse21 = notexist




                '22virtualThread/Virtual thread applied for Fasteners?
                comObj.VirtualThread22 = notexist

                '23defined feature
                comObj.defineFeatures23 = notexist

                '24Supperessed Feature
                comObj.suppressedFeature24 = notexist

                '25Vendorpartnumber
                comObj.vendorPartNumber25 = notexist

                '26Hardwareparts
                comObj.HardwareParts26 = notexist

                '27m2mSourceMarked
                comObj.m2mSourceMarked27 = notexist

                '28SEStatus
                comObj.SEstatus28 = notexist


                '29exist features
                comObj.ExistFeatures29 = notexist


                '30MatingParts
                comObj.MatingParts30 = notexist


                '31 environment
                comObj.Environment31 = notexist


                '32Geometry
                comObj.Geometry32 = notexist

                '33UpdateOnFileSave33
                'comObj.UpdateOnFileSave33 = If(xlsSheet1.Rows(8).Cells(i).value = Nothing, er, xlsSheet1.Rows(8).Cells(i).value.ToString())

                '34ConstrainedFeatures34
                comObj.ConstrainedFeatures34 = If(xlsSheet1.Rows(2).Cells(i).value = Nothing, er, xlsSheet1.Rows(2).Cells(i).value.ToString())

                '35AllFeaturesRemoved35
                comObj.AllFeaturesRemoved35 = If(xlsSheet1.Rows(3).Cells(i).value = Nothing, er, xlsSheet1.Rows(3).Cells(i).value.ToString())

                '36HardwarePartBox
                comObj.HardwarePartBox36 = If(xlsSheet1.Rows(6).Cells(i).value = Nothing, er, xlsSheet1.Rows(6).Cells(i).value.ToString())

                '-------------------------------------------------------------------------------
                'mannual Questions
                '37 RefPartOccurrenceProp/reference parts?(higher level, report To PL, physical properties, all Set To "NO") * (Use occurrence property macro to recorrect setting)
                comObj.C_RefPartOccurrenceProp40 = notexist

                '38 AssemblyFeatures/Are the assembly features minimized? *
                comObj.C_AssemblyFeatures41 = notexist

                'a+p
                '39ModelPreview/Is the Preview of the model saved in conformance to the GEE? (On button select,
                comObj.C_ModelPreview42 = notexist

                '40PartConstraint/Record number of parts that are NOT fully constrained. **Record "0" if all parts are constrained correctly. (All parts locked and constraints convey design intent - i.e incorrect to "ground" parts when actual intent is for mating surfaces) *
                comObj.C_PartConstraint43 = notexist

                '---------------------------------------------------------------------------------------------------------
                'mtc part

                'p+sheetmetal
                '41Material_Spec_Field/Is the material spec field filled out with the proper designation? (Modification,Purchased, Assembly, etc.) *
                comObj.C_Material_Spec_Field44 = notexist

                '--------------------------------------------------------------------------------------------------------------------------------
                'mtc sheetmetal +part mass and density

                '42MassAndDensity_Update/Is the density defined and the mass up to date? * (macro will auto-update if file not read only)
                comObj.C_MassAndDensity_Update45 = notexist

                '43Bend_Radius_Update/Did I ensure the bend radius was producible? (Evaluate based on NPM0001C1 and/or consult ProdE/ManE) *
                comObj.C_Bend_Radius_Update46 = notexist

                '44 FlatPattern_Update/Is the flat pattern up to date? *
                comObj.C_FlatPattern_Update47 = notexist

                '--------------------------------------------------------------------------------------------------------------------------------
                'MTC Baseline
                '45 HardwareInstancesAndStackups/Are all instances of hardware consistent and have logical stackups? (Examples: 1.a SS nut applied to a zinc coated bolt having a copper washer; 2. Nine(9) instances of grade 8 and one(1) instance of grade 5 within a mounting configuration) *
                comObj.C_HardwareInstancesAndStackups48 = notexist

                '46 DimensionalGeometry_Match_For_ALL/Has the dimensional geometry of the model been verified to match the vendor supplied data? (This is to include use of the hole tool for ALL features requiring fasteners) *
                comObj.C_DimensionalGeometry_Match_For_ALL49 = notexist

                '47 DimensionalGeometry_Match/Has the dimensional geometry of the model been verified to match the vendor supplied data? *
                comObj.C_DimensionalGeometry_Match50 = notexist

                '48 suppressedUnusedFeatures/ Have any suppressed (unused) features been removed from the model Pathfinder? *
                comObj.C_suppressedUnusedFeatures51 = notexist

                '49 ChildComponentsFullyConstrained/Are all child components fully constrained? (Ground constraints should be minimized for BEC generated models, but are acceptable for Vendor supplied assemblies) *
                comObj.C_ChildComponentsFullyConstrained52 = notexist

                '50 SimplifiedAssemblyModel/Is the assembly model simplified in design? (unnecessary details or "internal" parts should be removed if possible) *
                comObj.C_SimplifiedAssemblyModel53 = notexist

                '51 ChildPartOccurrenceProperties/Have the required occurrence properties for ALL child components been set to NO? (Reports/Parts List, Physical Properties, and Interference Analysis should all be set to NO) *
                comObj.C_ChildPartOccurrenceProperties54 = notexist

                '52 VendorMaterialData/Is the appropriate material applied to the model per Vendor supplied data? (If the component contains multiple material PURCHASED should be utilized) *
                comObj.C_VendorMaterialData55 = notexist

                '53 BaselineMassAndDensityUpdate/Is the density defined and the mass up to date? ("Update on file save" should be unchecked. When a Vendor weight is known "User-defined properties" should be checked and the mass applied) * (Manual check -if read-only)
                comObj.C_BaselineMassAndDensityUpdate56 = notexist

                '54 PMI_Instruction/Is PMI included for any other further or special situations? (kit identification, missing vendor information, etc.) *
                comObj.C_PMI_Instruction57 = notexist

                '55 BinderData/Does the model have a Binder attached and is properly filled out? (A hyperlink to the Vendor supplied data sheet located on the VAULT should be included) *
                comObj.C_BinderData58 = notexist

                '56 TerminalAssigned/Have the required terminals been assigned to the part model? *
                comObj.C_TerminalAssigned59 = notexist

                '57 ReferencePlanes/ Has the model been constructed on the correct reference planes per GEE0008? *
                comObj.C_ReferencePlanes60 = notexist

                '------------------------------------------------------------------------------------------------------------------------------------------
                'MTR assembly + sheetmetal+part 

                'assembly+part+sheetmetal
                '58 ConstrainedRefPlane/Verify the part model has been created and constrained to the correct reference plane.
                comObj.R_ConstrainedRefPlane61 = If(xlsSheet1.Rows(14).Cells(i).value = Nothing, er, xlsSheet1.Rows(14).Cells(i).value.ToString())

                '59 R_FullyConstrained/Verify that ALL parts/sub-assemblies have been fully constrained
                comObj.R_FullyConstrained62 = notexist

                '60 R_FullyConstrainedAndGround/Verify that ALL parts/sub-assemblies have been properly constrained***The “Ground” constraint should only be applied to the first component***
                comObj.R_FullyConstrainedAndGround63 = notexist

                'a+p
                '61 R_ProperConstrained/Verify proper constraints by utilizing “Assembly Relationship Manager”
                comObj.R_ProperConstrained64 = If(xlsSheet1.Rows(13).Cells(i).value = Nothing, er, xlsSheet1.Rows(13).Cells(i).value.ToString())

                'a+p+s
                '62 R_NeedSimplified/Verify if the assembly can be or needs to be simplified
                comObj.R_NeedSimplified65 = If(xlsSheet1.Rows(12).Cells(i).value = Nothing, er, xlsSheet1.Rows(12).Cells(i).value.ToString())

                '63 R_PreviewModelGEE/ matching part, background,to turn off shadows, sketches, coordinate systems And PMI reference planes off)
                comObj.R_PreviewModelGEE66 = If(xlsSheet1.Rows(11).Cells(i).value = Nothing, er, xlsSheet1.Rows(11).Cells(i).value.ToString())

                '64 R_WeightMass/Verify that the weight and mass are correct
                comObj.R_WeightMass67 = notexist

                '65 R_OccurrenceProperties/Verify the occurrence properties are set properly for the assembly model,
                comObj.R_OccurrenceProperties68 = notexist

                '66 R_PatternsWithinPatterns/Verify that “patterns” do not exist within other patterns. If present, they should be removed
                comObj.R_PatternsWithinPatterns69 = notexist

                '67 R_ExternallySuppliedModels/Do externally supplied models exist within the assembly? Are the models a surface model? Are the models a solid model?
                comObj.R_ExternallySuppliedModels70 = notexist

                '68 R_holePatterns/Verify that ALL hole patterns align properly with mating parts/components
                comObj.R_holePatterns71 = notexist

                '69 R_HoleDiameters/Verify that hole diameters are correct for size application
                comObj.R_HoleDiameters72 = If(xlsSheet1.Rows(16).Cells(i).value = Nothing, er, xlsSheet1.Rows(16).Cells(i).value.ToString())

                '70 R_HoleClearances/Verify that hole clearances are correct for fasteners used
                comObj.R_HoleClearances73 = If(xlsSheet1.Rows(17).Cells(i).value = Nothing, er, xlsSheet1.Rows(17).Cells(i).value.ToString())

                '71 R_HoleDistance/Verify that hole distance from edge of material is appropriate
                comObj.R_HoleDistance74 = If(xlsSheet1.Rows(18).Cells(i).value = Nothing, er, xlsSheet1.Rows(18).Cells(i).value.ToString())

                '72 R_PartLineupWithMatingParts/Verify that the part will line up correctly with mating parts
                comObj.R_PartLineupWithMatingParts75 = notexist

                '73 R_FITS_In_MatingParts/Verify that “FITS” are correct to mating parts Use Engineering Document MEA0001
                comObj.R_FITS_In_MatingParts76 = notexist

                'part
                '74 R_Density/Verify that the correct density has been applied to the model
                comObj.R_Density77 = If(xlsSheet1.Rows(15).Cells(i).value = Nothing, er, xlsSheet1.Rows(15).Cells(i).value.ToString())

                'p+s
                '75 R_Material/Verify the correct material is assigned to the part for the application
                comObj.R_Material = If(xlsSheet1.Rows(19).Cells(i).value = Nothing, er, xlsSheet1.Rows(19).Cells(i).value.ToString())

                'sheetmetal
                '76 R_UnusedSketches/Verify that any unused sketches have been removed
                comObj.R_UnusedSketches78 = notexist

                '77 R_linkSketches/Verify that any links within a sketch have been removed
                comObj.R_linkSketches79 = notexist

                '78 R_BinderProcess/Verify the Binder Document Process (GEE0001) was followed
                comObj.R_BinderProcess80 = notexist

                '79 R_BendRadius/Verify the bend radius meets requirements of the material selected
                comObj.R_BendRadius81 = notexist



                '-------------------------------------------------------------------------------------

                '38 date
                comObj.date38 = Date.Now()

                '39 lastModifiedDate
                comObj.lastModifiedDate39 = If(xlsSheet1.Rows(10).Cells(i).value = Nothing, er, xlsSheet1.Rows(10).Cells(i).value.ToString())
                '------------------------------------------------------------------
                row1Data(comObj)



                '3Rivisionlevel/What is the revision level? *
                'If xlsSheet1.Rows(4).Cells(i).value = Nothing Then
                '    mtcObj.revisonlevel3 = "Null"
                'Else
                '    mtcObj.revisonlevel3 = xlsSheet1.Rows(4).Cells(i).value.ToString()
                'End If

                '6Is the Revision number correct? *
                'If xlsSheet1.Rows(7).Cells(i).value = Nothing Then
                '    mtcObj.revisoncorrect6 = "Null"
                'Else
                '    mtcObj.revisoncorrect6 = xlsSheet1.Rows(7).Cells(i).value.ToString()
                'End If


                '23component Name
                'mtcObj.componentName23 = er
            Next



            xlsWB.Close()
            xlsApp.Quit()
            releaseObject(xlsApp)
            releaseObject(xlsWB)
            releaseObject(xlsSheet1)
            killProcess()


        Catch ex As Exception
            'MsgBox($"Error while reading MTR_Part sheet {vbNewLine}{ex.Message}{vbNewLine} {ex.StackTrace}")
        End Try

    End Sub

    Public Sub MTR_Sheetmetal(ByVal comObj As CommonReport)
        Try

            Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
            xlsWB = xlsApp.Workbooks.Open(comObj.excelFilePath, Nothing, False)

            Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing



            xlsSheet1 = xlsWB.Sheets("Sheetmetal")
            xlsSheet1.Activate()
            Dim RowCnt = xlsSheet1.UsedRange.Rows.Count

            Dim ColCnt = xlsSheet1.UsedRange.Columns.Count


            xlsCell1 = xlsSheet1.UsedRange()

            Dim count = Nothing
            For i = 4 To ColCnt
                Dim row1 As DataRow = dt.NewRow()
                'initialization

                '0file
                comObj.fileName0 = "MTR"


                '1PartName Value(Error)
                comObj.partName1 = If(xlsSheet1.Rows(1).Cells(i).value = Nothing, er, xlsSheet1.Rows(1).Cells(i).value.ToString())


                '2Category
                comObj.category2 = "Sheetmetal"




                '3author name/Who is the Author of the file?
                comObj.AuthorName3 = If(xlsSheet1.Rows(11).Cells(i).value = Nothing, er, xlsSheet1.Rows(11).Cells(i).value.ToString())


                '4projectname value/'  Is the correct "Project" used in the "Project Name" field? (Ensure consistent project titleIs used within a project) *
                comObj.projectName4 = Projectname


                '5part number match value/Is the part number match with M2M? *
                comObj.partNumberMatch5 = notexist





                '6document number/Is the Document Number correct? (BEC designed parts must have proper doc number.Hardware And other exceptions are N/A) *
                comObj.documentNumber6 = notexist



                '7authorNameIs the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) 
                comObj.AuthorName7 = notexist


                '8Dash/Verify the all categories below have been populated with relative Information Or, at a minimum, a dash (Summary, project, custom)
                comObj.DashInUnusedFiled8 = If(xlsSheet1.Rows(7).Cells(i).value = Nothing, er, xlsSheet1.Rows(7).Cells(i).value.ToString())

                '9 m2mdescription/Does the Model Title MATCH the M2M Item Master (INV) Description Field?
                comObj.m2mDiscriptionMatch9 = notexist


                '10uomsmatch/Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
                comObj.uomsMatch10 = notexist



                '11interface/Is interfernces found in assembly?
                comObj.interface11 = notexist


                '12broken interpartcopies/Verify that the inter-part copies are broken when released
                comObj.interpartcopies12 = If(xlsSheet1.Rows(5).Cells(i).value = Nothing, er, xlsSheet1.Rows(5).Cells(i).value.ToString())


                '13broken partcopies/Verify that the part copies are broken when released
                comObj.partcopies13 = If(xlsSheet1.Rows(6).Cells(i).value = Nothing, er, xlsSheet1.Rows(6).Cells(i).value.ToString())


                '14brokenfilebroken  file Path detected
                comObj.brokefile14 = notexist

                '15 Adjustable/Verify that the part model is NOT adjustable
                comObj.adjustable15 = If(xlsSheet1.Rows(4).Cells(i).value = Nothing, er, xlsSheet1.Rows(4).Cells(i).value.ToString())


                '16Mat'l spec/Is the mat'l spec field populated? (indicated by RED in metadata audit file)
                comObj.Matl_spec16 = notexist

                '17MaterialUsed/Is the material used field populated? *
                comObj.MaterialUsed17 = notexist

                '18RemovedUnusedFeatures/Have any suppressed (unused) features been removed from the model Pathfinder? *
                comObj.RemovedUnusedFeatures18 = notexist


                '19ASTMminimum/Is the bend radius of the part equal to or above the ASTM minimum? *
                comObj.ASTMminimum19 = notexist

                '20FlatPattern/. Is the flat pattern turned ON? (Applicable only to parts needing dxf or cut length) *
                comObj.FlatPattern20 = notexist

                '21HoleToolUse/Did I use the hole tool for ALL holes requiring fasteners? (include clearance holes forhardware, tapped holes, And slots) *
                comObj.HoleToolUse21 = notexist




                '22virtualThread/Virtual thread applied for Fasteners?
                comObj.VirtualThread22 = notexist

                '23defined feature
                comObj.defineFeatures23 = notexist

                '24Supperessed Feature
                comObj.suppressedFeature24 = notexist

                '25Vendorpartnumber
                comObj.vendorPartNumber25 = notexist

                '26Hardwareparts
                comObj.HardwareParts26 = notexist

                '27m2mSourceMarked
                comObj.m2mSourceMarked27 = notexist

                '28SEStatus
                comObj.SEstatus28 = notexist


                '29exist features
                comObj.ExistFeatures29 = notexist


                '30MatingParts
                comObj.MatingParts30 = notexist


                '31 environment
                comObj.Environment31 = notexist


                '32Geometry
                comObj.Geometry32 = If(xlsSheet1.Rows(10).Cells(i).value = Nothing, er, xlsSheet1.Rows(10).Cells(i).value.ToString())

                '33UpdateOnFileSave33
                'comObj.UpdateOnFileSave33 = If(xlsSheet1.Rows(9).Cells(i).value = Nothing, er, xlsSheet1.Rows(9).Cells(i).value.ToString())

                '34ConstrainedFeatures34
                comObj.ConstrainedFeatures34 = If(xlsSheet1.Rows(2).Cells(i).value = Nothing, er, xlsSheet1.Rows(2).Cells(i).value.ToString())

                '35AllFeaturesRemoved35
                comObj.AllFeaturesRemoved35 = If(xlsSheet1.Rows(3).Cells(i).value = Nothing, er, xlsSheet1.Rows(3).Cells(i).value.ToString())

                '36HardwarePartBox
                comObj.HardwarePartBox36 = notexist

                '-------------------------------------------------------------------------------
                'mannual Questions
                '37 RefPartOccurrenceProp/reference parts?(higher level, report To PL, physical properties, all Set To "NO") * (Use occurrence property macro to recorrect setting)
                comObj.C_RefPartOccurrenceProp40 = notexist

                '38 AssemblyFeatures/Are the assembly features minimized? *
                comObj.C_AssemblyFeatures41 = notexist

                'a+p
                '39ModelPreview/Is the Preview of the model saved in conformance to the GEE? (On button select,
                comObj.C_ModelPreview42 = If(xlsSheet1.Rows(18).Cells(i).value = Nothing, er, xlsSheet1.Rows(18).Cells(i).value.ToString())

                '40PartConstraint/Record number of parts that are NOT fully constrained. **Record "0" if all parts are constrained correctly. (All parts locked and constraints convey design intent - i.e incorrect to "ground" parts when actual intent is for mating surfaces) *
                comObj.C_PartConstraint43 = notexist

                '---------------------------------------------------------------------------------------------------------
                'mtc part

                'p+sheetmetal
                '41Material_Spec_Field/Is the material spec field filled out with the proper designation? (Modification,Purchased, Assembly, etc.) *
                comObj.C_Material_Spec_Field44 = notexist

                '--------------------------------------------------------------------------------------------------------------------------------
                'mtc sheetmetal +part mass and density

                '42MassAndDensity_Update/Is the density defined and the mass up to date? * (macro will auto-update if file not read only)
                comObj.C_MassAndDensity_Update45 = notexist

                '43Bend_Radius_Update/Did I ensure the bend radius was producible? (Evaluate based on NPM0001C1 and/or consult ProdE/ManE) *
                comObj.C_Bend_Radius_Update46 = notexist

                '44 FlatPattern_Update/Is the flat pattern up to date? *
                comObj.C_FlatPattern_Update47 = notexist

                '--------------------------------------------------------------------------------------------------------------------------------
                'MTC Baseline
                '45 HardwareInstancesAndStackups/Are all instances of hardware consistent and have logical stackups? (Examples: 1.a SS nut applied to a zinc coated bolt having a copper washer; 2. Nine(9) instances of grade 8 and one(1) instance of grade 5 within a mounting configuration) *
                comObj.C_HardwareInstancesAndStackups48 = notexist

                '46 DimensionalGeometry_Match_For_ALL/Has the dimensional geometry of the model been verified to match the vendor supplied data? (This is to include use of the hole tool for ALL features requiring fasteners) *
                comObj.C_DimensionalGeometry_Match_For_ALL49 = notexist

                '47 DimensionalGeometry_Match/Has the dimensional geometry of the model been verified to match the vendor supplied data? *
                comObj.C_DimensionalGeometry_Match50 = notexist

                '48 suppressedUnusedFeatures/ Have any suppressed (unused) features been removed from the model Pathfinder? *
                comObj.C_suppressedUnusedFeatures51 = notexist

                '49 ChildComponentsFullyConstrained/Are all child components fully constrained? (Ground constraints should be minimized for BEC generated models, but are acceptable for Vendor supplied assemblies) *
                comObj.C_ChildComponentsFullyConstrained52 = notexist

                '50 SimplifiedAssemblyModel/Is the assembly model simplified in design? (unnecessary details or "internal" parts should be removed if possible) *
                comObj.C_SimplifiedAssemblyModel53 = notexist

                '51 ChildPartOccurrenceProperties/Have the required occurrence properties for ALL child components been set to NO? (Reports/Parts List, Physical Properties, and Interference Analysis should all be set to NO) *
                comObj.C_ChildPartOccurrenceProperties54 = notexist

                '52 VendorMaterialData/Is the appropriate material applied to the model per Vendor supplied data? (If the component contains multiple material PURCHASED should be utilized) *
                comObj.C_VendorMaterialData55 = notexist

                '53 BaselineMassAndDensityUpdate/Is the density defined and the mass up to date? ("Update on file save" should be unchecked. When a Vendor weight is known "User-defined properties" should be checked and the mass applied) * (Manual check -if read-only)
                comObj.C_BaselineMassAndDensityUpdate56 = notexist

                '54 PMI_Instruction/Is PMI included for any other further or special situations? (kit identification, missing vendor information, etc.) *
                comObj.C_PMI_Instruction57 = notexist

                '55 BinderData/Does the model have a Binder attached and is properly filled out? (A hyperlink to the Vendor supplied data sheet located on the VAULT should be included) *
                comObj.C_BinderData58 = notexist

                '56 TerminalAssigned/Have the required terminals been assigned to the part model? *
                comObj.C_TerminalAssigned59 = notexist

                '57 ReferencePlanes/ Has the model been constructed on the correct reference planes per GEE0008? *
                comObj.C_ReferencePlanes60 = notexist

                '------------------------------------------------------------------------------------------------------------------------------------------
                'MTR assembly + sheetmetal+part 

                'assembly+part+sheetmetal
                '58 ConstrainedRefPlane/Verify the part model has been created and constrained to the correct reference plane.
                comObj.R_ConstrainedRefPlane61 = If(xlsSheet1.Rows(14).Cells(i).value = Nothing, er, xlsSheet1.Rows(14).Cells(i).value.ToString())

                '59 R_FullyConstrained/Verify that ALL parts/sub-assemblies have been fully constrained
                comObj.R_FullyConstrained62 = notexist

                '60 R_FullyConstrainedAndGround/Verify that ALL parts/sub-assemblies have been properly constrained***The “Ground” constraint should only be applied to the first component***
                comObj.R_FullyConstrainedAndGround63 = notexist

                'a+p
                '61 R_ProperConstrained/Verify that ALL features have been properly constrained
                comObj.R_ProperConstrained64 = If(xlsSheet1.Rows(15).Cells(i).value = Nothing, er, xlsSheet1.Rows(15).Cells(i).value.ToString())

                'a+p+s
                '62 R_NeedSimplified/Verify if the assembly can be or needs to be simplified
                comObj.R_NeedSimplified65 = If(xlsSheet1.Rows(13).Cells(i).value = Nothing, er, xlsSheet1.Rows(13).Cells(i).value.ToString())

                '63 R_PreviewModelGEE/ matching part, background,to turn off shadows, sketches, coordinate systems And PMI reference planes off)
                comObj.R_PreviewModelGEE66 = notexist

                '64 R_WeightMass/Verify that the weight and mass are correct
                comObj.R_WeightMass67 = If(xlsSheet1.Rows(19).Cells(i).value = Nothing, er, xlsSheet1.Rows(19).Cells(i).value.ToString())

                '65 R_OccurrenceProperties/Verify the occurrence properties are set properly for the assembly model,
                comObj.R_OccurrenceProperties68 = notexist

                '66 R_PatternsWithinPatterns/Verify that “patterns” do not exist within other patterns. If present, they should be removed
                comObj.R_PatternsWithinPatterns69 = notexist

                '67 R_ExternallySuppliedModels/Do externally supplied models exist within the assembly? Are the models a surface model? Are the models a solid model?
                comObj.R_ExternallySuppliedModels70 = notexist

                '68 R_holePatterns/Verify that ALL hole patterns align properly with mating parts/components
                comObj.R_holePatterns71 = If(xlsSheet1.Rows(21).Cells(i).value = Nothing, er, xlsSheet1.Rows(21).Cells(i).value.ToString())

                '69 R_HoleDiameters/Verify that hole diameters are correct for size application
                comObj.R_HoleDiameters72 = If(xlsSheet1.Rows(22).Cells(i).value = Nothing, er, xlsSheet1.Rows(22).Cells(i).value.ToString())

                '70 R_HoleClearances/Verify that hole clearances are correct for fasteners used
                comObj.R_HoleClearances73 = If(xlsSheet1.Rows(23).Cells(i).value = Nothing, er, xlsSheet1.Rows(23).Cells(i).value.ToString())

                '71 R_HoleDistance/Verify that hole distance from edge of material is appropriate
                comObj.R_HoleDistance74 = If(xlsSheet1.Rows(24).Cells(i).value = Nothing, er, xlsSheet1.Rows(24).Cells(i).value.ToString())

                '72 R_PartLineupWithMatingParts/Verify that the part will line up correctly with mating parts
                comObj.R_PartLineupWithMatingParts75 = notexist

                '73 R_FITS_In_MatingParts/Verify that “FITS” are correct to mating parts Use Engineering Document MEA0001
                comObj.R_FITS_In_MatingParts76 = If(xlsSheet1.Rows(25).Cells(i).value = Nothing, er, xlsSheet1.Rows(25).Cells(i).value.ToString())

                'part
                '74 R_Density/Verify that the correct density has been applied to the model
                comObj.R_Density77 = notexist

                'p+s
                '75 R_Material/Verify the correct material is assigned to the part for the application
                comObj.R_Material = If(xlsSheet1.Rows(26).Cells(i).value = Nothing, er, xlsSheet1.Rows(26).Cells(i).value.ToString())

                'sheetmetal
                '76 R_UnusedSketches/Verify that any unused sketches have been removed
                comObj.R_UnusedSketches78 = If(xlsSheet1.Rows(16).Cells(i).value = Nothing, er, xlsSheet1.Rows(16).Cells(i).value.ToString())

                '77 R_linkSketches/Verify that any links within a sketch have been removed
                comObj.R_linkSketches79 = If(xlsSheet1.Rows(17).Cells(i).value = Nothing, er, xlsSheet1.Rows(17).Cells(i).value.ToString())

                '78 R_BinderProcess/Verify the Binder Document Process (GEE0001) was followed
                comObj.R_BinderProcess80 = If(xlsSheet1.Rows(20).Cells(i).value = Nothing, er, xlsSheet1.Rows(20).Cells(i).value.ToString())

                '79 R_BendRadius/Verify the bend radius meets requirements of the material selected
                comObj.R_BendRadius81 = If(xlsSheet1.Rows(27).Cells(i).value = Nothing, er, xlsSheet1.Rows(27).Cells(i).value.ToString())



                '-------------------------------------------------------------------------------------

                '38 date
                comObj.date38 = Date.Now()
                '39 lastModifiedDate
                comObj.lastModifiedDate39 = If(xlsSheet1.Rows(12).Cells(i).value = Nothing, er, xlsSheet1.Rows(12).Cells(i).value.ToString())
                '------------------------------------------------------------------
                row1Data(comObj)



                '3Rivisionlevel/What is the revision level? *
                'If xlsSheet1.Rows(4).Cells(i).value = Nothing Then
                '    mtcObj.revisonlevel3 = "Null"
                'Else
                '    mtcObj.revisonlevel3 = xlsSheet1.Rows(4).Cells(i).value.ToString()
                'End If

                '6Is the Revision number correct? *
                'If xlsSheet1.Rows(7).Cells(i).value = Nothing Then
                '    mtcObj.revisoncorrect6 = "Null"
                'Else
                '    mtcObj.revisoncorrect6 = xlsSheet1.Rows(7).Cells(i).value.ToString()
                'End If


                '23component Name
                'mtcObj.componentName23 = er
            Next



            xlsWB.Close()
            xlsApp.Quit()
            releaseObject(xlsApp)
            releaseObject(xlsWB)
            releaseObject(xlsSheet1)
            killProcess()
        Catch ex As Exception
            'MsgBox($"Error while reading MTR_SheetMetal sheet {vbNewLine}{ex.Message}{vbNewLine} {ex.StackTrace}")
        End Try

    End Sub
    Public Sub row2Data(ByVal comObj As CommonReport)
        Try
            Dim row2 As DataRow = dt2.NewRow()
            dt2.Rows.Add(row2)


            If comObj.MTCfiles(comObj.i) = "" Then
                row2(0) = No
            Else
                row2(0) = comObj.MTCfiles(comObj.i)
            End If
            If comObj.MTRfiles(comObj.i) = "" Then
                row2(1) = No
            Else
                row2(1) = comObj.MTRfiles(comObj.i)
            End If


        Catch ex As Exception
            'Log.Error($"while getting Row2 Data {ex.Message} {ex.StackTrace}")
            'MsgBox($"Error while getting Row2 Data {vbNewLine}{ex.Message}{vbNewLine} {ex.StackTrace}")
        End Try
    End Sub

    Public Sub row1Data(ByVal comObj As CommonReport)
        Try

            If Count_4_MTC_MTR = 1 Then
#Region "MTC Row Data"
                Dim row1 As DataRow = dt.NewRow()
                dt.Rows.Add(row1)
                row1(0) = comObj.fileName0
                row1(1) = comObj.partName1
                row1(2) = comObj.category2

                Dim count As Integer = 0

                row1(3) = comObj.AuthorName3

                row1(4) = comObj.projectName4

                If comObj.partNumberMatch5 = "No" Then
                    row1(5) = No
                    count += 1
                ElseIf comObj.partNumberMatch5 = "Yes" Then
                    row1(5) = Yes
                Else
                    row1(5) = comObj.partNumberMatch5
                End If

                If comObj.documentNumber6 = "No" Then
                    row1(6) = No
                    count += 1
                ElseIf comObj.documentNumber6 = "Yes" Then
                    row1(6) = Yes
                Else
                    row1(6) = comObj.documentNumber6
                End If

                If comObj.AuthorName7 = "No" Then
                    row1(7) = No
                    count += 1
                ElseIf comObj.AuthorName7 = "Yes" Then
                    row1(7) = Yes
                Else
                    row1(7) = comObj.AuthorName7
                End If

                If comObj.DashInUnusedFiled8 = "No" Then
                    row1(8) = No
                    count += 1
                ElseIf comObj.DashInUnusedFiled8 = "Yes" Then
                    row1(8) = Yes
                Else
                    row1(8) = comObj.DashInUnusedFiled8
                End If
                '8
                If comObj.m2mDiscriptionMatch9 = "No" Then
                    row1(9) = No
                    count += 1
                ElseIf comObj.m2mDiscriptionMatch9 = "Yes" Then
                    row1(9) = Yes
                Else
                    row1(9) = comObj.m2mDiscriptionMatch9
                End If

                If comObj.uomsMatch10 = "No" Then
                    row1(10) = No
                    count += 1
                ElseIf comObj.uomsMatch10 = "Yes" Then
                    row1(10) = Yes
                Else
                    row1(10) = comObj.uomsMatch10
                End If
                '10 'mtc iterface no=0 yes=1
                If comObj.interface11 = "No" Then
                    row1(11) = Yes
                ElseIf comObj.interface11 = "Yes" Then
                    row1(11) = No
                    count += 1
                Else
                    row1(11) = comObj.interface11
                End If
                '11 mtc-interpartcopies no=0 yes=1
                If comObj.interpartcopies12 = "No" Then
                    row1(12) = Yes
                ElseIf comObj.interpartcopies12 = "Yes" Then
                    row1(12) = No
                    count += 1
                Else
                    row1(12) = comObj.interpartcopies12
                End If
                '12 mtc-partcopies no=0 yes=1
                If comObj.partcopies13 = "No" Then
                    row1(13) = Yes
                ElseIf comObj.partcopies13 = "Yes" Then
                    row1(13) = No
                    count += 1
                Else
                    row1(13) = comObj.partcopies13
                End If
                '13 mtc-brokefile no=0 yes=1
                If comObj.brokefile14 = "No" Then
                    row1(14) = Yes
                ElseIf comObj.brokefile14 = "Yes" Then
                    row1(14) = No
                    count += 1
                Else
                    row1(14) = comObj.brokefile14
                End If

                '14 mtc-adjustable no=0 yes=1
                If comObj.adjustable15 = "No" Then
                    row1(15) = Yes
                ElseIf comObj.adjustable15 = "Yes" Then
                    row1(15) = No
                    count += 1
                Else
                    row1(15) = comObj.adjustable15
                End If
                '15 
                If comObj.Matl_spec16 = "No" Then
                    row1(16) = No
                    count += 1
                ElseIf comObj.Matl_spec16 = "Yes" Then
                    row1(16) = Yes
                Else
                    row1(16) = comObj.MaterialUsed17
                End If
                '16
                If comObj.MaterialUsed17 = "No" Then
                    row1(17) = No
                    count += 1
                ElseIf comObj.MaterialUsed17 = "Yes" Then
                    row1(17) = Yes
                Else
                    row1(17) = comObj.MaterialUsed17
                End If
                '17
                If comObj.RemovedUnusedFeatures18 = "No" Then
                    row1(18) = No
                    count += 1
                ElseIf comObj.RemovedUnusedFeatures18 = "Yes" Then
                    row1(18) = Yes
                Else
                    row1(18) = comObj.RemovedUnusedFeatures18
                End If
                '18
                If comObj.ASTMminimum19 = "No" Then
                    row1(19) = No
                    count += 1
                ElseIf comObj.ASTMminimum19 = "Yes" Then
                    row1(19) = Yes
                Else
                    row1(19) = comObj.ASTMminimum19
                End If
                '19
                If comObj.FlatPattern20 = "No" Then
                    row1(20) = No
                    count += 1
                ElseIf comObj.FlatPattern20 = "Yes" Then
                    row1(20) = Yes
                Else
                    row1(20) = comObj.FlatPattern20
                End If
                '20
                If comObj.HoleToolUse21 = "No" Then
                    row1(21) = No
                    count += 1
                ElseIf comObj.HoleToolUse21 = "Yes" Then
                    row1(21) = Yes
                Else
                    row1(21) = comObj.HoleToolUse21
                End If


                If comObj.VirtualThread22 = "No" Then
                    row1(22) = No
                    count += 1
                ElseIf comObj.VirtualThread22 = "Yes" Then
                    row1(22) = Yes
                Else
                    row1(22) = comObj.VirtualThread22
                End If
                '22
                If comObj.defineFeatures23 = "No" Then
                    row1(23) = No
                    count += 1
                ElseIf comObj.defineFeatures23 = "Yes" Then
                    row1(23) = Yes
                Else
                    row1(23) = comObj.defineFeatures23
                End If
                '23 mtc-suppressedFeature no=0 yes=1
                If comObj.suppressedFeature24 = "No" Then
                    row1(24) = Yes
                ElseIf comObj.suppressedFeature24 = "Yes" Then
                    row1(24) = No
                    count += 1
                Else
                    row1(24) = comObj.suppressedFeature24
                End If

                If comObj.vendorPartNumber25 = "No" Then
                    row1(25) = No
                    count += 1
                ElseIf comObj.vendorPartNumber25 = "Yes" Then
                    row1(25) = Yes
                Else
                    row1(25) = comObj.vendorPartNumber25
                End If
                '25
                If comObj.HardwareParts26 = "No" Then
                    row1(26) = No
                    count += 1
                ElseIf comObj.HardwareParts26 = "Yes" Then
                    row1(26) = Yes
                Else
                    row1(26) = comObj.HardwareParts26
                End If
                '26
                If comObj.m2mSourceMarked27 = "No" Then
                    row1(27) = No
                    count += 1
                ElseIf comObj.m2mSourceMarked27 = "Yes" Then
                    row1(27) = Yes
                Else
                    row1(27) = comObj.m2mSourceMarked27
                End If
                '27
                If comObj.SEstatus28 = "No" Then
                    row1(28) = No
                    count += 1
                ElseIf comObj.SEstatus28 = "Yes" Then
                    row1(28) = Yes
                Else
                    row1(28) = comObj.SEstatus28
                End If

                If comObj.ExistFeatures29 = "No" Then
                    row1(29) = No
                    count += 1
                ElseIf comObj.ExistFeatures29 = "Yes" Then
                    row1(29) = Yes
                Else
                    row1(29) = comObj.ExistFeatures29
                End If

                If comObj.MatingParts30 = "No" Then
                    row1(30) = No
                    count += 1
                ElseIf comObj.MatingParts30 = "Yes" Then
                    row1(30) = Yes
                Else
                    row1(30) = comObj.MatingParts30
                End If


                If comObj.Environment31 = "No" Then
                    row1(31) = No
                    count += 1
                ElseIf comObj.Environment31 = "Yes" Then
                    row1(31) = Yes
                Else
                    row1(31) = comObj.Environment31
                End If

                If comObj.Geometry32 = "No" Then
                    row1(32) = No
                    count += 1
                ElseIf comObj.Geometry32 = "Yes" Then
                    row1(32) = Yes
                Else
                    row1(32) = comObj.Geometry32
                End If

                'If comObj.UpdateOnFileSave33 = "No" Then
                '    row1(33) = No
                '    count  += 1
                'ElseIf comObj.UpdateOnFileSave33 = "Yes" Then
                '    row1(33) = Yes
                'Else
                '    row1(33) = comObj.UpdateOnFileSave33
                'End If


                If comObj.ConstrainedFeatures34 = "No" Then
                    row1(33) = No
                    count += 1
                ElseIf comObj.ConstrainedFeatures34 = "Yes" Then
                    row1(33) = Yes
                Else
                    row1(33) = comObj.ConstrainedFeatures34
                End If


                If comObj.AllFeaturesRemoved35 = "No" Then
                    row1(34) = No
                    count += 1
                ElseIf comObj.AllFeaturesRemoved35 = "Yes" Then
                    row1(34) = Yes
                Else
                    row1(34) = comObj.AllFeaturesRemoved35
                End If

                If comObj.HardwarePartBox36 = "No" Then
                    row1(35) = No
                    count += 1
                ElseIf comObj.HardwarePartBox36 = "Yes" Then
                    row1(35) = Yes
                Else
                    row1(35) = comObj.HardwarePartBox36
                End If
                '---------------------------------------------------
                If comObj.C_RefPartOccurrenceProp40 = "No" Then
                    row1(36) = No
                    count += 1
                ElseIf comObj.C_RefPartOccurrenceProp40 = "Yes" Then
                    row1(36) = Yes
                Else
                    row1(36) = comObj.C_RefPartOccurrenceProp40
                End If

                If comObj.C_AssemblyFeatures41 = "No" Then
                    row1(37) = No
                    count += 1
                ElseIf comObj.C_AssemblyFeatures41 = "Yes" Then
                    row1(37) = Yes
                Else
                    row1(37) = comObj.C_AssemblyFeatures41
                End If

                If comObj.C_ModelPreview42 = "No" Then
                    row1(38) = No
                    count += 1
                ElseIf comObj.C_ModelPreview42 = "Yes" Then
                    row1(38) = Yes
                Else
                    row1(38) = comObj.C_ModelPreview42
                End If

                If comObj.C_PartConstraint43 = "No" Then
                    row1(39) = No
                    count += 1
                ElseIf comObj.C_PartConstraint43 = "Yes" Then
                    row1(39) = Yes
                ElseIf comObj.C_PartConstraint43 = "0" Then
                    row1(39) = No
                    count += 1
                Else
                    row1(39) = Yes
                    'row1(39) = comObj.C_PartConstraint43
                End If

                If comObj.C_Material_Spec_Field44 = "No" Then
                    row1(40) = No
                    count += 1
                ElseIf comObj.C_Material_Spec_Field44 = "Yes" Then
                    row1(40) = Yes
                Else
                    row1(40) = comObj.C_Material_Spec_Field44
                End If

                If comObj.C_MassAndDensity_Update45 = "No" Then
                    row1(41) = No
                    count += 1
                ElseIf comObj.C_MassAndDensity_Update45 = "Yes" Then
                    row1(41) = Yes
                Else
                    row1(41) = comObj.C_MassAndDensity_Update45
                End If

                If comObj.C_Bend_Radius_Update46 = "No" Then
                    row1(42) = No
                    count += 1
                ElseIf comObj.C_Bend_Radius_Update46 = "Yes" Then
                    row1(42) = Yes
                Else
                    row1(42) = comObj.C_Bend_Radius_Update46
                End If

                If comObj.C_FlatPattern_Update47 = "No" Then
                    row1(43) = No
                    count += 1
                ElseIf comObj.C_FlatPattern_Update47 = "Yes" Then
                    row1(43) = Yes
                Else
                    row1(43) = comObj.C_FlatPattern_Update47
                End If

                If comObj.C_HardwareInstancesAndStackups48 = "No" Then
                    row1(44) = No
                    count += 1
                ElseIf comObj.C_HardwareInstancesAndStackups48 = "Yes" Then
                    row1(44) = Yes
                Else
                    row1(44) = comObj.C_HardwareInstancesAndStackups48
                End If

                If comObj.C_DimensionalGeometry_Match_For_ALL49 = "No" Then
                    row1(45) = No
                    count += 1
                ElseIf comObj.C_DimensionalGeometry_Match_For_ALL49 = "Yes" Then
                    row1(45) = Yes
                Else
                    row1(45) = comObj.C_DimensionalGeometry_Match_For_ALL49
                End If

                If comObj.C_DimensionalGeometry_Match50 = "No" Then
                    row1(46) = No
                    count += 1
                ElseIf comObj.C_DimensionalGeometry_Match50 = "Yes" Then
                    row1(46) = Yes
                Else
                    row1(46) = comObj.C_DimensionalGeometry_Match50
                End If

                If comObj.C_suppressedUnusedFeatures51 = "No" Then
                    row1(47) = No
                    count += 1
                ElseIf comObj.C_suppressedUnusedFeatures51 = "Yes" Then
                    row1(47) = Yes
                Else
                    row1(47) = comObj.C_suppressedUnusedFeatures51
                End If

                If comObj.C_ChildComponentsFullyConstrained52 = "No" Then
                    row1(48) = No
                    count += 1
                ElseIf comObj.C_ChildComponentsFullyConstrained52 = "Yes" Then
                    row1(48) = Yes
                Else
                    row1(48) = comObj.C_ChildComponentsFullyConstrained52
                End If

                If comObj.C_SimplifiedAssemblyModel53 = "No" Then
                    row1(49) = No
                    count += 1
                ElseIf comObj.C_SimplifiedAssemblyModel53 = "Yes" Then
                    row1(49) = Yes
                Else
                    row1(49) = comObj.C_SimplifiedAssemblyModel53
                End If

                If comObj.C_ChildPartOccurrenceProperties54 = "No" Then
                    row1(50) = No
                    count += 1
                ElseIf comObj.C_ChildPartOccurrenceProperties54 = "Yes" Then
                    row1(50) = Yes
                Else
                    row1(50) = comObj.C_ChildPartOccurrenceProperties54
                End If

                If comObj.C_VendorMaterialData55 = "No" Then
                    row1(51) = No
                    count += 1
                ElseIf comObj.C_VendorMaterialData55 = "Yes" Then
                    row1(51) = Yes
                Else
                    row1(51) = comObj.C_VendorMaterialData55
                End If

                If comObj.C_BaselineMassAndDensityUpdate56 = "No" Then
                    row1(52) = No
                    count += 1
                ElseIf comObj.C_BaselineMassAndDensityUpdate56 = "Yes" Then
                    row1(52) = Yes
                Else
                    row1(52) = comObj.C_BaselineMassAndDensityUpdate56
                End If

                If comObj.C_PMI_Instruction57 = "No" Then
                    row1(53) = No
                    count += 1
                ElseIf comObj.C_PMI_Instruction57 = "Yes" Then
                    row1(53) = Yes
                Else
                    row1(53) = comObj.C_PMI_Instruction57
                End If

                If comObj.C_BinderData58 = "No" Then
                    row1(54) = No
                    count += 1
                ElseIf comObj.C_BinderData58 = "Yes" Then
                    row1(54) = Yes
                Else
                    row1(54) = comObj.C_BinderData58
                End If

                If comObj.C_TerminalAssigned59 = "No" Then
                    row1(55) = No
                    count += 1
                ElseIf comObj.C_TerminalAssigned59 = "Yes" Then
                    row1(55) = Yes
                Else
                    row1(55) = comObj.C_TerminalAssigned59
                End If

                If comObj.C_ReferencePlanes60 = "No" Then
                    row1(56) = No
                    count += 1
                ElseIf comObj.C_ReferencePlanes60 = "Yes" Then
                    row1(56) = Yes
                Else
                    row1(56) = comObj.C_ReferencePlanes60
                End If

                If comObj.R_ConstrainedRefPlane61 = "No" Then
                    row1(57) = No
                    count += 1
                ElseIf comObj.R_ConstrainedRefPlane61 = "Yes" Then
                    row1(57) = Yes
                Else
                    row1(57) = comObj.R_ConstrainedRefPlane61
                End If

                If comObj.R_FullyConstrained62 = "No" Then
                    row1(58) = No
                    count += 1
                ElseIf comObj.R_FullyConstrained62 = "Yes" Then
                    row1(58) = Yes
                Else
                    row1(58) = comObj.R_FullyConstrained62
                End If

                If comObj.R_FullyConstrainedAndGround63 = "No" Then
                    row1(59) = No
                    count += 1
                ElseIf comObj.R_FullyConstrainedAndGround63 = "Yes" Then
                    row1(59) = Yes
                Else
                    row1(59) = comObj.R_FullyConstrainedAndGround63
                End If

                If comObj.R_ProperConstrained64 = "No" Then
                    row1(60) = No
                    count += 1
                ElseIf comObj.R_ProperConstrained64 = "Yes" Then
                    row1(60) = Yes
                Else
                    row1(60) = comObj.R_ProperConstrained64
                End If

                'needsimplification no=0,yes=1
                If comObj.R_NeedSimplified65 = "No" Then
                    row1(61) = Yes
                ElseIf comObj.R_NeedSimplified65 = "Yes" Then
                    row1(61) = No
                    count += 1
                Else
                    row1(61) = comObj.R_NeedSimplified65
                End If

                If comObj.R_PreviewModelGEE66 = "No" Then
                    row1(62) = No
                    count += 1
                ElseIf comObj.R_PreviewModelGEE66 = "Yes" Then
                    row1(62) = Yes
                Else
                    row1(62) = comObj.R_PreviewModelGEE66
                End If

                If comObj.R_WeightMass67 = "No" Then
                    row1(63) = No
                    count += 1
                ElseIf comObj.R_WeightMass67 = "Yes" Then
                    row1(63) = Yes
                Else
                    row1(63) = comObj.R_WeightMass67
                End If

                If comObj.R_OccurrenceProperties68 = "No" Then
                    row1(64) = No
                    count += 1
                ElseIf comObj.R_OccurrenceProperties68 = "Yes" Then
                    row1(64) = Yes
                Else
                    row1(64) = comObj.R_OccurrenceProperties68
                End If

                If comObj.R_PatternsWithinPatterns69 = "No" Then
                    row1(65) = No
                    count += 1
                ElseIf comObj.R_PatternsWithinPatterns69 = "Yes" Then
                    row1(65) = Yes
                Else
                    row1(65) = comObj.R_PatternsWithinPatterns69
                End If

                'R_ExternallySuppliedModels70 no=0,yes=1
                If comObj.R_ExternallySuppliedModels70 = "No" Then
                    row1(66) = Yes
                ElseIf comObj.R_ExternallySuppliedModels70 = "Yes" Then
                    row1(66) = No
                    count += 1
                Else
                    row1(66) = comObj.R_ExternallySuppliedModels70
                End If

                If comObj.R_holePatterns71 = "No" Then
                    row1(67) = No
                    count += 1
                ElseIf comObj.R_holePatterns71 = "Yes" Then
                    row1(67) = Yes
                Else
                    row1(67) = comObj.R_holePatterns71
                End If

                If comObj.R_HoleDiameters72 = "No" Then
                    row1(68) = No
                    count += 1
                ElseIf comObj.R_HoleDiameters72 = "Yes" Then
                    row1(68) = Yes
                Else
                    row1(68) = comObj.R_HoleDiameters72
                End If

                If comObj.R_HoleClearances73 = "No" Then
                    row1(69) = No
                    count += 1
                ElseIf comObj.R_HoleClearances73 = "Yes" Then
                    row1(69) = Yes
                Else
                    row1(69) = comObj.R_HoleClearances73
                End If

                If comObj.R_HoleDistance74 = "No" Then
                    row1(70) = No
                    count += 1
                ElseIf comObj.R_HoleDistance74 = "Yes" Then
                    row1(70) = Yes
                Else
                    row1(70) = comObj.R_HoleDistance74
                End If

                If comObj.R_PartLineupWithMatingParts75 = "No" Then
                    row1(71) = No
                    count += 1
                ElseIf comObj.R_PartLineupWithMatingParts75 = "Yes" Then
                    row1(71) = Yes
                Else
                    row1(71) = comObj.R_PartLineupWithMatingParts75
                End If

                If comObj.R_FITS_In_MatingParts76 = "No" Then
                    row1(72) = No
                    count += 1
                ElseIf comObj.R_FITS_In_MatingParts76 = "Yes" Then
                    row1(72) = Yes
                Else
                    row1(72) = comObj.R_FITS_In_MatingParts76
                End If

                If comObj.R_Density77 = "No" Then
                    row1(73) = No
                    count += 1
                ElseIf comObj.R_Density77 = "Yes" Then
                    row1(73) = Yes
                Else
                    row1(73) = comObj.R_Density77
                End If

                If comObj.R_Material = "No" Then
                    row1(74) = No
                    count += 1
                ElseIf comObj.R_Material = "Yes" Then
                    row1(74) = Yes
                Else
                    row1(74) = comObj.R_Material
                End If

                If comObj.R_UnusedSketches78 = "No" Then
                    row1(75) = No
                    count += 1
                ElseIf comObj.R_UnusedSketches78 = "Yes" Then
                    row1(75) = Yes
                Else
                    row1(75) = comObj.R_UnusedSketches78
                End If

                If comObj.R_linkSketches79 = "No" Then
                    row1(76) = No
                    count += 1
                ElseIf comObj.R_linkSketches79 = "Yes" Then
                    row1(76) = Yes
                Else
                    row1(76) = comObj.R_linkSketches79
                End If

                If comObj.R_BinderProcess80 = "No" Then
                    row1(77) = No
                    count += 1
                ElseIf comObj.R_BinderProcess80 = "Yes" Then
                    row1(77) = Yes
                Else
                    row1(77) = comObj.R_BinderProcess80
                End If

                If comObj.R_BendRadius81 = "No" Then
                    row1(78) = No
                    count += 1
                ElseIf comObj.R_BendRadius81 = "Yes" Then
                    row1(78) = Yes
                Else
                    row1(78) = comObj.R_BendRadius81
                End If
                '-----------------------------------------------------------------------------
                '77 density -> material

                row1(79) = count
                row1(80) = comObj.date38

                If comObj.lastModifiedDate39 = "No" Then
                    row1(81) = "0"

                Else
                    row1(81) = comObj.lastModifiedDate39
                End If

                count = 0
#End Region
            ElseIf Count_4_MTC_MTR = 2 Then
#Region "MTR Row Data"
                Dim row1 As DataRow = dt.NewRow()
                dt.Rows.Add(row1)
                row1(0) = comObj.fileName0
                row1(1) = comObj.partName1
                row1(2) = comObj.category2

                Dim count As Integer = 0

                row1(3) = comObj.AuthorName3

                row1(4) = comObj.projectName4

                If comObj.partNumberMatch5 = "No" Then
                    row1(5) = No
                    count += 1
                ElseIf comObj.partNumberMatch5 = "Yes" Then
                    row1(5) = Yes
                Else
                    row1(5) = comObj.partNumberMatch5
                End If

                If comObj.documentNumber6 = "No" Then
                    row1(6) = No
                    count += 1
                ElseIf comObj.documentNumber6 = "Yes" Then
                    row1(6) = Yes
                Else
                    row1(6) = comObj.documentNumber6
                End If

                If comObj.AuthorName7 = "No" Then
                    row1(7) = No
                    count += 1
                ElseIf comObj.AuthorName7 = "Yes" Then
                    row1(7) = Yes
                Else
                    row1(7) = comObj.AuthorName7
                End If

                If comObj.DashInUnusedFiled8 = "No" Then
                    row1(8) = No
                    count += 1
                ElseIf comObj.DashInUnusedFiled8 = "Yes" Then
                    row1(8) = Yes
                Else
                    row1(8) = comObj.DashInUnusedFiled8
                End If
                '8
                If comObj.m2mDiscriptionMatch9 = "No" Then
                    row1(9) = No
                    count += 1
                ElseIf comObj.m2mDiscriptionMatch9 = "Yes" Then
                    row1(9) = Yes
                Else
                    row1(9) = comObj.m2mDiscriptionMatch9
                End If

                If comObj.uomsMatch10 = "No" Then
                    row1(10) = No
                    count += 1
                ElseIf comObj.uomsMatch10 = "Yes" Then
                    row1(10) = Yes
                Else
                    row1(10) = comObj.uomsMatch10
                End If
                '10
                If comObj.interface11 = "No" Then
                    row1(11) = No
                    count += 1
                ElseIf comObj.interface11 = "Yes" Then
                    row1(11) = Yes
                Else
                    row1(11) = comObj.interface11
                End If
                '11 mtr-interpartcopies no=0 yes=1
                If comObj.interpartcopies12 = "No" Then
                    row1(12) = Yes
                    count += 1
                ElseIf comObj.interpartcopies12 = "Yes" Then
                    row1(12) = No
                Else
                    row1(12) = comObj.interpartcopies12
                End If
                '12 mtr-partcopies no=0 yes=1
                If comObj.partcopies13 = "No" Then
                    row1(13) = Yes
                    count += 1
                ElseIf comObj.partcopies13 = "Yes" Then
                    row1(13) = No
                Else
                    row1(13) = comObj.partcopies13
                End If
                '13
                If comObj.brokefile14 = "No" Then
                    row1(14) = No
                    count += 1
                ElseIf comObj.brokefile14 = "Yes" Then
                    row1(14) = Yes
                Else
                    row1(14) = comObj.brokefile14
                End If
                '14 mtr-adjustable no=0 yes=1
                If comObj.adjustable15 = "No" Then
                    row1(15) = Yes
                    count += 1
                ElseIf comObj.adjustable15 = "Yes" Then
                    row1(15) = No
                Else
                    row1(15) = comObj.adjustable15
                End If
                '15
                If comObj.Matl_spec16 = "No" Then
                    row1(16) = No
                    count += 1
                ElseIf comObj.Matl_spec16 = "Yes" Then
                    row1(16) = Yes
                Else
                    row1(16) = comObj.MaterialUsed17
                End If
                '16
                If comObj.MaterialUsed17 = "No" Then
                    row1(17) = No
                    count += 1
                ElseIf comObj.MaterialUsed17 = "Yes" Then
                    row1(17) = Yes
                Else
                    row1(17) = comObj.MaterialUsed17
                End If
                '17
                If comObj.RemovedUnusedFeatures18 = "No" Then
                    row1(18) = No
                    count += 1
                ElseIf comObj.RemovedUnusedFeatures18 = "Yes" Then
                    row1(18) = Yes
                Else
                    row1(18) = comObj.RemovedUnusedFeatures18
                End If
                '18
                If comObj.ASTMminimum19 = "No" Then
                    row1(19) = No
                    count += 1
                ElseIf comObj.ASTMminimum19 = "Yes" Then
                    row1(19) = Yes
                Else
                    row1(19) = comObj.ASTMminimum19
                End If
                '19
                If comObj.FlatPattern20 = "No" Then
                    row1(20) = No
                    count += 1
                ElseIf comObj.FlatPattern20 = "Yes" Then
                    row1(20) = Yes
                Else
                    row1(20) = comObj.FlatPattern20
                End If
                '20
                If comObj.HoleToolUse21 = "No" Then
                    row1(21) = No
                    count += 1
                ElseIf comObj.HoleToolUse21 = "Yes" Then
                    row1(21) = Yes
                Else
                    row1(21) = comObj.HoleToolUse21
                End If


                If comObj.VirtualThread22 = "No" Then
                    row1(22) = No
                    count += 1
                ElseIf comObj.VirtualThread22 = "Yes" Then
                    row1(22) = Yes
                Else
                    row1(22) = comObj.VirtualThread22
                End If
                '22
                If comObj.defineFeatures23 = "No" Then
                    row1(23) = No
                    count += 1
                ElseIf comObj.defineFeatures23 = "Yes" Then
                    row1(23) = Yes
                Else
                    row1(23) = comObj.defineFeatures23
                End If
                '23
                If comObj.suppressedFeature24 = "No" Then
                    row1(24) = No
                    count += 1
                ElseIf comObj.suppressedFeature24 = "Yes" Then
                    row1(24) = Yes
                Else
                    row1(24) = comObj.suppressedFeature24
                End If

                If comObj.vendorPartNumber25 = "No" Then
                    row1(25) = No
                    count += 1
                ElseIf comObj.vendorPartNumber25 = "Yes" Then
                    row1(25) = Yes
                Else
                    row1(25) = comObj.vendorPartNumber25
                End If
                '25
                If comObj.HardwareParts26 = "No" Then
                    row1(26) = No
                    count += 1
                ElseIf comObj.HardwareParts26 = "Yes" Then
                    row1(26) = Yes
                Else
                    row1(26) = comObj.HardwareParts26
                End If
                '26
                If comObj.m2mSourceMarked27 = "No" Then
                    row1(27) = No
                    count += 1
                ElseIf comObj.m2mSourceMarked27 = "Yes" Then
                    row1(27) = Yes
                Else
                    row1(27) = comObj.m2mSourceMarked27
                End If
                '27
                If comObj.SEstatus28 = "No" Then
                    row1(28) = No
                    count += 1
                ElseIf comObj.SEstatus28 = "Yes" Then
                    row1(28) = Yes
                Else
                    row1(28) = comObj.SEstatus28
                End If

                If comObj.ExistFeatures29 = "No" Then
                    row1(29) = No
                    count += 1
                ElseIf comObj.ExistFeatures29 = "Yes" Then
                    row1(29) = Yes
                Else
                    row1(29) = comObj.ExistFeatures29
                End If

                If comObj.MatingParts30 = "No" Then
                    row1(30) = No
                    count += 1
                ElseIf comObj.MatingParts30 = "Yes" Then
                    row1(30) = Yes
                Else
                    row1(30) = comObj.MatingParts30
                End If


                If comObj.Environment31 = "No" Then
                    row1(31) = No
                    count += 1
                ElseIf comObj.Environment31 = "Yes" Then
                    row1(31) = Yes
                Else
                    row1(31) = comObj.Environment31
                End If

                If comObj.Geometry32 = "No" Then
                    row1(32) = No
                    count += 1
                ElseIf comObj.Geometry32 = "Yes" Then
                    row1(32) = Yes
                Else
                    row1(32) = comObj.Geometry32
                End If

                'If comObj.UpdateOnFileSave33 = "No" Then
                '    row1(33) = No
                '    count  += 1
                'ElseIf comObj.UpdateOnFileSave33 = "Yes" Then
                '    row1(33) = Yes
                'Else
                '    row1(33) = comObj.UpdateOnFileSave33
                'End If


                If comObj.ConstrainedFeatures34 = "No" Then
                    row1(33) = No
                    count += 1
                ElseIf comObj.ConstrainedFeatures34 = "Yes" Then
                    row1(33) = Yes
                Else
                    row1(33) = comObj.ConstrainedFeatures34
                End If

                'mtr-AllFeaturesRemoved no=0 yes=1
                If comObj.AllFeaturesRemoved35 = "No" Then
                    row1(34) = Yes
                    count += 1
                ElseIf comObj.AllFeaturesRemoved35 = "Yes" Then
                    row1(34) = No
                Else
                    row1(34) = comObj.AllFeaturesRemoved35
                End If

                If comObj.HardwarePartBox36 = "No" Then
                    row1(35) = No
                    count += 1
                ElseIf comObj.HardwarePartBox36 = "Yes" Then
                    row1(35) = Yes
                Else
                    row1(35) = comObj.HardwarePartBox36
                End If

                '---------------------------------------------------
                If comObj.C_RefPartOccurrenceProp40 = "No" Then
                    row1(36) = No
                    count += 1
                ElseIf comObj.C_RefPartOccurrenceProp40 = "Yes" Then
                    row1(36) = Yes
                Else
                    row1(36) = comObj.C_RefPartOccurrenceProp40
                End If

                If comObj.C_AssemblyFeatures41 = "No" Then
                    row1(37) = No
                    count += 1
                ElseIf comObj.C_AssemblyFeatures41 = "Yes" Then
                    row1(37) = Yes
                Else
                    row1(37) = comObj.C_AssemblyFeatures41
                End If

                If comObj.C_ModelPreview42 = "No" Then
                    row1(38) = No
                    count += 1
                ElseIf comObj.C_ModelPreview42 = "Yes" Then
                    row1(38) = Yes
                Else
                    row1(38) = comObj.C_ModelPreview42
                End If

                If comObj.C_PartConstraint43 = "No" Then
                    row1(39) = No
                    count += 1
                ElseIf comObj.C_PartConstraint43 = "Yes" Then
                    row1(39) = Yes
                ElseIf comObj.C_PartConstraint43 = "0" Then
                    row1(39) = No
                    count += 1
                Else
                    row1(39) = Yes
                    'row1(39) = comObj.C_PartConstraint43
                End If

                If comObj.C_Material_Spec_Field44 = "No" Then
                    row1(40) = No
                    count += 1
                ElseIf comObj.C_Material_Spec_Field44 = "Yes" Then
                    row1(40) = Yes
                Else
                    row1(40) = comObj.C_Material_Spec_Field44
                End If

                If comObj.C_MassAndDensity_Update45 = "No" Then
                    row1(41) = No
                    count += 1
                ElseIf comObj.C_MassAndDensity_Update45 = "Yes" Then
                    row1(41) = Yes
                Else
                    row1(41) = comObj.C_MassAndDensity_Update45
                End If

                If comObj.C_Bend_Radius_Update46 = "No" Then
                    row1(42) = No
                    count += 1
                ElseIf comObj.C_Bend_Radius_Update46 = "Yes" Then
                    row1(42) = Yes
                Else
                    row1(42) = comObj.C_Bend_Radius_Update46
                End If

                If comObj.C_FlatPattern_Update47 = "No" Then
                    row1(43) = No
                    count += 1
                ElseIf comObj.C_FlatPattern_Update47 = "Yes" Then
                    row1(43) = Yes
                Else
                    row1(43) = comObj.C_FlatPattern_Update47
                End If

                If comObj.C_HardwareInstancesAndStackups48 = "No" Then
                    row1(44) = No
                    count += 1
                ElseIf comObj.C_HardwareInstancesAndStackups48 = "Yes" Then
                    row1(44) = Yes
                Else
                    row1(44) = comObj.C_HardwareInstancesAndStackups48
                End If

                If comObj.C_DimensionalGeometry_Match_For_ALL49 = "No" Then
                    row1(45) = No
                    count += 1
                ElseIf comObj.C_DimensionalGeometry_Match_For_ALL49 = "Yes" Then
                    row1(45) = Yes
                Else
                    row1(45) = comObj.C_DimensionalGeometry_Match_For_ALL49
                End If

                If comObj.C_DimensionalGeometry_Match50 = "No" Then
                    row1(46) = No
                    count += 1
                ElseIf comObj.C_DimensionalGeometry_Match50 = "Yes" Then
                    row1(46) = Yes
                Else
                    row1(46) = comObj.C_DimensionalGeometry_Match50
                End If

                If comObj.C_suppressedUnusedFeatures51 = "No" Then
                    row1(47) = No
                    count += 1
                ElseIf comObj.C_suppressedUnusedFeatures51 = "Yes" Then
                    row1(47) = Yes
                Else
                    row1(47) = comObj.C_suppressedUnusedFeatures51
                End If

                If comObj.C_ChildComponentsFullyConstrained52 = "No" Then
                    row1(48) = No
                    count += 1
                ElseIf comObj.C_ChildComponentsFullyConstrained52 = "Yes" Then
                    row1(48) = Yes
                Else
                    row1(48) = comObj.C_ChildComponentsFullyConstrained52
                End If

                If comObj.C_SimplifiedAssemblyModel53 = "No" Then
                    row1(49) = No
                    count += 1
                ElseIf comObj.C_SimplifiedAssemblyModel53 = "Yes" Then
                    row1(49) = Yes
                Else
                    row1(49) = comObj.C_SimplifiedAssemblyModel53
                End If

                If comObj.C_ChildPartOccurrenceProperties54 = "No" Then
                    row1(50) = No
                    count += 1
                ElseIf comObj.C_ChildPartOccurrenceProperties54 = "Yes" Then
                    row1(50) = Yes
                Else
                    row1(50) = comObj.C_ChildPartOccurrenceProperties54
                End If

                If comObj.C_VendorMaterialData55 = "No" Then
                    row1(51) = No
                    count += 1
                ElseIf comObj.C_VendorMaterialData55 = "Yes" Then
                    row1(51) = Yes
                Else
                    row1(51) = comObj.C_VendorMaterialData55
                End If

                If comObj.C_BaselineMassAndDensityUpdate56 = "No" Then
                    row1(52) = No
                    count += 1
                ElseIf comObj.C_BaselineMassAndDensityUpdate56 = "Yes" Then
                    row1(52) = Yes
                Else
                    row1(52) = comObj.C_BaselineMassAndDensityUpdate56
                End If

                If comObj.C_PMI_Instruction57 = "No" Then
                    row1(53) = No
                    count += 1
                ElseIf comObj.C_PMI_Instruction57 = "Yes" Then
                    row1(53) = Yes
                Else
                    row1(53) = comObj.C_PMI_Instruction57
                End If

                If comObj.C_BinderData58 = "No" Then
                    row1(54) = No
                    count += 1
                ElseIf comObj.C_BinderData58 = "Yes" Then
                    row1(54) = Yes
                Else
                    row1(54) = comObj.C_BinderData58
                End If

                If comObj.C_TerminalAssigned59 = "No" Then
                    row1(55) = No
                    count += 1
                ElseIf comObj.C_TerminalAssigned59 = "Yes" Then
                    row1(55) = Yes
                Else
                    row1(55) = comObj.C_TerminalAssigned59
                End If

                If comObj.C_ReferencePlanes60 = "No" Then
                    row1(56) = No
                    count += 1
                ElseIf comObj.C_ReferencePlanes60 = "Yes" Then
                    row1(56) = Yes
                Else
                    row1(56) = comObj.C_ReferencePlanes60
                End If

                If comObj.R_ConstrainedRefPlane61 = "No" Then
                    row1(57) = No
                    count += 1
                ElseIf comObj.R_ConstrainedRefPlane61 = "Yes" Then
                    row1(57) = Yes
                Else
                    row1(57) = comObj.R_ConstrainedRefPlane61
                End If

                If comObj.R_FullyConstrained62 = "No" Then
                    row1(58) = No
                    count += 1
                ElseIf comObj.R_FullyConstrained62 = "Yes" Then
                    row1(58) = Yes
                Else
                    row1(58) = comObj.R_FullyConstrained62
                End If

                If comObj.R_FullyConstrainedAndGround63 = "No" Then
                    row1(59) = No
                    count += 1
                ElseIf comObj.R_FullyConstrainedAndGround63 = "Yes" Then
                    row1(59) = Yes
                Else
                    row1(59) = comObj.R_FullyConstrainedAndGround63
                End If

                If comObj.R_ProperConstrained64 = "No" Then
                    row1(60) = No
                    count += 1
                ElseIf comObj.R_ProperConstrained64 = "Yes" Then
                    row1(60) = Yes
                Else
                    row1(60) = comObj.R_ProperConstrained64
                End If

                'needsimplification no=0,yes=1
                If comObj.R_NeedSimplified65 = "No" Then
                    row1(61) = Yes
                ElseIf comObj.R_NeedSimplified65 = "Yes" Then
                    row1(61) = No
                    count += 1
                Else
                    row1(61) = comObj.R_NeedSimplified65
                End If

                If comObj.R_PreviewModelGEE66 = "No" Then
                    row1(62) = No
                    count += 1
                ElseIf comObj.R_PreviewModelGEE66 = "Yes" Then
                    row1(62) = Yes
                Else
                    row1(62) = comObj.R_PreviewModelGEE66
                End If

                If comObj.R_WeightMass67 = "No" Then
                    row1(63) = No
                    count += 1
                ElseIf comObj.R_WeightMass67 = "Yes" Then
                    row1(63) = Yes
                Else
                    row1(63) = comObj.R_WeightMass67
                End If

                If comObj.R_OccurrenceProperties68 = "No" Then
                    row1(64) = No
                    count += 1
                ElseIf comObj.R_OccurrenceProperties68 = "Yes" Then
                    row1(64) = Yes
                Else
                    row1(64) = comObj.R_OccurrenceProperties68
                End If

                If comObj.R_PatternsWithinPatterns69 = "No" Then
                    row1(65) = No
                    count += 1
                ElseIf comObj.R_PatternsWithinPatterns69 = "Yes" Then
                    row1(65) = Yes
                Else
                    row1(65) = comObj.R_PatternsWithinPatterns69
                End If

                'R_ExternallySuppliedModels70 no=0,yes=1
                If comObj.R_ExternallySuppliedModels70 = "No" Then
                    row1(66) = Yes
                ElseIf comObj.R_ExternallySuppliedModels70 = "Yes" Then
                    row1(66) = No
                    count += 1
                Else
                    row1(66) = comObj.R_ExternallySuppliedModels70
                End If

                If comObj.R_holePatterns71 = "No" Then
                    row1(67) = No
                    count += 1
                ElseIf comObj.R_holePatterns71 = "Yes" Then
                    row1(67) = Yes
                Else
                    row1(67) = comObj.R_holePatterns71
                End If

                If comObj.R_HoleDiameters72 = "No" Then
                    row1(68) = No
                    count += 1
                ElseIf comObj.R_HoleDiameters72 = "Yes" Then
                    row1(68) = Yes
                Else
                    row1(68) = comObj.R_HoleDiameters72
                End If

                If comObj.R_HoleClearances73 = "No" Then
                    row1(69) = No
                    count += 1
                ElseIf comObj.R_HoleClearances73 = "Yes" Then
                    row1(69) = Yes
                Else
                    row1(69) = comObj.R_HoleClearances73
                End If

                If comObj.R_HoleDistance74 = "No" Then
                    row1(70) = No
                    count += 1
                ElseIf comObj.R_HoleDistance74 = "Yes" Then
                    row1(70) = Yes
                Else
                    row1(70) = comObj.R_HoleDistance74
                End If

                If comObj.R_PartLineupWithMatingParts75 = "No" Then
                    row1(71) = No
                    count += 1
                ElseIf comObj.R_PartLineupWithMatingParts75 = "Yes" Then
                    row1(71) = Yes
                Else
                    row1(71) = comObj.R_PartLineupWithMatingParts75
                End If

                If comObj.R_FITS_In_MatingParts76 = "No" Then
                    row1(72) = No
                    count += 1
                ElseIf comObj.R_FITS_In_MatingParts76 = "Yes" Then
                    row1(72) = Yes
                Else
                    row1(72) = comObj.R_FITS_In_MatingParts76
                End If

                If comObj.R_Density77 = "No" Then
                    row1(73) = No
                    count += 1
                ElseIf comObj.R_Density77 = "Yes" Then
                    row1(73) = Yes
                Else
                    row1(73) = comObj.R_Density77
                End If

                If comObj.R_Material = "No" Then
                    row1(74) = No
                    count += 1
                ElseIf comObj.R_Material = "Yes" Then
                    row1(74) = Yes
                Else
                    row1(74) = comObj.R_Material
                End If

                If comObj.R_UnusedSketches78 = "No" Then
                    row1(75) = No
                    count += 1
                ElseIf comObj.R_UnusedSketches78 = "Yes" Then
                    row1(75) = Yes
                Else
                    row1(75) = comObj.R_UnusedSketches78
                End If

                If comObj.R_linkSketches79 = "No" Then
                    row1(76) = No
                    count += 1
                ElseIf comObj.R_linkSketches79 = "Yes" Then
                    row1(76) = Yes
                Else
                    row1(76) = comObj.R_linkSketches79
                End If

                If comObj.R_BinderProcess80 = "No" Then
                    row1(77) = No
                    count += 1
                ElseIf comObj.R_BinderProcess80 = "Yes" Then
                    row1(77) = Yes
                Else
                    row1(77) = comObj.R_BinderProcess80
                End If

                If comObj.R_BendRadius81 = "No" Then
                    row1(78) = No
                    count += 1
                ElseIf comObj.R_BendRadius81 = "Yes" Then
                    row1(78) = Yes
                Else
                    row1(78) = comObj.R_BendRadius81
                End If
                '-----------------------------------------------------------------------------
                '77 density -> material

                row1(79) = count
                row1(80) = comObj.date38

                If comObj.lastModifiedDate39 = "No" Then
                    row1(81) = "0"

                Else
                    row1(81) = comObj.lastModifiedDate39
                End If

                count = 0


#End Region
            End If

        Catch ex As Exception
            'MsgBox(ex.Message + ex.StackTrace)
        End Try

    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    Public Sub killProcess()
        Dim _proceses As Process()
        _proceses = Process.GetProcessesByName("excel")
        For Each proces As Process In _proceses
            proces.Kill()
        Next
    End Sub
End Class
