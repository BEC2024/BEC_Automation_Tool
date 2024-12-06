
Public Class MTC_ReportExcelUtil
    Implements IMTC_ReportExcelUtil

    Dim dt As New DataTable()
    Dim Projectname As String = Nothing
    Dim er As String = "No"
    Dim notexist As String = "0"
    Dim No As String = "1"
    Dim Yes As String = "0"
    Public Sub MTCReport(mtcObj As MTC_Report) Implements IMTC_ReportExcelUtil.MTCReport
        Try
            If Not dt.Columns.Count > 0 Then
                'heading
                dt.Columns.Add(New DataColumn("Part Name")) '1
                dt.Columns.Add(New DataColumn("Category")) '2
                dt.Columns.Add(New DataColumn("Part number match")) '3
                'dt.Columns.Add(New DataColumn("Rivision level"))'4
                dt.Columns.Add(New DataColumn("Autor Name")) '4
                dt.Columns.Add(New DataColumn("Project Name")) '5
                dt.Columns.Add(New DataColumn("Document number")) '6
                dt.Columns.Add(New DataColumn("Autor name")) '7
                dt.Columns.Add(New DataColumn("Dash in unused field")) '8
                dt.Columns.Add(New DataColumn("M2M Description match")) '9
                dt.Columns.Add(New DataColumn("UOMs match")) '10
                dt.Columns.Add(New DataColumn("Interfenrce ")) '11
                dt.Columns.Add(New DataColumn("inter part copies")) '12
                dt.Columns.Add(New DataColumn("part copies")) '13
                dt.Columns.Add(New DataColumn("broken file")) '14
                dt.Columns.Add(New DataColumn("Adjustable")) '15 till assembly
                dt.Columns.Add(New DataColumn("Mat'l spec")) '16
                dt.Columns.Add(New DataColumn("Material used")) '17
                dt.Columns.Add(New DataColumn("Removed Unused Features")) '18
                dt.Columns.Add(New DataColumn("ASTM minimum")) '19
                dt.Columns.Add(New DataColumn("Flat Pattern")) '20
                dt.Columns.Add(New DataColumn(" Hole Tool Use")) '21
                ' dt.Columns.Add(New DataColumn("Revision Number correct")) '23
                'dt.Columns.Add(New DataColumn("Component Name")) '22
                dt.Columns.Add(New DataColumn("Virtual Thread")) '22
                dt.Columns.Add(New DataColumn("Defined Feature")) '23
                dt.Columns.Add(New DataColumn("Supprresed Feature")) '24
                dt.Columns.Add(New DataColumn("Vendor Part Number")) '25
                dt.Columns.Add(New DataColumn("Hardware Parts")) '26
                dt.Columns.Add(New DataColumn("M2M Sourced Marked")) '27
                dt.Columns.Add(New DataColumn("SE Status")) '28
                dt.Columns.Add(New DataColumn("Total")) '29
                dt.Columns.Add(New DataColumn("Report Date")) '30
            End If

            'function call
            Assembly(mtcObj)
            Part(mtcObj)
            Sheetmetal(mtcObj)
            Baseline(mtcObj)
            Electrical(mtcObj)
            creteMTCExcel(mtcObj)



        Catch ex As Exception
            MsgBox(ex.Message, ex.StackTrace)
        End Try

#Region "Comment"
        ''data table title
        'dt.Columns.Add(New DataColumn("Author"))
        'dt.Columns.Add(New DataColumn("Category"))
        'dt.Columns.Add(New DataColumn("Part"))
        'dt.Columns.Add(New DataColumn("Projectname"))
        'dt.Columns.Add(New DataColumn("Error Count"))


        'Dim sheetName As ArrayList = New ArrayList()

        'sheetName.Add("Assembly")
        'sheetName.Add("Part")
        'sheetName.Add("Sheetmetal")
        'sheetName.Add("Baseline")
        'sheetName.Add("Electrical")



        'Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
        'Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
        'xlsWB = xlsApp.Workbooks.Open(excelFilePath, Nothing, False)

        'Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        'Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing
        'For c = 0 To sheetName.Count - 1


        '    xlsSheet1 = xlsWB.Sheets(sheetName(c))
        '    xlsSheet1.Activate()
        '    Dim RowCnt = xlsSheet1.UsedRange.Rows.Count

        '    Dim ColCnt = xlsSheet1.UsedRange.Columns.Count
        '    Dim range1 As Microsoft.Office.Interop.Excel.Range

        '    xlsCell1 = xlsSheet1.UsedRange()

        '    Dim ans = 0
        '    Dim count = Nothing

        '    For i = 4 To ColCnt
        '        'initialization
        '        Dim Aunthor As String = Nothing
        '        Dim category As String = Nothing
        '        Dim projectName As String = Nothing
        '        Dim part = xlsSheet1.Rows(1).Cells(i).value.ToString()
        '        'MsgBox(xlsSheet1.Rows(1).Cells(i).value.ToString)
        '        For j = 3 To RowCnt
        '            If (xlsSheet1.Rows(j).Cells(i).value = Nothing) Then

        '                count = j - 1
        '                Exit For
        '            End If
        '        Next
        '        Dim row1 As DataRow = dt.NewRow()
        '        category = sheetName(c)

        '        For k = 3 To count
        '            If category = "Baseline" Then
        '                If k = 4 Then
        '                    Aunthor = xlsSheet1.Rows(k).Cells(i).value.ToString()
        '                End If

        '                If k = 17 Then

        '                    Dim projectNameStr() As String
        '                    Dim Details As String
        '                    If xlsSheet1.Rows(k).Cells(i).Comment.Text() Is Nothing Then
        '                        Details = "NA"
        '                    End If
        '                    Details = xlsSheet1.Rows(k).Cells(i).Comment.Text()
        '                    projectNameStr = Details.Split(New Char() {":"c})
        '                    projectName = projectNameStr(1)

        '                End If

        '                If (xlsSheet1.Rows(k).Cells(i).value.ToString = "No") Then
        '                    ans = ans + 1


        '                End If

        '            Else
        '                If k = 5 Then
        '                    Aunthor = xlsSheet1.Rows(k).Cells(i).value.ToString()
        '                    ' dt.Rows.Add(xlsSheet1.Rows(k).Cells(i).value.ToString())
        '                End If

        '                If k = 6 Then

        '                    Dim projectNameStr() As String
        '                    Dim Details As String
        '                    Details = xlsSheet1.Rows(k).Cells(i).Comment.Text()
        '                    projectNameStr = Details.Split(New Char() {":"c})
        '                    projectName = projectNameStr(1)
        '                    'MsgBox(projectName)

        '                    'dt.Rows.Add(projectName)
        '                End If

        '                'MsgBox(xlsSheet1.Rows(k).Cells(i).value.ToString())
        '                If (xlsSheet1.Rows(k).Cells(i).value.ToString = "No") Then
        '                    ans = ans + 1


        '                End If
        '            End If



        '        Next
        '        row1(0) = Aunthor
        '        row1(1) = category
        '        row1(2) = part
        '        row1(3) = projectName
        '        row1(4) = ans
        '        dt.Rows.Add(row1)
        '        'dt.Rows.Add(ans)
        '        ' MsgBox(xlsSheet1.Cells(24, i).value.ToString())
        '        ans = 0

        '    Next





        '    'row title
        '    'Dim ColumnName = "NO"
        '    'If Not xlsSheet1.Cells(RowCnt, 3).value.ToString() = ColumnName Then
        '    '    xlsSheet1.Cells(RowCnt + 1, 3) = ColumnName
        '    'End If


        '    'range1 = xlsSheet1.Range("D" & 2 & ":D" & 23 & "")
        '    'Dim ans = xlsApp.WorksheetFunction.CountIf(range1, "No")
        '    'xlsSheet1.Cells(24, 4) = ans





        'Next
        'xlsWB.Close()
        'xlsApp.Quit()
        'releaseObject(xlsApp)
        'releaseObject(xlsWB)
        'releaseObject(xlsSheet1)
        'killProcess()
#End Region
        'connecting to excel application
        'Dim xlApp As New Microsoft.Office.Interop.Excel.Application
        ''connecting workbook and excel sheet
        'Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook

        'Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        'Dim misValue As Object = System.Reflection.Missing.Value


        'xlWorkBook = xlApp.Workbooks.Add(Type.Missing)
        'xlWorkSheet = xlWorkBook.ActiveSheet
        'xlWorkSheet.Name = "MTCReport"

        'xlWorkSheet.ImportDataTable(dt, True, 1, 1)
        'xlWorkSheet.UsedRange.AutofitColumns()
        'xlWorkBook.SaveAs("C:\Users\pratikg\Desktop\sunny\MTCReport.xlsx")
        'xlWorkBook.Close()
        'xlApp.Quit()





        'Return dt

        'xlWorkBook.Close()
        'xlApp.Quit()
        'releaseObject(xlApp)
        'releaseObject(xlWorkBook)
        'releaseObject(xlWorkSheet)
        'killProcess()


    End Sub

    Public Sub creteMTCExcel(ByVal mtcObj As MTC_Report) Implements IMTC_ReportExcelUtil.CreteMTCExce
        Dim _excel As New Microsoft.Office.Interop.Excel.Application
        Dim wBook As Microsoft.Office.Interop.Excel.Workbook
        Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet

        wBook = _excel.Workbooks.Add()
        wSheet = wBook.ActiveSheet()
        wSheet.Name = "MTC Report"



        Dim dc As System.Data.DataColumn
        Dim dr As System.Data.DataRow
        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = 0

        For Each dc In dt.Columns
            colIndex += 1
            _excel.Cells(1, colIndex) = dc.ColumnName
        Next
        For Each dr In dt.Rows
            rowIndex += 1
            colIndex = 0
            For Each dc In dt.Columns
                colIndex += 1
                _excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
            Next
        Next

        Dim formatRange As Microsoft.Office.Interop.Excel.Range
        wSheet.Columns.AutoFit()
        formatRange = wSheet.Range("a1", "ad1")

        formatRange.EntireRow.Font.Bold = True
        formatRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        formatRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
        formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SandyBrown)

        formatRange = wSheet.Range("a1", "ad1")
        formatRange.RowHeight = 20
        wSheet.Range("A1:AD1").Columns.EntireColumn.AutoFit()
        formatRange.Font.Name = "Arial"
        formatRange.Font.Size = 11
        Dim str2 = wSheet.UsedRange.Rows.Count + 1
        For i = 2 To str2
            formatRange = wSheet.Range("a" & i & ":ad" & i & "")
            formatRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            formatRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
            formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow)
            wSheet.Columns.AutoFit()

        Next

        Dim strFileName As String = mtcObj.dir + "\MTC_MTR_KPI_Report.xlsx"
        If System.IO.File.Exists(strFileName) Then
            System.IO.File.Delete(strFileName)
        End If


        If mtcObj.i = mtcObj.files.Count - 1 Then
            wBook.SaveAs(strFileName)
            wBook.Close()
            _excel.Quit()

        End If

    End Sub
    Public Sub Assembly(ByVal mtcObj As MTC_Report) Implements IMTC_ReportExcelUtil.Assembly
        Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
        xlsWB = xlsApp.Workbooks.Open(mtcObj.excelFilePath, Nothing, False)

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

            '1PartName Value(Error)
            mtcObj.partName1 = xlsSheet1.Rows(1).Cells(i).value.ToString()


            '2part number match value/Is the part number match with M2M? *
            mtcObj.partNumberMatch2 = If(xlsSheet1.Rows(3).Cells(i).value = Nothing, er, xlsSheet1.Rows(3).Cells(i).value.ToString())


            'Category
            mtcObj.category = "Assembly"

            '3Rivisionlevel/What is the revision level? *
            'If xlsSheet1.Rows(4).Cells(i).value = Nothing Then
            '    mtcObj.revisonlevel3 = "Null"
            'Else
            '    mtcObj.revisonlevel3 = xlsSheet1.Rows(4).Cells(i).value.ToString()
            'End If


            '4author name/Who is the Author of the file?
            mtcObj.AuthorName4 = xlsSheet1.Rows(5).Cells(i).value.ToString()


            '5projectname value/'  Is the correct "Project" used in the "Project Name" field? (Ensure consistent project titleIs used within a project) *

            mtcObj.projectName5 = If(xlsSheet1.Rows(1).Cells(4).value = Nothing, er, xlsSheet1.Rows(1).Cells(4).value.ToString())
            Projectname = mtcObj.projectName5

            '6Is the Revision number correct? *
            'If xlsSheet1.Rows(7).Cells(i).value = Nothing Then
            '    mtcObj.revisoncorrect6 = "Null"
            'Else
            '    mtcObj.revisoncorrect6 = xlsSheet1.Rows(7).Cells(i).value.ToString()
            'End If



            '7document number/Is the Document Number correct? (BEC designed parts must have proper doc number.Hardware And other exceptions are N/A) *
            mtcObj.documentNumber7 = If(xlsSheet1.Rows(8).Cells(i).value = Nothing, er, xlsSheet1.Rows(8).Cells(i).value.ToString())



            '8authorNameIs the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) 
            mtcObj.AuthorName8 = If(xlsSheet1.Rows(9).Cells(i).value = Nothing, er, xlsSheet1.Rows(9).Cells(i).value.ToString())


            '9dashinunusedfield/Do all technically unused properties have a "dash" populated? 
            mtcObj.DashInUnusedFiled9 = If(xlsSheet1.Rows(10).Cells(i).value = Nothing, er, xlsSheet1.Rows(10).Cells(i).value.ToString())

            '10 m2mdescription/Does the Model Title MATCH the M2M Item Master (INV) Description Field?
            mtcObj.m2mDiscriptionMatch10 = If(xlsSheet1.Rows(11).Cells(i).value = Nothing, er, xlsSheet1.Rows(11).Cells(i).value.ToString())


            '11uomsmatch/Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
            mtcObj.uomsMatch11 = If(xlsSheet1.Rows(12).Cells(i).value = Nothing, er, xlsSheet1.Rows(12).Cells(i).value.ToString())



            '12interface/Is interfernces found in assembly?
            mtcObj.interface12 = If(xlsSheet1.Rows(14).Cells(i).value = Nothing, er, xlsSheet1.Rows(14).Cells(i).value.ToString())


            '13interpartcopiesIs interfernces found in assembly?
            mtcObj.interpartcopies13 = If(xlsSheet1.Rows(15).Cells(i).value = Nothing, er, xlsSheet1.Rows(15).Cells(i).value.ToString())


            '14partcopiespart copies detected
            mtcObj.partcopies14 = If(xlsSheet1.Rows(16).Cells(i).value = Nothing, er, xlsSheet1.Rows(16).Cells(i).value.ToString())


            '15brokenfilebroken  file Path detected
            mtcObj.brokefile15 = If(xlsSheet1.Rows(17).Cells(i).value = Nothing, er, xlsSheet1.Rows(17).Cells(i).value.ToString())


            '16Adjustable/Is the part "Adjustable"? (part should NOT be adjustable) *
            mtcObj.adjustable16 = If(xlsSheet1.Rows(18).Cells(i).value = Nothing, er, xlsSheet1.Rows(17).Cells(i).value.ToString())


            '17Mat'l spec/Is the mat'l spec field populated? (indicated by RED in metadata audit file)
            mtcObj.Matl_spec17 = notexist

            '18MaterialUsed/Is the material used field populated? *
            mtcObj.MaterialUsed18 = notexist

            '19RemovedUnusedFeatures/Have any suppressed (unused) features been removed from the model Pathfinder? *
            mtcObj.RemovedUnusedFeatures19 = notexist


            '20ASTMminimum/Is the bend radius of the part equal to or above the ASTM minimum? *
            mtcObj.ASTMminimum20 = notexist

            '21FlatPattern/. Is the flat pattern turned ON? (Applicable only to parts needing dxf or cut length) *
            mtcObj.FlatPattern21 = notexist

            '22HoleToolUse/Did I use the hole tool for ALL holes requiring fasteners? (include clearance holes forhardware, tapped holes, And slots) *
            mtcObj.HoleToolUse22 = notexist



            '23component Name
            'mtcObj.componentName23 = er

            '24virtualThread/Virtual thread applied for Fasteners?
            mtcObj.VirtualThread24 = notexist

            '25defined feature
            mtcObj.defineFeatures25 = notexist

            '26Supperessed Feature
            mtcObj.suppressedFeature26 = notexist

            '27Vendorpartnumber
            mtcObj.vendorPartNumber27 = notexist

            '28Hardwareparts
            mtcObj.HardwareParts28 = notexist

            '28m2mSourceMarked
            mtcObj.m2mSourceMarked29 = notexist

            '29SEStatus
            mtcObj.SEstatus30 = notexist

            '30 date
            mtcObj.date31 = Date.Now()
            '------------------------------------------------------------------
            Row1Data(mtcObj)

        Next


        xlsWB.Close()
        xlsApp.Quit()
        ReleaseObject(xlsApp)
        ReleaseObject(xlsWB)
        ReleaseObject(xlsSheet1)
        KillProcess()

    End Sub

    Public Sub Part(ByVal mtcObj As MTC_Report) Implements IMTC_ReportExcelUtil.Part
        Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
        xlsWB = xlsApp.Workbooks.Open(mtcObj.excelFilePath, Nothing, False)

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

            '1PartName Value(Error)
            mtcObj.partName1 = xlsSheet1.Rows(1).Cells(i).value.ToString()

            '2part number match value/Is the part number match with M2M? *
            mtcObj.partNumberMatch2 = If(xlsSheet1.Rows(3).Cells(i).value = Nothing, er, xlsSheet1.Rows(3).Cells(i).value.ToString())


            'Category
            mtcObj.category = xlsSheet1.Name

            '4author name/Who is the Author of the file?
            mtcObj.AuthorName4 = xlsSheet1.Rows(5).Cells(i).value.ToString()

            '5projectname value/'  Is the correct "Project" used in the "Project Name" field? (Ensure consistent project titleIs used within a project) *
            'from assembly sheet first cell name(variable ProjectName)


            '6document number/Is the Document Number correct? (BEC designed parts must have proper doc number.Hardware And other exceptions are N/A) *
            mtcObj.documentNumber7 = If(xlsSheet1.Rows(9).Cells(i).value = Nothing, er, xlsSheet1.Rows(9).Cells(i).value.ToString())


            '8authorNameIs the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) 
            mtcObj.AuthorName8 = If(xlsSheet1.Rows(6).Cells(i).value = Nothing, er, xlsSheet1.Rows(6).Cells(i).value.ToString())

            '9dashinunusedfield/Do all technically unused properties have a "dash" populated? 
            mtcObj.DashInUnusedFiled9 = If(xlsSheet1.Rows(10).Cells(i).value = Nothing, er, xlsSheet1.Rows(10).Cells(i).value.ToString())


            '10 m2mdescription/Does the Model Title MATCH the M2M Item Master (INV) Description Field?
            mtcObj.m2mDiscriptionMatch10 = If(xlsSheet1.Rows(11).Cells(i).value = Nothing, er, xlsSheet1.Rows(11).Cells(i).value.ToString())


            '11uomsmatch/Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
            mtcObj.uomsMatch11 = If(xlsSheet1.Rows(12).Cells(i).value = Nothing, er, xlsSheet1.Rows(12).Cells(i).value.ToString)



            '12interface/Is interfernces found in assembly?
            mtcObj.interface12 = notexist


            '13interpartcopiesIs interfernces found in assembly?
            mtcObj.interpartcopies13 = notexist


            '14partcopiespart copies detected
            mtcObj.partcopies14 = notexist


            '15brokenfilebroken  file Path detected
            mtcObj.brokefile15 = notexist

            '16Adjustable/Is the part "Adjustable"? (part should NOT be adjustable) *
            mtcObj.adjustable16 = If(xlsSheet1.Rows(18).Cells(i).value = Nothing, er, xlsSheet1.Rows(17).Cells(i).value.ToString())



            '17Mat'l spec/Is the mat'l spec field populated? (indicated by RED in metadata audit file)
            mtcObj.Matl_spec17 = If(xlsSheet1.Rows(13).Cells(i).value = Nothing, er, xlsSheet1.Rows(13).Cells(i).value.ToString)


            '18MaterialUsed/Is the material used field populated? *
            mtcObj.MaterialUsed18 = If(xlsSheet1.Rows(14).Cells(i).value = Nothing, er, xlsSheet1.Rows(14).Cells(i).value.ToString)

            '19RemovedUnusedFeatures/Have any suppressed (unused) features been removed from the model Pathfinder? *
            mtcObj.RemovedUnusedFeatures19 = If(xlsSheet1.Rows(17).Cells(i).value = Nothing, er, xlsSheet1.Rows(17).Cells(i).value.ToString)

            '20ASTMminimum/Is the bend radius of the part equal to or above the ASTM minimum? *
            mtcObj.ASTMminimum20 = notexist

            '21FlatPattern/. Is the flat pattern turned ON? (Applicable only to parts needing dxf or cut length) *
            mtcObj.FlatPattern21 = notexist

            '22HoleToolUse/Did I use the hole tool for ALL holes requiring fasteners? (include clearance holes forhardware, tapped holes, And slots) *
            mtcObj.HoleToolUse22 = notexist



            '23component Name
            'mtcObj.componentName23 = er

            '24virtualThread/Virtual thread applied for Fasteners?
            mtcObj.VirtualThread24 = notexist

            '25defined feature
            mtcObj.defineFeatures25 = notexist

            '26Supperessed Feature
            mtcObj.suppressedFeature26 = notexist

            '27Vendorpartnumber
            mtcObj.vendorPartNumber27 = notexist

            'Hardwareparts
            mtcObj.HardwareParts28 = notexist

            '28m2mSourceMarked
            mtcObj.m2mSourceMarked29 = notexist

            '28SEStatus
            mtcObj.SEstatus30 = notexist
            '----------------------------------------------------------------------------------------------------------
            Row1Data(mtcObj)

        Next



        xlsWB.Close()
        xlsApp.Quit()
        ReleaseObject(xlsApp)
        ReleaseObject(xlsWB)
        ReleaseObject(xlsSheet1)
        KillProcess()

    End Sub

    Public Sub Sheetmetal(ByVal mtcObj As MTC_Report) Implements IMTC_ReportExcelUtil.Sheetmetal
        Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
        xlsWB = xlsApp.Workbooks.Open(mtcObj.excelFilePath, Nothing, False)

        Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing



        xlsSheet1 = xlsWB.Sheets("Sheetmetal")
        xlsSheet1.Activate()
        Dim RowCnt = xlsSheet1.UsedRange.Rows.Count

        Dim ColCnt = xlsSheet1.UsedRange.Columns.Count


        xlsCell1 = xlsSheet1.UsedRange()

        Dim count = Nothing
        For i = 4 To ColCnt


            '1PartName Value(Error)
            mtcObj.partName1 = xlsSheet1.Rows(1).Cells(i).value.ToString()

            '2part number match value/Is the part number match with M2M? *
            mtcObj.partNumberMatch2 = If(xlsSheet1.Rows(3).Cells(i).value = Nothing, er, xlsSheet1.Rows(3).Cells(i).value.ToString())


            'Category
            mtcObj.category = xlsSheet1.Name

            '4author name/Who is the Author of the file?
            mtcObj.AuthorName4 = xlsSheet1.Rows(5).Cells(i).value.ToString()

            '5projectname value/'  Is the correct "Project" used in the "Project Name" field? (Ensure consistent project titleIs used within a project) *
            'from assembly sheet first cell name(variable ProjectName)


            '6document number/Is the Document Number correct? (BEC designed parts must have proper doc number.Hardware And other exceptions are N/A) *
            mtcObj.documentNumber7 = If(xlsSheet1.Rows(8).Cells(i).value = Nothing, er, xlsSheet1.Rows(9).Cells(i).value.ToString())


            '8authorNameIs the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) 
            mtcObj.AuthorName8 = If(xlsSheet1.Rows(9).Cells(i).value = Nothing, er, xlsSheet1.Rows(6).Cells(i).value.ToString())

            '9dashinunusedfield/Do all technically unused properties have a "dash" populated? 
            mtcObj.DashInUnusedFiled9 = If(xlsSheet1.Rows(10).Cells(i).value = Nothing, er, xlsSheet1.Rows(10).Cells(i).value.ToString())


            '10 m2mdescription/Does the Model Title MATCH the M2M Item Master (INV) Description Field?
            mtcObj.m2mDiscriptionMatch10 = If(xlsSheet1.Rows(11).Cells(i).value = Nothing, er, xlsSheet1.Rows(11).Cells(i).value.ToString())


            '11uomsmatch/Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
            mtcObj.uomsMatch11 = If(xlsSheet1.Rows(12).Cells(i).value = Nothing, er, xlsSheet1.Rows(12).Cells(i).value.ToString)



            '12interface/Is interfernces found in assembly?
            mtcObj.interface12 = notexist


            '13interpartcopiesIs interfernces found in assembly?
            mtcObj.interpartcopies13 = notexist


            '14partcopiespart copies detected
            mtcObj.partcopies14 = notexist


            '15brokenfilebroken  file Path detected
            mtcObj.brokefile15 = notexist

            '16Adjustable/Is the part "Adjustable"? (part should NOT be adjustable) *
            ' mtcObj.adjustable16 = "No"
            mtcObj.adjustable16 = If(xlsSheet1.Rows(18).Cells(i).value = Nothing, er, xlsSheet1.Rows(18).Cells(i).value.ToString())



            '17Mat'l spec/Is the mat'l spec field populated? (indicated by RED in metadata audit file)
            mtcObj.Matl_spec17 = If(xlsSheet1.Rows(13).Cells(i).value = Nothing, er, xlsSheet1.Rows(13).Cells(i).value.ToString)


            '18MaterialUsed/Is the material used field populated? *
            mtcObj.MaterialUsed18 = If(xlsSheet1.Rows(14).Cells(i).value = Nothing, er, xlsSheet1.Rows(14).Cells(i).value.ToString)

            '19RemovedUnusedFeatures/Have any suppressed (unused) features been removed from the model Pathfinder? *
            mtcObj.RemovedUnusedFeatures19 = notexist

            '20ASTMminimum/Is the bend radius of the part equal to or above the ASTM minimum? *
            mtcObj.ASTMminimum20 = If(xlsSheet1.Rows(15).Cells(i).value = Nothing, er, xlsSheet1.Rows(15).Cells(i).value.ToString)

            '21FlatPattern/. Is the flat pattern turned ON? (Applicable only to parts needing dxf or cut length) *
            mtcObj.FlatPattern21 = If(xlsSheet1.Rows(16).Cells(i).value = Nothing, er, xlsSheet1.Rows(16).Cells(i).value.ToString)

            '22HoleToolUse/Did I use the hole tool for ALL holes requiring fasteners? (include clearance holes forhardware, tapped holes, And slots) *
            mtcObj.HoleToolUse22 = If(xlsSheet1.Rows(17).Cells(i).value = Nothing, er, xlsSheet1.Rows(17).Cells(i).value.ToString)


            '23component Name
            'mtcObj.componentName23 = er

            '24virtualThread/Virtual thread applied for Fasteners?
            mtcObj.VirtualThread24 = notexist

            '25defined feature
            mtcObj.defineFeatures25 = notexist

            '26Supperessed Feature
            mtcObj.suppressedFeature26 = notexist

            '27Vendorpartnumber
            mtcObj.vendorPartNumber27 = notexist

            'Hardwareparts
            mtcObj.HardwareParts28 = notexist

            '28m2mSourceMarked
            mtcObj.m2mSourceMarked29 = notexist

            '29SEStatus
            mtcObj.SEstatus30 = notexist
            '----------------------------------------------------------------------------------------------------------

            Row1Data(mtcObj)



        Next


        xlsWB.Close()
        xlsApp.Quit()
        ReleaseObject(xlsApp)
        ReleaseObject(xlsWB)
        ReleaseObject(xlsSheet1)
        KillProcess()


    End Sub
    Public Sub Baseline(ByVal mtcObj As MTC_Report) Implements IMTC_ReportExcelUtil.Baseline
        Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
        xlsWB = xlsApp.Workbooks.Open(mtcObj.excelFilePath, Nothing, False)

        Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing



        xlsSheet1 = xlsWB.Sheets("Baseline")
        xlsSheet1.Activate()
        Dim RowCnt = xlsSheet1.UsedRange.Rows.Count

        Dim ColCnt = xlsSheet1.UsedRange.Columns.Count


        xlsCell1 = xlsSheet1.UsedRange()

        Dim count = Nothing
        For i = 4 To ColCnt


            '1PartName Value(Error)
            mtcObj.partName1 = xlsSheet1.Rows(1).Cells(i).value.ToString()

            'done
            '2part number match value/Is the part number match with M2M? *
            mtcObj.partNumberMatch2 = If(xlsSheet1.Rows(2).Cells(i).value = Nothing, er, xlsSheet1.Rows(2).Cells(i).value.ToString())


            'Category
            mtcObj.category = xlsSheet1.Name

            'done
            '4author name/Who is the Author of the file?
            mtcObj.AuthorName4 = xlsSheet1.Rows(4).Cells(i).value.ToString()

            '5projectname value/'  Is the correct "Project" used in the "Project Name" field? (Ensure consistent project titleIs used within a project) *
            'from assembly sheet first cell name(variable ProjectName)

            'done
            '6document number/Is the "Document Number" field populated with the correct part number? (This should MATCH the M2M Item Master Part Number field) *
            mtcObj.documentNumber7 = If(xlsSheet1.Rows(15).Cells(i).value = Nothing, er, xlsSheet1.Rows(15).Cells(i).value.ToString())

            'done
            '8authorNameIs the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) 
            mtcObj.AuthorName8 = If(xlsSheet1.Rows(12).Cells(i).value = Nothing, er, xlsSheet1.Rows(12).Cells(i).value.ToString())


            '9dashinunusedfield/Do all technically unused properties have a "dash" populated? 
            mtcObj.DashInUnusedFiled9 = If(xlsSheet1.Rows(19).Cells(i).value = Nothing, er, xlsSheet1.Rows(19).Cells(i).value.ToString())

            'here
            '10 m2mdescription/Does the Model Title MATCH the M2M Item Master (INV) Description Field?
            mtcObj.m2mDiscriptionMatch10 = If(xlsSheet1.Rows(11).Cells(i).value = Nothing, er, xlsSheet1.Rows(11).Cells(i).value.ToString())


            '11uomsmatch/Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
            mtcObj.uomsMatch11 = If(xlsSheet1.Rows(20).Cells(i).value = Nothing, er, xlsSheet1.Rows(20).Cells(i).value.ToString)



            '12interface/Is interfernces found in assembly?
            mtcObj.interface12 = notexist


            '13interpartcopiesIs interfernces found in assembly?
            mtcObj.interpartcopies13 = If(xlsSheet1.Rows(24).Cells(i).value = Nothing, er, xlsSheet1.Rows(24).Cells(i).value.ToString)


            '14partcopiespart copies detected
            mtcObj.partcopies14 = If(xlsSheet1.Rows(25).Cells(i).value = Nothing, er, xlsSheet1.Rows(25).Cells(i).value.ToString)


            '15brokenfilebroken  file Path detected
            mtcObj.brokefile15 = If(xlsSheet1.Rows(26).Cells(i).value = Nothing, er, xlsSheet1.Rows(26).Cells(i).value.ToString)

            'done
            '16Adjustable/Is the part "Adjustable"? (part should NOT be adjustable) *
            mtcObj.adjustable16 = If(xlsSheet1.Rows(27).Cells(i).value = Nothing, er, xlsSheet1.Rows(27).Cells(i).value.ToString())


            'here
            '17Mat'l spec/Is the "Mat'l Spec" field populated? (PURCHASED for library components) *
            mtcObj.Matl_spec17 = If(xlsSheet1.Rows(9).Cells(i).value = Nothing, er, xlsSheet1.Rows(9).Cells(i).value.ToString)


            '18MaterialUsed/Is the "Material Used" field populated? (PURCHASED for library components) *
            mtcObj.MaterialUsed18 = If(xlsSheet1.Rows(10).Cells(i).value = Nothing, er, xlsSheet1.Rows(10).Cells(i).value.ToString)

            '19RemovedUnusedFeatures/Have any suppressed (unused) features been removed from the model Pathfinder? *
            mtcObj.RemovedUnusedFeatures19 = If(xlsSheet1.Rows(35).Cells(i).value = Nothing, er, xlsSheet1.Rows(35).Cells(i).value.ToString)

            '20ASTMminimum/Is the bend radius of the part equal to or above the ASTM minimum? *
            mtcObj.ASTMminimum20 = notexist

            '21FlatPattern/. Is the flat pattern turned ON? (Applicable only to parts needing dxf or cut length) *
            mtcObj.FlatPattern21 = notexist

            '22HoleToolUse/Did I use the hole tool for ALL holes requiring fasteners? (include clearance holes forhardware, tapped holes, And slots) *
            mtcObj.HoleToolUse22 = If(xlsSheet1.Rows(22).Cells(i).value = Nothing, er, xlsSheet1.Rows(22).Cells(i).value.ToString)



            '23component Name/What type of component?
            'mtcObj.componentName23 = If(xlsSheet1.Rows(5).Cells(i).value = Nothing, er, xlsSheet1.Rows(5).Cells(i).value.ToString)

            '24virtualThread/Virtual thread applied for Fasteners?
            mtcObj.VirtualThread24 = If(xlsSheet1.Rows(6).Cells(i).value = Nothing, er, xlsSheet1.Rows(6).Cells(i).value.ToString)

            '25defined feature/Are ALL features fully defined in the model Pathfinder? * (Sketch fully constraint?)
            mtcObj.defineFeatures25 = If(xlsSheet1.Rows(7).Cells(i).value = Nothing, er, xlsSheet1.Rows(7).Cells(i).value.ToString)

            '26Supperessed Feature
            mtcObj.suppressedFeature26 = If(xlsSheet1.Rows(8).Cells(i).value = Nothing, er, xlsSheet1.Rows(8).Cells(i).value.ToString)

            '27Vendorpartnumber
            mtcObj.vendorPartNumber27 = If(xlsSheet1.Rows(14).Cells(i).value = Nothing, er, xlsSheet1.Rows(14).Cells(i).value.ToString)

            'Hardwareparts
            mtcObj.HardwareParts28 = If(xlsSheet1.Rows(18).Cells(i).value = Nothing, er, xlsSheet1.Rows(18).Cells(i).value.ToString)

            '28m2mSourceMarked
            mtcObj.m2mSourceMarked29 = If(xlsSheet1.Rows(21).Cells(i).value = Nothing, er, xlsSheet1.Rows(21).Cells(i).value.ToString)

            '29SEStatus
            mtcObj.SEstatus30 = If(xlsSheet1.Rows(29).Cells(i).value = Nothing, er, xlsSheet1.Rows(29).Cells(i).value.ToString)
            '----------------------------------------------------------------------------------------------------------

            Row1Data(mtcObj)



        Next


        xlsWB.Close()
        xlsApp.Quit()
        ReleaseObject(xlsApp)
        ReleaseObject(xlsWB)
        ReleaseObject(xlsSheet1)
        KillProcess()


    End Sub
    Public Sub Electrical(ByVal mtcObj As MTC_Report) Implements IMTC_ReportExcelUtil.Electrical
        Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
        xlsWB = xlsApp.Workbooks.Open(mtcObj.excelFilePath, Nothing, False)

        Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing



        xlsSheet1 = xlsWB.Sheets("Electrical")
        xlsSheet1.Activate()
        Dim RowCnt = xlsSheet1.UsedRange.Rows.Count

        Dim ColCnt = xlsSheet1.UsedRange.Columns.Count


        xlsCell1 = xlsSheet1.UsedRange()

        Dim count = Nothing
        For i = 4 To ColCnt


            '1PartName Value(Error)
            mtcObj.partName1 = xlsSheet1.Rows(1).Cells(i).value.ToString()

            'done
            '2part number match value/Is the part number match with M2M? *
            mtcObj.partNumberMatch2 = If(xlsSheet1.Rows(3).Cells(i).value = Nothing, er, xlsSheet1.Rows(3).Cells(i).value.ToString())


            'Category
            mtcObj.category = xlsSheet1.Name

            'done
            '4author name/Who is the Author of the file?
            mtcObj.AuthorName4 = xlsSheet1.Rows(5).Cells(i).value.ToString()

            '5projectname value/'  Is the correct "Project" used in the "Project Name" field? (Ensure consistent project titleIs used within a project) *
            'from assembly sheet first cell name(variable ProjectName)

            'done
            '6document number/Is the "Document Number" field populated with the correct part number? (This should MATCH the M2M Item Master Part Number field) *
            mtcObj.documentNumber7 = If(xlsSheet1.Rows(8).Cells(i).value = Nothing, er, xlsSheet1.Rows(8).Cells(i).value.ToString())

            'done
            '8authorNameIs the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) 
            mtcObj.AuthorName8 = If(xlsSheet1.Rows(9).Cells(i).value = Nothing, er, xlsSheet1.Rows(9).Cells(i).value.ToString())


            '9dashinunusedfield/Do all technically unused properties have a "dash" populated? 
            mtcObj.DashInUnusedFiled9 = If(xlsSheet1.Rows(10).Cells(i).value = Nothing, er, xlsSheet1.Rows(10).Cells(i).value.ToString())

            'here
            '10 m2mdescription/Does the Model Title MATCH the M2M Item Master (INV) Description Field?
            mtcObj.m2mDiscriptionMatch10 = If(xlsSheet1.Rows(11).Cells(i).value = Nothing, er, xlsSheet1.Rows(11).Cells(i).value.ToString())


            '11uomsmatch/Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
            mtcObj.uomsMatch11 = notexist



            '12interface/Is interfernces found in assembly?
            mtcObj.interface12 = notexist


            '13interpartcopiesIs interfernces found in assembly?
            mtcObj.interpartcopies13 = notexist


            '14partcopiespart copies detected
            mtcObj.partcopies14 = notexist


            '15brokenfilebroken  file Path detected
            mtcObj.brokefile15 = notexist

            'done
            '16Adjustable/Is the part "Adjustable"? (part should NOT be adjustable) *
            mtcObj.adjustable16 = notexist


            'here
            '17Mat'l spec/Is the "Mat'l Spec" field populated? (PURCHASED for library components) *
            mtcObj.Matl_spec17 = notexist


            '18MaterialUsed/Is the "Material Used" field populated? (PURCHASED for library components) *
            mtcObj.MaterialUsed18 = notexist

            '19RemovedUnusedFeatures/Have any suppressed (unused) features been removed from the model Pathfinder? *
            mtcObj.RemovedUnusedFeatures19 = notexist

            '20ASTMminimum/Is the bend radius of the part equal to or above the ASTM minimum? *
            mtcObj.ASTMminimum20 = notexist

            '21FlatPattern/. Is the flat pattern turned ON? (Applicable only to parts needing dxf or cut length) *
            mtcObj.FlatPattern21 = notexist

            '22HoleToolUse/Did I use the hole tool for ALL holes requiring fasteners? (include clearance holes forhardware, tapped holes, And slots) *
            mtcObj.HoleToolUse22 = notexist



            '23component Name/What type of component?
            'mtcObj.componentName23 = er

            '24virtualThread/Virtual thread applied for Fasteners?
            mtcObj.VirtualThread24 = notexist

            '25defined feature/Are ALL features fully defined in the model Pathfinder? * (Sketch fully constraint?)
            mtcObj.defineFeatures25 = notexist

            '26Supperessed Feature
            mtcObj.suppressedFeature26 = notexist

            '27Vendorpartnumber
            mtcObj.vendorPartNumber27 = notexist

            'Hardwareparts
            mtcObj.HardwareParts28 = notexist

            '28m2mSourceMarked
            mtcObj.m2mSourceMarked29 = notexist

            '29SEStatus
            mtcObj.SEstatus30 = notexist
            '----------------------------------------------------------------------------------------------------------

            Row1Data(mtcObj)



        Next


        xlsWB.Close()
        xlsApp.Quit()
        ReleaseObject(xlsApp)
        ReleaseObject(xlsWB)
        ReleaseObject(xlsSheet1)
        KillProcess()
    End Sub
    Public Sub Row1Data(ByVal mtcObj As MTC_Report) Implements IMTC_ReportExcelUtil.Row1Data
        Dim row1 As DataRow = dt.NewRow()
        dt.Rows.Add(row1)
        row1(0) = mtcObj.partName1
        row1(1) = mtcObj.category

        Dim count As Integer = 0
        '2
        If mtcObj.partNumberMatch2 = "No" Then
            row1(2) = No
            count += 1
        ElseIf mtcObj.partNumberMatch2 = "Yes" Then
            row1(2) = Yes
        Else
            row1(2) = mtcObj.partNumberMatch2
        End If


        'row1(3) = mtcObj.revisonlevel3
        row1(3) = mtcObj.AuthorName4
        row1(4) = Projectname
        'row1(6) = mtcObj.revisoncorrect6
        '5
        If mtcObj.documentNumber7 = "No" Then
            row1(5) = No
            count += 1
        ElseIf mtcObj.documentNumber7 = "Yes" Then
            row1(5) = Yes
        Else
            row1(5) = mtcObj.documentNumber7
        End If
        '6
        If mtcObj.AuthorName8 = "No" Then
            row1(6) = No
            count += 1
        ElseIf mtcObj.AuthorName8 = "Yes" Then
            row1(6) = Yes
        Else
            row1(6) = mtcObj.AuthorName8
        End If
        '7
        If mtcObj.DashInUnusedFiled9 = "No" Then
            row1(7) = No
            count += 1
        ElseIf mtcObj.DashInUnusedFiled9 = "Yes" Then
            row1(7) = Yes
        Else
            row1(7) = mtcObj.DashInUnusedFiled9
        End If
        '8
        If mtcObj.m2mDiscriptionMatch10 = "No" Then
            row1(8) = No
            count += 1
        ElseIf mtcObj.m2mDiscriptionMatch10 = "Yes" Then
            row1(8) = Yes
        Else
            row1(8) = mtcObj.m2mDiscriptionMatch10
        End If
        '9
        If mtcObj.uomsMatch11 = "No" Then
            row1(9) = No
            count += 1
        ElseIf mtcObj.uomsMatch11 = "Yes" Then
            row1(9) = Yes
        Else
            row1(9) = mtcObj.uomsMatch11
        End If
        '10
        If mtcObj.interface12 = "No" Then
            row1(10) = No
            count += 1
        ElseIf mtcObj.interface12 = "Yes" Then
            row1(10) = Yes
        Else
            row1(10) = mtcObj.interface12
        End If
        '11
        If mtcObj.interpartcopies13 = "No" Then
            row1(11) = No
            count += 1
        ElseIf mtcObj.interpartcopies13 = "Yes" Then
            row1(11) = Yes
        Else
            row1(11) = mtcObj.interpartcopies13
        End If
        '12
        If mtcObj.partcopies14 = "No" Then
            row1(12) = No
            count += 1
        ElseIf mtcObj.partcopies14 = "Yes" Then
            row1(12) = Yes
        Else
            row1(12) = mtcObj.partcopies14
        End If
        '13
        If mtcObj.brokefile15 = "No" Then
            row1(13) = No
            count += 1
        ElseIf mtcObj.brokefile15 = "Yes" Then
            row1(13) = Yes
        Else
            row1(13) = mtcObj.brokefile15
        End If
        '14
        If mtcObj.adjustable16 = "No" Then
            row1(14) = No
            count += 1
        ElseIf mtcObj.adjustable16 = "Yes" Then
            row1(14) = Yes
        Else
            row1(14) = mtcObj.adjustable16
        End If
        '15
        If mtcObj.Matl_spec17 = "No" Then
            row1(15) = No
            count += 1
        ElseIf mtcObj.Matl_spec17 = "Yes" Then
            row1(15) = Yes
        Else
            row1(15) = mtcObj.Matl_spec17
        End If
        '16
        If mtcObj.MaterialUsed18 = "No" Then
            row1(16) = No
            count += 1
        ElseIf mtcObj.MaterialUsed18 = "Yes" Then
            row1(16) = Yes
        Else
            row1(16) = mtcObj.MaterialUsed18
        End If
        '17
        If mtcObj.RemovedUnusedFeatures19 = "No" Then
            row1(17) = No
            count += 1
        ElseIf mtcObj.RemovedUnusedFeatures19 = "Yes" Then
            row1(17) = Yes
        Else
            row1(17) = mtcObj.RemovedUnusedFeatures19
        End If
        '18
        If mtcObj.ASTMminimum20 = "No" Then
            row1(18) = No
            count += 1
        ElseIf mtcObj.ASTMminimum20 = "Yes" Then
            row1(18) = Yes
        Else
            row1(18) = mtcObj.ASTMminimum20
        End If
        '19
        If mtcObj.FlatPattern21 = "No" Then
            row1(19) = No
            count += 1
        ElseIf mtcObj.FlatPattern21 = "Yes" Then
            row1(19) = Yes
        Else
            row1(19) = mtcObj.FlatPattern21
        End If
        '20
        If mtcObj.HoleToolUse22 = "No" Then
            row1(20) = No
            count += 1
        ElseIf mtcObj.HoleToolUse22 = "Yes" Then
            row1(20) = Yes
        Else
            row1(20) = mtcObj.HoleToolUse22
        End If

        'If mtcObj.componentName23 = "No" Then
        '    row1(21) = No
        '    count += 1
        'ElseIf mtcObj.componentName23 = "Yes" Then
        '    row1(21) = Yes
        'Else
        '    row1(21) = mtcObj.componentName23
        'End If


        '21
        If mtcObj.VirtualThread24 = "No" Then
            row1(21) = No
            count += 1
        ElseIf mtcObj.VirtualThread24 = "Yes" Then
            row1(21) = Yes
        Else
            row1(21) = mtcObj.VirtualThread24
        End If
        '22
        If mtcObj.defineFeatures25 = "No" Then
            row1(22) = No
            count += 1
        ElseIf mtcObj.defineFeatures25 = "Yes" Then
            row1(22) = Yes
        Else
            row1(22) = mtcObj.defineFeatures25
        End If
        '23
        If mtcObj.suppressedFeature26 = "No" Then
            row1(23) = No
            count += 1
        ElseIf mtcObj.suppressedFeature26 = "Yes" Then
            row1(23) = Yes
        Else
            row1(23) = mtcObj.suppressedFeature26
        End If
        '24
        If mtcObj.vendorPartNumber27 = "No" Then
            row1(24) = No
            count += 1
        ElseIf mtcObj.vendorPartNumber27 = "Yes" Then
            row1(24) = Yes
        Else
            row1(24) = mtcObj.vendorPartNumber27
        End If
        '25
        If mtcObj.HardwareParts28 = "No" Then
            row1(25) = No
            count += 1
        ElseIf mtcObj.HardwareParts28 = "Yes" Then
            row1(25) = Yes
        Else
            row1(25) = mtcObj.HardwareParts28
        End If
        '26
        If mtcObj.m2mSourceMarked29 = "No" Then
            row1(26) = No
            count += 1
        ElseIf mtcObj.m2mSourceMarked29 = "Yes" Then
            row1(26) = Yes
        Else
            row1(26) = mtcObj.m2mSourceMarked29
        End If
        '27
        If mtcObj.SEstatus30 = "No" Then
            row1(27) = No
            count += 1
        ElseIf mtcObj.SEstatus30 = "Yes" Then
            row1(27) = Yes
        Else
            row1(27) = mtcObj.SEstatus30
        End If
        '28
        row1(28) = count
        '29
        row1(29) = mtcObj.date31
        count = 0
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

    Public Sub KillProcess() Implements IMTC_ReportExcelUtil.KillProcess
        Dim _proceses As Process()
        _proceses = Process.GetProcessesByName("excel")
        For Each proces As Process In _proceses
            proces.Kill()
        Next
    End Sub

End Class
