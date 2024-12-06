Public Class MTR_ReportExcelUtil

    Dim dt As New DataTable()
    Dim Projectname As String = Nothing
    Dim er As String = "No"
    Dim notexist As String = "0"
    Dim No As String = "1"
    Dim Yes As String = "0"
    Public Sub MTR_Report(mtrObj As MTR_Report)
        Try

            'heading
            If Not dt.Columns.Count > 0 Then
                dt.Columns.Add(New DataColumn("Part Name")) '1
                dt.Columns.Add(New DataColumn("Category")) '2
                dt.Columns.Add(New DataColumn("Broken inter-part copies")) '3
                dt.Columns.Add(New DataColumn("Broken part copies")) '4
                dt.Columns.Add(New DataColumn("Dash")) '5
                dt.Columns.Add(New DataColumn("Exist Features")) '6
                dt.Columns.Add(New DataColumn("Mating Parts")) '7
                dt.Columns.Add(New DataColumn("Environment")) '8
                dt.Columns.Add(New DataColumn("Broken Geometry")) '9
                dt.Columns.Add(New DataColumn("UpdateOnFileSave")) '10
                dt.Columns.Add(New DataColumn("ConstrainedFeatures")) '11
                dt.Columns.Add(New DataColumn("UnusedFeatures")) '12
                dt.Columns.Add(New DataColumn("Adjustable")) '13
                dt.Columns.Add(New DataColumn("HardwarePartBox")) '14
                dt.Columns.Add(New DataColumn("Author Name")) '15
                dt.Columns.Add(New DataColumn("Total")) '16
                dt.Columns.Add(New DataColumn("Report Date")) '17
            End If

            Assembly(mtrObj)
            Part(mtrObj)
            Sheetmetal(mtrObj)
            createMTRExcel(mtrObj)
        Catch ex As Exception
            MsgBox(ex)
        End Try

    End Sub
    Public Sub createMTRExcel(mtrObj As MTR_Report)
        Dim strFileName As String = mtrObj.dir + "\MTC_MTR_KPI_Report.xlsx"
        If System.IO.File.Exists(strFileName) Then
            Dim _excel As New Microsoft.Office.Interop.Excel.Application
            Dim wBook As Microsoft.Office.Interop.Excel.Workbook
            Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet
            Dim misValue As Object = System.Reflection.Missing.Value
            wBook = _excel.Workbooks.Open(strFileName)
            'wSheet = wBook.Sheets("MTR Report")
            wSheet = wBook.Sheets.Add()
            wSheet = wBook.ActiveSheet()
            wSheet.Name = "MTR Report"

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
            wSheet.Columns.AutoFit()
            Dim formatRange As Microsoft.Office.Interop.Excel.Range

            formatRange = wSheet.Range("a1", "q1")

            formatRange.EntireRow.Font.Bold = True

            ' formatRange.BorderInside(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Color.Blue)

            formatRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
            formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SandyBrown)

            formatRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous


            formatRange = wSheet.Range("a1", "q1")
            formatRange.RowHeight = 20
            wSheet.Range("A1:Q1").Columns.EntireColumn.AutoFit()
            formatRange.Font.Name = "Arial"
            formatRange.Font.Size = 11
            Dim str2 = wSheet.UsedRange.Rows.Count + 1
            For i = 2 To str2
                formatRange = wSheet.Range("a" & i & ":q" & i & "")
                formatRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                ' formatRange.BorderInside(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Color.Blue)
                formatRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
                formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow)
                wSheet.Columns.AutoFit()

            Next


            '        Dim strFileName As String = mtrObj.dir + "\MTR_Report.xlsx"
            'If System.IO.File.Exists(strFileName) Then
            '    System.IO.File.Delete(strFileName)
            'End If
            If mtrObj.i = mtrObj.files.Count - 1 Then
                _excel.DisplayAlerts = False
                wBook.SaveAs(strFileName)
                wBook.Close()
                _excel.Quit()
            End If
        Else
            Dim _excel As New Microsoft.Office.Interop.Excel.Application
            Dim wBook As Microsoft.Office.Interop.Excel.Workbook
            Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet

            wBook = _excel.Workbooks.Add()

            wSheet = wBook.ActiveSheet()
            wSheet.Name = "MTR Report"

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
            wSheet.Columns.AutoFit()
            Dim formatRange As Microsoft.Office.Interop.Excel.Range

            formatRange = wSheet.Range("a1", "q1")

            formatRange.EntireRow.Font.Bold = True

            ' formatRange.BorderInside(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Color.Blue)

            formatRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
            formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SandyBrown)

            formatRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous


            formatRange = wSheet.Range("a1", "Q1")
            formatRange.RowHeight = 20
            wSheet.Range("A1:Q1").Columns.EntireColumn.AutoFit()
            formatRange.Font.Name = "Arial"
            formatRange.Font.Size = 11
            Dim str2 = wSheet.UsedRange.Rows.Count + 1
            For i = 2 To str2
                formatRange = wSheet.Range("a" & i & ":q" & i & "")
                formatRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                ' formatRange.BorderInside(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Color.Blue)
                formatRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
                formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow)
                wSheet.Columns.AutoFit()

            Next


            '        Dim strFileName As String = mtrObj.dir + "\MTR_Report.xlsx"
            'If System.IO.File.Exists(strFileName) Then
            '    System.IO.File.Delete(strFileName)
            'End If




            If mtrObj.i = mtrObj.files.Count - 1 Then
                _excel.DisplayAlerts = False
                wBook.SaveAs(strFileName)
                wBook.Close()
                _excel.Quit()
            End If
        End If

    End Sub
    Public Sub Assembly(mtrObj As MTR_Report)
        Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
        xlsWB = xlsApp.Workbooks.Open(mtrObj.excelFilePath, Nothing, False)

        Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing



        xlsSheet1 = xlsWB.Sheets("Assembly")
        xlsSheet1.Activate()
        Dim RowCnt = xlsSheet1.UsedRange.Rows.Count

        Dim ColCnt = xlsSheet1.UsedRange.Columns.Count


        xlsCell1 = xlsSheet1.UsedRange()
        For i = 4 To ColCnt
            '1PartName 
            mtrObj.PartName1 = xlsSheet1.Rows(1).Cells(i).value.ToString()

            '2Category
            mtrObj.Category2 = xlsSheet1.Name

            '3 broken interpartcopies/Verify that the inter-part copies are broken when released
            mtrObj.interpartcopies3 = If(xlsSheet1.Rows(3).Cells(i).value = Nothing, er, xlsSheet1.Rows(3).Cells(i).value.ToString())

            '4 broken partcopies/Verify that the part copies are broken when released
            mtrObj.partcopies4 = If(xlsSheet1.Rows(4).Cells(i).value = Nothing, er, xlsSheet1.Rows(4).Cells(i).value.ToString())

            '5 Dash/ Verify the all categories below have been populated with relative Information Or, at a minimum, a dash (Summary, project, custom)
            mtrObj.Dash5 = If(xlsSheet1.Rows(5).Cells(i).value = Nothing, er, xlsSheet1.Rows(5).Cells(i).value.ToString())

            '6 Exist Features/Do assembly features exist within the assembly model? If present, can they be removed?
            mtrObj.ExistFeatures6 = If(xlsSheet1.Rows(6).Cells(i).value = Nothing, er, xlsSheet1.Rows(6).Cells(i).value.ToString())

            '7 mating parts/Verify that mating parts have been checked for interferences
            mtrObj.MatingParts7 = If(xlsSheet1.Rows(7).Cells(i).value = Nothing, er, xlsSheet1.Rows(7).Cells(i).value.ToString())

            '8 environment/Verify that there are no interferences with objects in the environment
            mtrObj.Environment8 = If(xlsSheet1.Rows(8).Cells(i).value = Nothing, er, xlsSheet1.Rows(8).Cells(i).value.ToString())

            '9 broken geometry/Verify that the included geometry is broken when released
            mtrObj.Geometry9 = If(xlsSheet1.Rows(10).Cells(i).value = Nothing, er, xlsSheet1.Rows(10).Cells(i).value.ToString())

            '10 UpdateOnFileSave10/Verify that the “Update on File Save” is UNCHECKED
            mtrObj.UpdateOnFileSave10 = If(xlsSheet1.Rows(9).Cells(i).value = Nothing, er, xlsSheet1.Rows(9).Cells(i).value.ToString())

            '11ConstrainedFeatures/
            mtrObj.ConstrainedFeatures11 = notexist

            '12 UnusedFeatures/
            mtrObj.UnusedFeatures12 = notexist

            '13 Adjustable/
            mtrObj.Adjustable13 = notexist

            '14HardwarePartBox
            mtrObj.HardwarePartBox14 = notexist

            '15Author Name/Who is the Author of the file?
            mtrObj.author15 = If(xlsSheet1.Rows(11).Cells(i).value = Nothing, er, xlsSheet1.Rows(11).Cells(i).value.ToString())
            '------------------------------------------------------------------
            row1Data(mtrObj)

        Next


        xlsWB.Close()
        xlsApp.Quit()
        releaseObject(xlsApp)
        releaseObject(xlsWB)
        releaseObject(xlsSheet1)
        killProcess()



    End Sub

    Public Sub Part(mtrObj As MTR_Report)
        Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
        xlsWB = xlsApp.Workbooks.Open(mtrObj.excelFilePath, Nothing, False)

        Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing



        xlsSheet1 = xlsWB.Sheets("Part")
        xlsSheet1.Activate()
        Dim RowCnt = xlsSheet1.UsedRange.Rows.Count

        Dim ColCnt = xlsSheet1.UsedRange.Columns.Count


        xlsCell1 = xlsSheet1.UsedRange()
        For i = 4 To ColCnt
            '1PartName 
            mtrObj.PartName1 = xlsSheet1.Rows(1).Cells(i).value.ToString()

            '2Category
            mtrObj.Category2 = xlsSheet1.Name

            '3 broken interpartcopies/Verify that the inter-part copies are broken when released
            mtrObj.interpartcopies3 = notexist

            '4 broken partcopies/Verify that the part copies are broken when released
            mtrObj.partcopies4 = notexist

            '5 Dash/ Verify the all categories below have been populated with relative Information Or, at a minimum, a dash (Summary, project, custom)
            mtrObj.Dash5 = If(xlsSheet1.Rows(5).Cells(i).value = Nothing, er, xlsSheet1.Rows(5).Cells(i).value.ToString())

            '6 Exist Features/Do assembly features exist within the assembly model? If present, can they be removed?
            mtrObj.ExistFeatures6 = notexist

            '7 mating parts/Verify that mating parts have been checked for interferences
            mtrObj.MatingParts7 = notexist

            '8 environment/Verify that there are no interferences with objects in the environment
            mtrObj.Environment8 = notexist

            '9 broken geometry/Verify that the included geometry is broken when released
            mtrObj.Geometry9 = notexist

            '10 UpdateOnFileSave10/Verify that the “Update on File Save” is UNCHECKED
            mtrObj.UpdateOnFileSave10 = If(xlsSheet1.Rows(8).Cells(i).value = Nothing, er, xlsSheet1.Rows(8).Cells(i).value.ToString())

            '11 ConstrainedFeatures/Verify that ALL features have been fully constrained
            mtrObj.ConstrainedFeatures11 = If(xlsSheet1.Rows(2).Cells(i).value = Nothing, er, xlsSheet1.Rows(2).Cells(i).value.ToString())

            '12 UnusedFeatures/Verify that ALL suppressed and unused features have been removed
            mtrObj.UnusedFeatures12 = If(xlsSheet1.Rows(3).Cells(i).value = Nothing, er, xlsSheet1.Rows(3).Cells(i).value.ToString())

            '13 Adjustable/Verify that the part model is NOT adjustable
            mtrObj.Adjustable13 = If(xlsSheet1.Rows(4).Cells(i).value = Nothing, er, xlsSheet1.Rows(4).Cells(i).value.ToString())

            '14 HardwarePartBox/If the model is a fastener then the HARDWARE PART box should be checked
            mtrObj.HardwarePartBox14 = If(xlsSheet1.Rows(6).Cells(i).value = Nothing, er, xlsSheet1.Rows(6).Cells(i).value.ToString())

            '15Author Name/Who is the Author of the file?
            mtrObj.author15 = If(xlsSheet1.Rows(9).Cells(i).value = Nothing, er, xlsSheet1.Rows(9).Cells(i).value.ToString())
            '------------------------------------------------------------------
            row1Data(mtrObj)

        Next


        xlsWB.Close()
        xlsApp.Quit()
        releaseObject(xlsApp)
        releaseObject(xlsWB)
        releaseObject(xlsSheet1)
        killProcess()
    End Sub

    Public Sub Sheetmetal(mtrObj As MTR_Report)
        Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
        xlsWB = xlsApp.Workbooks.Open(mtrObj.excelFilePath, Nothing, False)

        Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing



        xlsSheet1 = xlsWB.Sheets("Sheetmetal")
        xlsSheet1.Activate()
        Dim RowCnt = xlsSheet1.UsedRange.Rows.Count

        Dim ColCnt = xlsSheet1.UsedRange.Columns.Count


        xlsCell1 = xlsSheet1.UsedRange()
        For i = 4 To ColCnt
            '1PartName 
            mtrObj.PartName1 = xlsSheet1.Rows(1).Cells(i).value.ToString()

            '2Category
            mtrObj.Category2 = xlsSheet1.Name

            '3 broken interpartcopies/Verify that the inter-part copies are broken when released
            mtrObj.interpartcopies3 = If(xlsSheet1.Rows(5).Cells(i).value = Nothing, er, xlsSheet1.Rows(5).Cells(i).value.ToString())

            '4 broken partcopies/Verify that the part copies are broken when released
            mtrObj.partcopies4 = If(xlsSheet1.Rows(6).Cells(i).value = Nothing, er, xlsSheet1.Rows(6).Cells(i).value.ToString())

            '5 Dash/ Verify the all categories below have been populated with relative Information Or, at a minimum, a dash (Summary, project, custom)
            mtrObj.Dash5 = If(xlsSheet1.Rows(7).Cells(i).value = Nothing, er, xlsSheet1.Rows(7).Cells(i).value.ToString())

            '6 Exist Features/Do assembly features exist within the assembly model? If present, can they be removed?
            mtrObj.ExistFeatures6 = notexist

            '7 mating parts/Verify that mating parts have been checked for interferences
            mtrObj.MatingParts7 = notexist

            '8 environment/Verify that there are no interferences with objects in the environment
            mtrObj.Environment8 = notexist

            '9 broken geometry/Verify that the included geometry is broken when released
            mtrObj.Geometry9 = If(xlsSheet1.Rows(10).Cells(i).value = Nothing, er, xlsSheet1.Rows(10).Cells(i).value.ToString())

            '10 UpdateOnFileSave10/Verify that the “Update on File Save” is UNCHECKED
            mtrObj.UpdateOnFileSave10 = If(xlsSheet1.Rows(9).Cells(i).value = Nothing, er, xlsSheet1.Rows(9).Cells(i).value.ToString())

            '11 ConstrainedFeatures/Verify that ALL features have been fully constrained
            mtrObj.ConstrainedFeatures11 = If(xlsSheet1.Rows(2).Cells(i).value = Nothing, er, xlsSheet1.Rows(2).Cells(i).value.ToString())

            '12 UnusedFeatures/Verify that ALL suppressed and unused features have been removed
            mtrObj.UnusedFeatures12 = If(xlsSheet1.Rows(3).Cells(i).value = Nothing, er, xlsSheet1.Rows(3).Cells(i).value.ToString())

            '13 Adjustable/Verify that the part model is NOT adjustable
            mtrObj.Adjustable13 = If(xlsSheet1.Rows(4).Cells(i).value = Nothing, er, xlsSheet1.Rows(4).Cells(i).value.ToString())

            '14 HardwarePartBox/If the model is a fastener then the HARDWARE PART box should be checked
            mtrObj.HardwarePartBox14 = notexist

            '15Author Name/Who is the Author of the file?
            mtrObj.author15 = If(xlsSheet1.Rows(11).Cells(i).value = Nothing, er, xlsSheet1.Rows(11).Cells(i).value.ToString())
            '------------------------------------------------------------------
            row1Data(mtrObj)

        Next


        xlsWB.Close()
        xlsApp.Quit()
        releaseObject(xlsApp)
        releaseObject(xlsWB)
        releaseObject(xlsSheet1)
        killProcess()
    End Sub

    Public Sub row1Data(mtrObj As MTR_Report)
        Dim row1 As DataRow = dt.NewRow()
        Dim count As Integer = 0
        dt.Rows.Add(row1)
        row1(0) = mtrObj.PartName1
        row1(1) = mtrObj.Category2

        If mtrObj.interpartcopies3 = "No" Then
            row1(2) = No
            count += 1
        ElseIf mtrObj.interpartcopies3 = "Yes" Then
            row1(2) = Yes
        Else
            row1(2) = mtrObj.interpartcopies3
        End If


        If mtrObj.partcopies4 = "No" Then
            row1(3) = No
            count += 1
        ElseIf mtrObj.partcopies4 = "Yes" Then
            row1(3) = Yes
        Else
            row1(3) = mtrObj.partcopies4
        End If


        If mtrObj.Dash5 = "No" Then
            row1(4) = No
            count += 1
        ElseIf mtrObj.Dash5 = "Yes" Then
            row1(4) = Yes
        Else
            row1(4) = mtrObj.Dash5
        End If

        If mtrObj.ExistFeatures6 = "No" Then
            row1(5) = No
            count += 1
        ElseIf mtrObj.ExistFeatures6 = "Yes" Then
            row1(5) = Yes
        Else
            row1(5) = mtrObj.ExistFeatures6
        End If

        If mtrObj.MatingParts7 = "No" Then
            row1(6) = No
            count += 1
        ElseIf mtrObj.MatingParts7 = "Yes" Then
            row1(6) = Yes
        Else
            row1(6) = mtrObj.MatingParts7
        End If


        If mtrObj.Environment8 = "No" Then
            row1(7) = No
            count += 1
        ElseIf mtrObj.Environment8 = "Yes" Then
            row1(7) = Yes
        Else
            row1(7) = mtrObj.Environment8
        End If


        If mtrObj.Geometry9 = "No" Then
            row1(8) = No
            count += 1
        ElseIf mtrObj.Geometry9 = "Yes" Then
            row1(8) = Yes
        Else
            row1(8) = mtrObj.Geometry9
        End If



        If mtrObj.UpdateOnFileSave10 = "No" Then
            row1(9) = No
            count += 1
        ElseIf mtrObj.UpdateOnFileSave10 = "Yes" Then
            row1(9) = Yes
        Else
            row1(9) = mtrObj.UpdateOnFileSave10
        End If


        If mtrObj.ConstrainedFeatures11 = "No" Then
            row1(10) = No
            count += 1
        ElseIf mtrObj.ConstrainedFeatures11 = "Yes" Then
            row1(10) = Yes
        Else

            row1(10) = mtrObj.ConstrainedFeatures11

        End If


        If mtrObj.UnusedFeatures12 = "No" Then
            row1(11) = No
            count += 1
        ElseIf mtrObj.UnusedFeatures12 = "Yes" Then
            row1(11) = Yes
        Else
            row1(11) = mtrObj.UnusedFeatures12
        End If


        If mtrObj.Adjustable13 = "No" Then
            row1(12) = No

        ElseIf mtrObj.Adjustable13 = "Yes" Then
            row1(12) = Yes
        Else
            row1(12) = mtrObj.Adjustable13
        End If


        If mtrObj.HardwarePartBox14 = "No" Or "False" Then
            row1(13) = No
            count += 1
        ElseIf mtrObj.HardwarePartBox14 = "Yes" Then
            row1(13) = Yes
        Else
            row1(13) = mtrObj.HardwarePartBox14
        End If
        row1(14) = mtrObj.author15
        row1(15) = count
        '17 Report Date
        mtrObj.date16 = Date.Now
        row1(16) = mtrObj.date16
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
