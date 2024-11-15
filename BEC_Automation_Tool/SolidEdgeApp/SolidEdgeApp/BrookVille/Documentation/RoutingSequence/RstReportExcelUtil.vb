Imports Microsoft.Office.Interop
Imports System.IO
Public Class RstReportExcelUtil
    Dim dt As New DataTable()
    Dim dt2 As New DataTable()
    Dim ProdTotal As Decimal = Nothing
    Dim MoveTotal As Integer = Nothing
    Dim ProjectName As String = Nothing
    Dim ProjectNumber As String = Nothing
    Public Sub mtcReport(rstObj As RountingSequenceClass)
        Try

            'Column
            dt.Columns.Add(New DataColumn("--->"))
            For i = 10 To 100 Step 10
                dt.Columns.Add(New DataColumn(i))
            Next
            dt.Columns.Add(New DataColumn("Total"))
            GetDgvProcessData(rstObj)
            RS_Sheetmetal(rstObj)

            ' creteNewMTCReport(rstObj)
        Catch ex As Exception
            'MessageBox.Show(ex.Message + ex.StackTrace, "Message")
        End Try
    End Sub
    Public Sub RS_Sheetmetal(rstObj As RountingSequenceClass)
        Try
            Dim a As String
            Dim notexist = "FALSE"
            Dim order As New ArrayList()


            Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
            xlsWB = xlsApp.Workbooks.Open(rstObj.excelFilepath, Nothing, False)

            Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim xlsSheet2 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing



            xlsSheet1 = xlsWB.Sheets(rstObj.CategoryName)
            xlsSheet1.Activate()
            rstObj.CategoryName = xlsSheet1.Name.ToUpper
            Dim RowCnt = xlsSheet1.UsedRange.Rows.Count

            Dim ColCnt = xlsSheet1.UsedRange.Columns.Count


            xlsCell1 = xlsSheet1.UsedRange()

            rstObj.order.Clear()
            rstObj.prodtime.Clear()
            rstObj.movetime.Clear()
            ProdTotal = Nothing
            MoveTotal = Nothing
            Dim count = Nothing

            ''project name
            'rstObj.ProjectName = xlsSheet1.Rows(2).Cells(1).Text
            For i = 2 To RowCnt


                rstObj.order.Clear()
                rstObj.prodtime.Clear()
                rstObj.movetime.Clear()
                ProdTotal = Nothing
                MoveTotal = Nothing
                '1PartName Value(Error)
                rstObj.PartName = xlsSheet1.Rows(i).Cells(1).Text

                'order10/9020
                For j = 2 To ColCnt Step 3
                    a = If(xlsSheet1.Rows(i).Cells(j).Text = Nothing, "", xlsSheet1.Rows(i).Cells(j).Text)
                    If Not a = notexist And Not a = "" Then
                        rstObj.order.Add(a)
                    End If
                Next


                'prodtime heading
                a = "ProdTime"
                rstObj.prodtime.Add(a)

                'prodtime value 1 
                For j = 3 To ColCnt Step 3
                    a = If(xlsSheet1.Rows(i).Cells(j).Text = Nothing, "", xlsSheet1.Rows(i).Cells(j).Text.ToString)
                    If Not a = "" And Not a = notexist Then
                        rstObj.prodtime.Add(a)
                        ProdTotal += a
                    End If
                Next

                'movetime heading
                a = "MoveTime"
                rstObj.movetime.Add(a)
                For j = 4 To ColCnt Step 3
                    a = If(xlsSheet1.Rows(i).Cells(j).Text = Nothing, "", xlsSheet1.Rows(i).Cells(j).Text)
                    If Not a = notexist And Not a = "" Then
                        rstObj.movetime.Add(a)
                        MoveTotal += a
                    End If
                Next

                '----------------------------------------------------------------------------------------------------------
                'If rstObj.CategoryName = "SheetMetal".ToUpper Then
                '    row1Data(rstObj)
                'ElseIf rstObj.CategoryName = "Assembly".ToUpper Then
                '    'data of sheetmetal row1data()
                '    row1Data(rstObj)
                'ElseIf rstObj.CategoryName = "Structure".ToUpper Then
                '    'data of sheetmetal row1data()
                row1Data(rstObj)
                'End If


                '--------------------------------------------------------------------------------------------------------------------------------

            Next
            'project name
            xlsSheet2 = xlsWB.Sheets("Assembly Data")
            xlsSheet1.Activate()
            rstObj.ProjectName = xlsSheet2.Rows(2).Cells(1).text
            ProjectName = xlsSheet2.Rows(2).Cells(5).text
            ProjectNumber = xlsSheet2.Rows(2).Cells(1).text
            rstObj.ProjectName += " ||" + xlsSheet2.Rows(2).Cells(5).text
            rstObj.user = xlsSheet2.Rows(2).Cells(8).text
            xlsWB.Close()
            xlsApp.Quit()
            releaseObject(xlsApp)
            releaseObject(xlsWB)
            releaseObject(xlsSheet1)
            releaseObject(xlsSheet2)
            ' killProcess()
            'Return dt
            rstObj.Maindt = dt
        Catch ex As Exception
            'MessageBox.Show($"While Reading the excel Sheet :[{rstObj.CategoryName}]{vbNewLine} {ex.Message}{vbNewLine}{ex.StackTrace}", "Message")
        End Try
    End Sub

#Region "Get&set for sheetmetal"
    Public Sub Get_dt3data(ByVal rstObj As RountingSequenceClass)
        Try
            Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
            xlsWB = xlsApp.Workbooks.Open(rstObj.excelFilepath, Nothing, False)
            Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing

            xlsSheet1 = xlsWB.Sheets(rstObj.CategoryName.ToLower + " Data")
            xlsSheet1.Activate()
            Dim RowCnt = xlsSheet1.UsedRange.Rows.Count
            Dim ColCnt = xlsSheet1.UsedRange.Columns.Count
            xlsCell1 = xlsSheet1.UsedRange()
            Dim a As String = Nothing
            Dim partdetails As New ArrayList()
            Dim partdetailsValues As New ArrayList()
            partdetails.Clear()
            partdetailsValues.Clear()
            rstObj.dt3.Clear()

            a = rstObj.PartName
            rstObj.dt3.Columns.Clear()
            rstObj.dt3.Columns.Add(New DataColumn(a))
            rstObj.dt3.Columns.Add(New DataColumn("Data"))
            For i = 2 To ColCnt
                partdetails.Add(xlsSheet1.Rows(1).Cells(i).Text.ToString)

            Next

            For i = 2 To RowCnt

                If xlsSheet1.Rows(i).Cells(1).Text = rstObj.PartName Then
                    For j = 2 To ColCnt
                        ' MsgBox(xlsSheet1.Rows(i).Cells(j).Text.ToString)

                        partdetailsValues.Add(xlsSheet1.Rows(i).Cells(j).Text.ToString)
                    Next

                    For j = 0 To partdetails.Count - 1
                        Dim row1 As DataRow = rstObj.dt3.NewRow()
                        rstObj.dt3.Rows.Add(row1)
                        row1(0) = partdetails(j)
                        row1(1) = partdetailsValues(j)
                        If partdetails(j) = "Material Description" Then
                            rstObj.MaterialDescription = partdetailsValues(j)
                        End If
                        If partdetails(j) = "FilePath" Then
                            rstObj.FilePath = partdetailsValues(j)
                        End If
                    Next

                End If
            Next




            xlsWB.Close()
            xlsApp.Quit()
            releaseObject(xlsApp)
            releaseObject(xlsWB)
            releaseObject(xlsSheet1)

            ' killProcess()
        Catch ex As Exception
            'MessageBox.Show($"Error While Reading the Excel Sheet:[{rstObj.CategoryName} Data] {vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}", "Message")
        End Try
    End Sub
    Public Sub Set_dt3data(ByVal rstObj As RountingSequenceClass)
        Try
            Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
            xlsWB = xlsApp.Workbooks.Open(rstObj.excelFilepath, Nothing, False)
            Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing

            xlsSheet1 = xlsWB.Sheets(rstObj.CategoryName + " Data")
            xlsSheet1.Activate()
            Dim RowCnt = xlsSheet1.UsedRange.Rows.Count
            Dim ColCnt = xlsSheet1.UsedRange.Columns.Count
            xlsCell1 = xlsSheet1.UsedRange()


            For i = 2 To RowCnt

                If xlsSheet1.Rows(i).Cells(1).Text = rstObj.PartName Then
                    For j = 0 To rstObj.dt3.Rows.Count - 1
                        'MsgBox(rstObj.dt3.Rows(j)(1))
                        'MsgBox(xlsSheet1.Rows(i).Cells(j + 2).Text)

                        xlsSheet1.Rows(i).Cells(j + 2) = rstObj.dt3.Rows(j)(1)
                    Next

                End If
            Next
            xlsApp.Application.DisplayAlerts = False
            xlsWB.SaveAs(rstObj.excelFilepath)
            xlsWB.Close()
            xlsApp.Quit()
            releaseObject(xlsApp)
            releaseObject(xlsWB)
            releaseObject(xlsSheet1)
            killProcess()

        Catch ex As Exception
            'MessageBox.Show($"(btn Apply Values) {vbNewLine} While Updating the Data into Excel Sheet:[{rstObj.CategoryName} Data] {vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}")
        End Try
    End Sub
    Public Sub Set_Maindt_data(ByVal rstObj As RountingSequenceClass)
        Try
            Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
            xlsWB = xlsApp.Workbooks.Open(rstObj.excelFilepath, Nothing, False)

            Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim xlsSheet2 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing


            xlsSheet1 = xlsWB.Sheets(rstObj.CategoryName)
            xlsSheet1.Activate()
            Dim RowCnt = xlsSheet1.UsedRange.Rows.Count

            Dim ColCnt = xlsSheet1.UsedRange.Columns.Count


            xlsCell1 = xlsSheet1.UsedRange()
            Dim a As String
            Dim notexist = "FALSE"
            ' Dim order As ArrayList = New ArrayList()


            rstObj.order.Clear()
            rstObj.prodtime.Clear()
            rstObj.movetime.Clear()
            ProdTotal = Nothing
            MoveTotal = Nothing
            Dim count = Nothing
            Dim Copycnt = Nothing
            For i = 2 To RowCnt
                If rstObj.PartName IsNot "" And rstObj.PartName = xlsSheet1.Rows(i).Cells(1).Text Then
                    For j = 2 To ColCnt Step 3
                        'Part name
                        a = If(xlsSheet1.Rows(i).Cells(j).Text = Nothing, "", xlsSheet1.Rows(i).Cells(j).Text)
                        If Not a = notexist And Not a = "" And Not rstObj.order.Contains(a) Then
                            rstObj.order.Add(a)
                        ElseIf rstObj.order.Contains(a) Then
                            Copycnt += 1
                        End If
                    Next

                    'a = "ProdTime"
                    'rstObj.prodtime.Add(a)

                    'prodtime value 1 
                    Dim countMatch = 0
                    For j = 3 To ColCnt Step 3
                        countMatch += 1
                        a = If(xlsSheet1.Rows(i).Cells(j).Text = Nothing, "", xlsSheet1.Rows(i).Cells(j).Text.ToString)
                        If Not a = "" And Not a = notexist And Not Copycnt = countMatch Then

                            rstObj.prodtime.Add(a)
                            ProdTotal += a
                        End If
                    Next

                    'a = "MoveTime"
                    'rstObj.movetime.Add(a)
                    For j = 4 To ColCnt Step 3
                        countMatch += 1
                        a = If(xlsSheet1.Rows(i).Cells(j).Text = Nothing, "", xlsSheet1.Rows(i).Cells(j).Text)
                        If Not a = notexist And Not a = "" And Not Copycnt = countMatch Then
                            rstObj.movetime.Add(a)
                            MoveTotal += a
                        End If
                    Next
                    'Dim cnt2 As Integer = 0
                    'For cnt = 0 To rstObj.order.Count - 1
                    '    Dim CopyValue As String = rstObj.order.Item(i).ToString



                    'Next

                    For k = 0 To rstObj.Maindt.Rows.Count - 1
                        'MsgBox(rstObj.Maindt.Rows(k)(0))
                        If (rstObj.Maindt.Rows(k)(0)) = rstObj.PartName Then
                            For m = 1 To rstObj.order.Count
                                rstObj.Maindt.Rows(k)(m) = rstObj.order(m - 1)

                            Next

                            'this for loop for Remove Duplicate or extra data after update value.
                            For nextcount = rstObj.order.Count + 1 To 10
                                'remove Duplicate or extra Order Column
                                rstObj.Maindt.Rows(k)(nextcount) = ""

                                'remove Duplicate or extra ProdTime Column
                                rstObj.Maindt.Rows(k + 1)(nextcount) = ""

                                'remove Duplicate or extra MoveTime Column
                                rstObj.Maindt.Rows(k + 2)(nextcount) = ""


                            Next



                            For l = 1 To rstObj.movetime.Count
                                rstObj.Maindt.Rows(k + 1)(l) = rstObj.prodtime(l - 1)
                                rstObj.Maindt.Rows(k + 2)(l) = rstObj.movetime(l - 1)
                                If l = rstObj.movetime.Count Then
                                    rstObj.Maindt.Rows(k + 1)(11) = ProdTotal
                                    rstObj.Maindt.Rows(k + 2)(11) = MoveTotal
                                End If
                            Next

                        End If


                    Next
                    count = 1
                End If
                If count = 1 Then
                    Exit For
                End If
            Next


            xlsWB.Close()
            xlsApp.Quit()
            releaseObject(xlsApp)
            releaseObject(xlsWB)
            releaseObject(xlsSheet1)
            killProcess()
            'Return dt





        Catch ex As Exception
            'MessageBox.Show($"While Updating the Data into Grid for :{rstObj.CategoryName}{vbNewLine} {ex.Message}{vbNewLine}{ex.StackTrace}")
        End Try
    End Sub
#End Region





#Region "get DgvProcessData"

    Public Sub GetDgvProcessData(rstObj As RountingSequenceClass)
        Try
            Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
            xlsWB = xlsApp.Workbooks.Open(rstObj.excelFilepath, Nothing, False)
            Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing
            If rstObj.CategoryName = "SheetMetal" Then
                rstObj.CategoryName = "Sheemetal"
            End If
            xlsSheet1 = xlsWB.Sheets(rstObj.CategoryName + " Rules")
            xlsSheet1.Activate()
            Dim RowCnt = xlsSheet1.UsedRange.Rows.Count
            Dim ColCnt = xlsSheet1.UsedRange.Columns.Count
            xlsCell1 = xlsSheet1.UsedRange()
            ' MsgBox(xlsSheet1.Name)

            'MsgBox(ColCnt)
            'MsgBox(xlsSheet1.Rows(1).Cells(2).Text.ToString)



            'heading of dtProcess
            For i = 2 To ColCnt
                If i <= 5 Then
                    rstObj.dtProcess.Columns.Add(New DataColumn((xlsSheet1.Rows(1).Cells(i).Text.ToString)))
                    If i = 5 Then
                        Exit For
                    End If
                End If

            Next


            Dim ProcessNames As New ArrayList
            ProcessNames.Clear()
            Dim WC As New ArrayList
            WC.Clear()
            Dim ProdTime As New ArrayList
            ProdTime.Clear()
            Dim MoveTime As New ArrayList
            MoveTime.Clear()

            For i = 2 To RowCnt
                If xlsSheet1.Rows(i).Cells(2).Text.ToString = "" Then
                    Exit For
                End If
                For j = 2 To 5 'ColCnt
                    If Not ProcessNames.Contains((xlsSheet1.Rows(i).Cells(j).Text.ToString)) Then

                        If j = 2 Then
                            ProcessNames.Add((xlsSheet1.Rows(i).Cells(j).Text.ToString))
                        End If
                        If j = 3 Then
                            WC.Add((xlsSheet1.Rows(i).Cells(j).Text.ToString))
                        End If

                        If j = 4 Then
                            MoveTime.Add((xlsSheet1.Rows(i).Cells(j).Text.ToString))
                        End If
                        If j = 5 Then
                            ProdTime.Add((xlsSheet1.Rows(i).Cells(j).Text.ToString)) 'ProdTime.Add("0") 
                        End If

                    End If
                Next
            Next

            'rows of dtProcess



            For j = 0 To ProcessNames.Count - 1
                Dim row1 As DataRow = rstObj.dtProcess.NewRow()
                rstObj.dtProcess.Rows.Add(row1)
                row1(0) = ProcessNames(j)
                row1(1) = WC(j)
                row1(2) = MoveTime(j)
                row1(3) = ProdTime(j)
            Next

            If rstObj.CategoryName = "Sheemetal" Then
                rstObj.CategoryName = "SheetMetal"
            End If
            xlsWB.Close()
            xlsApp.Quit()
            releaseObject(xlsApp)
            releaseObject(xlsWB)
            releaseObject(xlsSheet1)
            ' killProcess()
        Catch ex As Exception
            'MessageBox.Show($"Error While Reading {rstObj.CategoryName} Rules Sheet :{vbNewLine} {ex.Message}{vbNewLine}{ ex.StackTrace}", "Message")
        End Try
    End Sub
#End Region




    'copy of approvesequence
    Public Sub creteNewMTCReport(ByVal rstObj As RountingSequenceClass)


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

        For Each dc In rstObj.Maindt.Columns
            colIndex += 1
            _excel.Cells(1, colIndex) = dc.ColumnName
        Next
        For Each dr In rstObj.Maindt.Rows
            rowIndex += 1
            colIndex = 0
            For Each dc In rstObj.Maindt.Columns
                colIndex += 1
                _excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
            Next
        Next

        Dim strFileName As String = rstObj.dir + "\MTC_RST_Report.xlsx"
        If System.IO.File.Exists(strFileName) Then
            System.IO.File.Delete(strFileName)
        End If

        wBook.SaveAs(strFileName)

        wBook.Close()
        _excel.Quit()
    End Sub
    Public Sub ApproveSequence(rstObj As RountingSequenceClass)
        Try
            Dim _excel As New Microsoft.Office.Interop.Excel.Application
            Dim wBook As Microsoft.Office.Interop.Excel.Workbook
            Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet
            Dim dc As System.Data.DataColumn
            Dim dr As System.Data.DataRow
            Dim colIndex As Integer = 0
            Dim rowIndex As Integer = 0
            Dim colIndex1 As Integer = 0
            Dim rowIndex1 As Integer = 0
            Dim formatRange As Microsoft.Office.Interop.Excel.Range

            Dim existSheet As String = Nothing
            Dim strFileName As String = rstObj.dir + "\MTC_RST_Report.xlsx"

            'existing excel 
            'if sheet is exist it will delete otherwise create a report in new sheet
            If System.IO.File.Exists(strFileName) Then
                wBook = _excel.Workbooks.Open(strFileName)

                Dim strSheetName As New List(Of String)
                For Each wSheet In _excel.Sheets
                    strSheetName.Add(wSheet.Name)
                Next
                wBook.Close()
                _excel.Quit()



                For Each wSheetName In strSheetName
                    If wSheetName = "MTC " + rstObj.CategoryName + " Report" Then
                        existSheet = "MTC " + rstObj.CategoryName + " Report"
                        'System.IO.File.Delete(strFileName)
                    End If
                Next



                If existSheet = "MTC " + rstObj.CategoryName + " Report" Then
                    System.IO.File.Delete(strFileName)
                End If

                '
                wBook = _excel.Workbooks.Open(strFileName)
                wBook.Sheets.Add()
                wSheet = wBook.ActiveSheet()
                wSheet.Name = "MTC " + rstObj.CategoryName + " Report"


                _excel.Cells(1, 1) = Date.Now
                _excel.Cells(2, 1) = "Project Name"
                _excel.Cells(2, 2) = ProjectName.ToUpper
                _excel.Cells(3, 1) = "Project#"
                _excel.Cells(3, 2) = ProjectNumber.ToUpper
                _excel.Cells(4, 1) = "Approved By"
                _excel.Cells(4, 2) = rstObj.user.ToUpper
                _excel.Cells(5, 1) = "Category"
                _excel.Cells(5, 2) = rstObj.CategoryName.ToUpper

                rowIndex = 6

                For Each dc In rstObj.Maindt.Columns
                    colIndex += 1
                    _excel.Cells(rowIndex, colIndex) = dc.ColumnName
                Next
                wSheet.Range("A1:L6").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SlateGray)
                wSheet.Range("A1:L6").Style.Font.Color = Color.White
                formatRange = wSheet.Range("A1:L6")
                formatRange.ColumnWidth = 20
                ' formatRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                formatRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
                For Each dr In rstObj.Maindt.Rows
                    rowIndex += 1
                    colIndex = 0

                    For Each dc In rstObj.Maindt.Columns
                        colIndex += 1
                        _excel.Cells(rowIndex, colIndex) = dr(dc.ColumnName)
                    Next

                Next
                formatRange = wSheet.Range("A7:L" + Convert.ToString(rowIndex))
                formatRange.ColumnWidth = 20
                formatRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                formatRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
                rowIndex1 = 7
                For i = rowIndex1 To rowIndex
                    Dim rng = "A" + Convert.ToString(i) + ":L" + Convert.ToString(i)
                    wSheet.Range(rng).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue)
                    wSheet.Range(rng).Style.Font.Color = Color.Black
                    i += 2
                Next
                rowIndex1 = 8
                For i = rowIndex1 To rowIndex
                    Dim rng = "A" + Convert.ToString(i) + ":L" + Convert.ToString(i)
                    wSheet.Range(rng).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.AliceBlue)
                    wSheet.Range(rng).Style.Font.Color = Color.Black
                    i += 2
                Next
                rowIndex1 = 9
                For i = rowIndex1 To rowIndex
                    Dim rng = "A" + Convert.ToString(i) + ":L" + Convert.ToString(i)
                    wSheet.Range(rng).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.AliceBlue)
                    wSheet.Range(rng).Style.Font.Color = Color.Black
                    i += 2
                Next

                'For i = 1 To 3
                '    Dim rng = "A" + Convert.ToString(rowIndex) + ":L" + Convert.ToString(rowIndex)
                '    If i = 1 Then
                '        wSheet.Range(rng).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.AliceBlue)
                '    Else
                '        wSheet.Range(rng).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue)
                '    End If

                'Next
                ' Dim rng = "A7:L" + Convert.ToString(rowIndex)
                ' wSheet.Range(rng).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.AliceBlue)
                'wSheet.Range(rng).Style.Font.Color = Color.Black
                wSheet.Range("A1:L" + Convert.ToString(rowIndex1)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft


            Else


                'new excel report
                wBook = _excel.Workbooks.Add()
                wSheet = wBook.ActiveSheet()
                wSheet.Name = "MTC " + rstObj.CategoryName + " Report"





                colIndex = 0
                rowIndex = 0
                colIndex1 = 0
                rowIndex1 = 0


                _excel.Cells(1, 1) = Date.Now
                _excel.Cells(2, 1) = "Project Name"
                _excel.Cells(2, 2) = ProjectName.ToUpper
                _excel.Cells(3, 1) = "Project#"
                _excel.Cells(3, 2) = ProjectNumber.ToUpper
                _excel.Cells(4, 1) = "Approved By"
                _excel.Cells(4, 2) = rstObj.user.ToUpper
                _excel.Cells(5, 1) = "Category"
                _excel.Cells(5, 2) = rstObj.CategoryName.ToUpper

                rowIndex = 6

                For Each dc In rstObj.Maindt.Columns
                    colIndex += 1
                    _excel.Cells(rowIndex, colIndex) = dc.ColumnName
                Next
                wSheet.Range("A1:L6").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SlateGray)
                wSheet.Range("A1:L6").Style.Font.Color = Color.White
                formatRange = wSheet.Range("A1:L6")
                formatRange.ColumnWidth = 20
                ' formatRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                formatRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
                For Each dr In rstObj.Maindt.Rows
                    rowIndex += 1
                    colIndex = 0

                    For Each dc In rstObj.Maindt.Columns
                        colIndex += 1
                        _excel.Cells(rowIndex, colIndex) = dr(dc.ColumnName)
                    Next

                Next
                formatRange = wSheet.Range("A7:L" + Convert.ToString(rowIndex))
                formatRange.ColumnWidth = 20
                formatRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                formatRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
                rowIndex1 = 7
                For i = rowIndex1 To rowIndex
                    Dim rng = "A" + Convert.ToString(i) + ":L" + Convert.ToString(i)
                    wSheet.Range(rng).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue)
                    wSheet.Range(rng).Style.Font.Color = Color.Black
                    i += 2
                Next
                rowIndex1 = 8
                For i = rowIndex1 To rowIndex
                    Dim rng = "A" + Convert.ToString(i) + ":L" + Convert.ToString(i)
                    wSheet.Range(rng).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.AliceBlue)
                    wSheet.Range(rng).Style.Font.Color = Color.Black
                    i += 2
                Next
                rowIndex1 = 9
                For i = rowIndex1 To rowIndex
                    Dim rng = "A" + Convert.ToString(i) + ":L" + Convert.ToString(i)
                    wSheet.Range(rng).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.AliceBlue)
                    wSheet.Range(rng).Style.Font.Color = Color.Black
                    i += 2
                Next

                'For i = 1 To 3
                '    Dim rng = "A" + Convert.ToString(rowIndex) + ":L" + Convert.ToString(rowIndex)
                '    If i = 1 Then
                '        wSheet.Range(rng).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.AliceBlue)
                '    Else
                '        wSheet.Range(rng).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSteelBlue)
                '    End If

                'Next
                ' Dim rng = "A7:L" + Convert.ToString(rowIndex)
                ' wSheet.Range(rng).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.AliceBlue)
                'wSheet.Range(rng).Style.Font.Color = Color.Black
                wSheet.Range("A1:L" + Convert.ToString(rowIndex1)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
            End If
#Region "Delete Empty column"
            Dim MyRange = wSheet.UsedRange 'Should be on GR4100
            For iCounter = MyRange.Columns.Count To 1 Step -1
                If _excel.WorksheetFunction.CountA(wSheet.Columns(iCounter).EntireColumn) = 1 Then
                    wSheet.Columns(iCounter).Delete
                End If
            Next iCounter
#End Region
            _excel.DisplayAlerts = False
            wBook.SaveAs(strFileName)
            wBook.Close()
            _excel.Quit()
        Catch ex As Exception
            'MessageBox.Show($"Error While Creating RoutingSequence Report{vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}")
        End Try
    End Sub

    Public Sub row1Data(ByVal rstObj As RountingSequenceClass)
        Try
            Dim row1 As DataRow = dt.NewRow()
            Dim row2 As DataRow = dt.NewRow()
            Dim row3 As DataRow = dt.NewRow()

            For i = 1 To 3
                If i = 1 Then

                    dt.Rows.Add(row1)
                    row1(0) = rstObj.PartName
                    For j = 0 To rstObj.order.Count - 1
                        row1(j + 1) = rstObj.order(j)
                    Next

                End If
                If i = 2 Then
                    dt.Rows.Add(row2)
                    For j = 0 To rstObj.prodtime.Count - 1
                        row2(j) = rstObj.prodtime(j)
                        If j = rstObj.prodtime.Count - 1 Then
                            row2(11) = ProdTotal
                        End If
                    Next

                End If
                If i = 3 Then
                    dt.Rows.Add(row3)
                    For j = 0 To rstObj.movetime.Count - 1
                        row3(j) = rstObj.movetime(j)
                        If j = rstObj.movetime.Count - 1 Then
                            row3(11) = MoveTotal
                        End If
                    Next

                End If
            Next

        Catch ex As Exception
            'MessageBox.Show($"While creating Row Data For Category: {rstObj.CategoryName} {vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}", "Message")
        End Try

    End Sub
#Region "KillProcess"
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

#End Region

#Region "Extras"
    Public Sub ShowExistingMtcReport(ByVal rstObj As RountingSequenceClass)

        Try

            Dim a As String
            Dim xlApp As Excel.Application
            Dim xlWorkBook As Excel.Workbook
            Dim xlWorkSheet As Excel.Worksheet
            xlApp = New Excel.Application
            Dim i As Integer
            Dim strFileName As String = rstObj.dir + "\MTC_RST_Report.xlsx"
            xlWorkBook = xlApp.Workbooks.Open(strFileName)
            xlWorkSheet = xlWorkBook.Sheets("MTC Report")
            xlWorkSheet.Activate()
            Dim rowCnt = xlWorkSheet.UsedRange.Rows.Count
            Dim colCnt = xlWorkSheet.UsedRange.Columns.Count
            ' Dim rowCnt = 13
            'Dim colCnt = 9
            dt.Dispose()
            'MsgBox(xlWorkSheet.Cells(1, 1).value.ToString)

            For j = 1 To colCnt


                dt.Columns.Add(New DataColumn(xlWorkSheet.Cells(1, j).value.ToString))
            Next
            For i = 2 To rowCnt
                a = xlWorkSheet.Rows(i).Cells(1).value.ToString()
                rstObj.prodtime.Add(a)


                'order10
                a = If(xlWorkSheet.Rows(i).Cells(2).value = Nothing, "", xlWorkSheet.Rows(i).Cells(2).value.ToString())
                rstObj.prodtime.Add(a)


                'order20
                a = If(xlWorkSheet.Rows(i).Cells(3).value = Nothing, "", xlWorkSheet.Rows(i).Cells(3).value.ToString())
                rstObj.prodtime.Add(a)

                'order30

                'order40
                a = If(xlWorkSheet.Rows(i).Cells(4).value = Nothing, "", xlWorkSheet.Rows(i).Cells(4).value.ToString())
                rstObj.prodtime.Add(a)

                'order50
                a = If(xlWorkSheet.Rows(i).Cells(5).value = Nothing, "", xlWorkSheet.Rows(i).Cells(5).value.ToString())
                rstObj.prodtime.Add(a)

                'order60
                a = If(xlWorkSheet.Rows(i).Cells(6).value = Nothing, "", xlWorkSheet.Rows(i).Cells(6).value.ToString())
                rstObj.prodtime.Add(a)

                'order70
                a = If(xlWorkSheet.Rows(i).Cells(7).value = Nothing, "", xlWorkSheet.Rows(i).Cells(7).value.ToString())
                rstObj.prodtime.Add(a)

                'order80
                a = If(xlWorkSheet.Rows(i).Cells(8).value = Nothing, "", xlWorkSheet.Rows(i).Cells(8).value.ToString())
                rstObj.prodtime.Add(a)

                a = If(xlWorkSheet.Rows(i).Cells(9).value = Nothing, "", xlWorkSheet.Rows(i).Cells(9).value.ToString())
                rstObj.prodtime.Add(a)


                row2Data(rstObj)
            Next
            xlApp.Close()
            xlApp.Quit()
            releaseObject(xlApp)
            releaseObject(xlApp)
            releaseObject(xlWorkSheet)
            killProcess()
            'Return dt
            rstObj.Maindt = dt
        Catch ex As Exception
            'MessageBox.Show($"Error While Showing Existing Report{vbNewLine}{ex.Message }{vbNewLine}{ex.StackTrace}", "Message")
        End Try
    End Sub
    Public Sub row2Data(ByVal rstObj As RountingSequenceClass)
        Dim row2 As DataRow = dt.NewRow()
        dt.Rows.Add(row2)
        row2(0) = rstObj.prodtime(0)
        row2(1) = rstObj.prodtime(1)
        row2(2) = rstObj.prodtime(2)
        row2(3) = rstObj.prodtime(3)
        row2(4) = rstObj.prodtime(4)
        row2(5) = rstObj.prodtime(5)
        row2(6) = rstObj.prodtime(6)
        row2(7) = rstObj.prodtime(7)
        row2(8) = rstObj.prodtime(8)
    End Sub

#End Region

End Class
