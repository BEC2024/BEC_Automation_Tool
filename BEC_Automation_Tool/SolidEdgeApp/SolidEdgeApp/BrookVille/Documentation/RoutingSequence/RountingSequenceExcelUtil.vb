Imports Microsoft.Office.Interop
Imports System.IO
Public Class RountingSequenceExcelUtil
    Dim rstObj As New RountingSequenceClass
    Dim dt As New DataTable()
    Public Function ViewData(rstObj As RountingSequenceClass) As DataTable
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        xlApp = New Excel.Application
        Dim i As Integer

        xlWorkBook = xlApp.Workbooks.Open("C:\Users\pratikg\source\repos\RoutingSequenceTool\RoutingSequenceTool\resource\Excel\Manufacturing Sequence Task.xlsx")
        xlWorkSheet = xlWorkBook.Sheets("Output ")
        xlWorkSheet.Activate()
        'Dim rowCnt = xlWorkSheet.UsedRange.Rows.Count
        'Dim colCnt = xlWorkSheet.UsedRange.Columns.Count
        Dim rowCnt = 15
        Dim colCnt = 8
        dt.Dispose()

        dt.Columns.Add(New DataColumn("PartName"))
        For j = 2 To colCnt
            dt.Columns.Add(New DataColumn(xlWorkSheet.Cells(1, j).value.ToString))
        Next
        'MsgBox((xlWorkSheet.Cells(2, 2).value.ToString))

        'rstObj.partname = If((xlWorkSheet.Cells(2, 1).value) = Nothing, " ", (xlWorkSheet.Cells(2, 1).value.ToString))
        'rstObj.ten = If((xlWorkSheet.Cells(2, 2).value) = Nothing, "", (xlWorkSheet.Cells(2, 2).value.ToString))
        'rstObj.twenty = If((xlWorkSheet.Cells(2, 3).value) = Nothing, "", (xlWorkSheet.Cells(2, 3).value.ToString))
        'rstObj.thirty = If((xlWorkSheet.Cells(2, 4).value) = Nothing, "", (xlWorkSheet.Cells(2, 4).value.ToString))
        'rstObj.fourty = If((xlWorkSheet.Cells(2, 5).value) = Nothing, "", (xlWorkSheet.Cells(2, 5).value.ToString))
        'rstObj.fifty = If((xlWorkSheet.Cells(2, 6).value) = Nothing, "", (xlWorkSheet.Cells(2, 6).value.ToString))
        'rstObj.sixty = If((xlWorkSheet.Cells(2, 7).value) = Nothing, "", (xlWorkSheet.Cells(2, 7).value.ToString))
        'rstObj.eighty = If((xlWorkSheet.Cells(2, 8).value) = Nothing, "", (xlWorkSheet.Cells(2, 8).value.ToString))
        'MsgBox((xlWorkSheet.Cells(2, 2).value.ToString))
        For i = 2 To rowCnt
            rstObj.partname = If((xlWorkSheet.Cells(i, 1).value) = Nothing, "", (xlWorkSheet.Cells(i, 1).value.ToString))
            rstObj.ten = If((xlWorkSheet.Cells(i, 2).value) = Nothing, "", (xlWorkSheet.Cells(i, 2).value.ToString))
            rstObj.twenty = If((xlWorkSheet.Cells(i, 3).value) = Nothing, "", (xlWorkSheet.Cells(i, 3).value.ToString))
            rstObj.thirty = If((xlWorkSheet.Cells(i, 4).value) = Nothing, "", (xlWorkSheet.Cells(i, 4).value.ToString))
            rstObj.fourty = If((xlWorkSheet.Cells(i, 5).value) = Nothing, "", (xlWorkSheet.Cells(i, 5).value.ToString))
            rstObj.fifty = If((xlWorkSheet.Cells(i, 6).value) = Nothing, "", (xlWorkSheet.Cells(i, 6).value.ToString))
            rstObj.sixty = If((xlWorkSheet.Cells(i, 7).value) = Nothing, "", (xlWorkSheet.Cells(i, 7).value.ToString))
            rstObj.eighty = If((xlWorkSheet.Cells(i, 8).value) = Nothing, "", (xlWorkSheet.Cells(i, 8).value.ToString))

            Row1Data(rstObj)



        Next



        xlWorkBook.Close()
        xlApp.Quit()

        ReleaseObject(xlApp)
        ReleaseObject(xlWorkBook)
        ReleaseObject(xlWorkSheet)
        Killprocess()
        Return dt
    End Function
    Public Sub Row1Data(rstObj As RountingSequenceClass)
        Dim dtime As DateTime
        Dim row1 As DataRow = dt.NewRow()
        dt.Rows.Add(row1)
        row1(0) = rstObj.PartName


        If (rstObj.ten.Contains(".") Or rstObj.ten.Length = "1") Then
            dtime = (New DateTime()).AddDays(rstObj.ten)
            row1(1) = dtime
        Else
            row1(1) = rstObj.ten
        End If

        If (rstObj.twenty.Contains(".") Or rstObj.twenty.Length = "1") Then
            dtime = (New DateTime()).AddDays(rstObj.twenty)
            row1(2) = dtime
        Else
            row1(2) = rstObj.twenty
        End If

        If (rstObj.thirty.Contains(".") Or rstObj.thirty.Length = "1") Then
            dtime = (New DateTime()).AddDays(rstObj.thirty)
            row1(3) = dtime
        Else
            row1(3) = rstObj.thirty
        End If

        If (rstObj.fourty.Contains(".") Or rstObj.fourty.Length = "1") Then
            dtime = (New DateTime()).AddDays(rstObj.fourty)
            row1(4) = dtime
        Else
            row1(4) = rstObj.fourty
        End If

        If (rstObj.fifty.Contains(".") Or rstObj.fifty.Length = "1") Then
            dtime = (New DateTime()).AddDays(rstObj.fifty)
            row1(5) = dtime
        Else
            row1(5) = rstObj.fifty
        End If

        If (rstObj.sixty.Contains(".") Or rstObj.sixty.Length = "1") Then
            dtime = (New DateTime()).AddDays(rstObj.sixty)
            row1(6) = dtime
        Else
            row1(6) = rstObj.sixty
        End If

        If (rstObj.eighty.Contains(".") Or rstObj.eighty.Length = "1") Then
            dtime = (New DateTime()).AddDays(rstObj.eighty)
            row1(7) = dtime
        Else
            row1(7) = rstObj.eighty
        End If



        'row1(1) = If(rstObj.ten.Contains("."), dtime = (New DateTime()).AddDays(rstObj.ten), rstObj.ten)
        'row1(2) = If(rstObj.twenty.Contains("."), dtime = (New DateTime()).AddDays(rstObj.twenty), rstObj.twenty)
        'row1(3) = If(rstObj.thirty.Contains("."), dtime = (New DateTime()).AddDays(rstObj.thirty), rstObj.thirty)
        'row1(4) = If(rstObj.fourty.Contains("."), dtime = (New DateTime()).AddDays(rstObj.fourty), rstObj.fourty)
        'row1(5) = If(rstObj.fifty.Contains("."), dtime = (New DateTime()).AddDays(rstObj.fifty), rstObj.fifty)
        'row1(6) = If(rstObj.sixty.Contains("."), dtime = (New DateTime()).AddDays(rstObj.sixty), rstObj.sixty)
        'row1(7) = If(rstObj.eighty.Contains("."), dtime = (New DateTime()).AddDays(rstObj.eighty), rstObj.eighty)

    End Sub
    Public Sub EditData2(rstObj As RountingSequenceClass)
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        xlApp = New Excel.Application
        Dim i As Integer

        xlWorkBook = xlApp.Workbooks.Open("C:\Users\pratikg\source\repos\RoutingSequenceTool\RoutingSequenceTool\resource\Excel\Manufacturing Sequence Task.xlsx")
        xlWorkSheet = xlWorkBook.Sheets("Output ")
        xlWorkSheet.Activate()
        Dim count = xlWorkSheet.UsedRange.Rows.Count
        ' MsgBox(xlWorkSheet.Cells(2, 0).value.ToString)
        MsgBox(rstObj.dt2.Rows(0)(2))
        For i = 1 To count
            Dim partname = If((xlWorkSheet.Cells(i, 1).value) = Nothing, "", (xlWorkSheet.Cells(i, 1).value.ToString))
            If rstObj.PartName = partname Then
                xlWorkSheet.Cells(i, 1) = partname
                xlWorkSheet.Cells(i, 2) = rstObj.dt2.Rows(0)(2)
                xlWorkSheet.Cells(i, 3) = rstObj.dt2.Rows(1)(2)
                xlWorkSheet.Cells(i, 4) = rstObj.dt2.Rows(2)(2)
                xlWorkSheet.Cells(i, 5) = rstObj.dt2.Rows(3)(2)
                xlWorkSheet.Cells(i, 6) = rstObj.dt2.Rows(4)(2)
                xlWorkSheet.Cells(i, 7) = rstObj.dt2.Rows(5)(2)
                xlWorkSheet.Cells(i, 8) = rstObj.dt2.Rows(6)(2)

            End If
        Next

        xlWorkSheet.SaveAs("C:\Users\pratikg\source\repos\RoutingSequenceTool\RoutingSequenceTool\resource\Excel\Manufacturing Sequence Task.xlsx")

        xlWorkBook.Close()
        xlApp.Quit()

        ReleaseObject(xlApp)
        ReleaseObject(xlWorkBook)
        ReleaseObject(xlWorkSheet)
        Killprocess()
        'If rstObj.dt2.Rows.Count > 0 Then
        '    For j As Integer = 0 To rstObj.dt2.Rows.Count - 1

        '        MsgBox(rstObj.dt2.Rows(1)(0))
        '    Next
        'End If
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
    Private Sub Killprocess()
        Dim _proceses As Process()
        _proceses = Process.GetProcessesByName("excel")
        For Each proces As Process In _proceses
            proces.Kill()
        Next
    End Sub
End Class
