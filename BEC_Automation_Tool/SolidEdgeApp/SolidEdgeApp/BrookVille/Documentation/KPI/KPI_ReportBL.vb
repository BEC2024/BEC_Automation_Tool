Public Class KPI_ReportBL
    Dim ResSave As New MTC_ReportExcelUtil()
    Dim ResSave1 As New MTR_ReportExcelUtil()
    Dim ResSave2 As New CommonReportExcelUtil()
    Public Sub MTCReport(ByVal mtcObj As MTC_Report)
        ResSave.MTCReport(mtcObj)
    End Sub

    Public Sub MTR_Report(ByVal mtrObj As MTR_Report)
        ResSave1.MTR_Report(mtrObj)
    End Sub

    Public Sub CommonReport(ByVal comObj As CommonReport)
        ResSave2.CommonReport(comObj)
    End Sub

End Class
