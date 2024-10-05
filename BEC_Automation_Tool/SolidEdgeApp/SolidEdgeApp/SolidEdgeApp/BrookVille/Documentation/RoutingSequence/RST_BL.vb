Public Class RST_BL
    Dim a As New RstReportExcelUtil
    ' Dim SolidObj As SolidEdgeUtil = New SolidEdgeUtil
    Dim SolidObj As New SEUtill
    Dim rstObj As New RountingSequenceClass

    Public Sub MtcReport(rstObj As RountingSequenceClass)
        a.mtcReport(rstObj)
    End Sub

    Public Sub ShowExistingMtcReport(rstObj As RountingSequenceClass)
        a.ShowExistingMtcReport(rstObj)
    End Sub
    Public Sub ApproveSequence(rstObj As RountingSequenceClass)
        a.ApproveSequence(rstObj)
    End Sub

    Public Sub Get_dt3data(rstObj As RountingSequenceClass)
        a.Get_dt3data(rstObj)

    End Sub
    Public Sub Set_dt3data(rstObj As RountingSequenceClass)
        a.Set_dt3data(rstObj)
    End Sub

    Public Sub Get_Maindt_data(rstObj As RountingSequenceClass)
        a.Set_Maindt_data(rstObj)
    End Sub
    Public Sub OpenSEPart(rstObj As RountingSequenceClass)
        SolidObj.OpenSEPart(rstObj)
    End Sub
    Public Sub OpenSEDocument(rstObj As RountingSequenceClass)
        SolidObj.OpenDocument(rstObj)
    End Sub
End Class
