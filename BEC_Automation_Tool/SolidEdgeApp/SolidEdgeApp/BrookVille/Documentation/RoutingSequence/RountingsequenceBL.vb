Public Class RountingsequenceBL
    Dim a As New RountingSequenceExcelUtil()
    Dim rstObj As New RountingSequenceClass
    Public Function ViewData() As DataTable
        Dim dt As DataTable
        dt = a.ViewData(rstObj)
        Return dt
    End Function
    Public Sub EditData2(rstObj As RountingSequenceClass)
        a.EditData2(rstObj)
    End Sub
End Class
