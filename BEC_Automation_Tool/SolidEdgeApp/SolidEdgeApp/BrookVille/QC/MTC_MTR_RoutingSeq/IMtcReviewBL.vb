Public Interface IMtcReviewBL

    Function GetAssemblyPartList() As DataTable

    Function ReadMTCExcel(ByVal mtcExcelPath As String) As DataSet

    Function ReadPropSeedFile(ByVal propseedFilePath As String) As PropSeedFile

End Interface
