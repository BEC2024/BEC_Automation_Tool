Public Class MTC_MTR_Model

    Public dtM2M As System.Data.DataTable = New DataTable("M2MData")

    Public dtAuthorData As System.Data.DataTable = New DataTable("AuthorData")

    'Public M2MprojectNameList As List(Of String) = New List(Of String)()

    Public assemblyPath As String = String.Empty

    Public projectNameList As List(Of String) = New List(Of String)()

    Public authorList As List(Of String) = New List(Of String)()

    Public dtCurrentAssemblyData As System.Data.DataTable = New DataTable("Assembly Data")

    Public dtFilteredAssemblyData As System.Data.DataTable = New DataTable("Filtered Assembly Data")

    Public dsMTCReport As System.Data.DataSet = New DataSet("MTC Report")

    Public dtBECAuthorAssemblyData As System.Data.DataTable = New DataTable("BEC Authors Data")

    Public dtNonBECAuthorAssemblyData As System.Data.DataTable = New DataTable("NON BEC Authors Data")

    Public exportDirectoryLocation As String = String.Empty

    '2nd Sep 2024
    Public export_MTR_Report_DirectoryLocation As String = String.Empty

    Public export_Routing_Report_DirectoryLocation As String = String.Empty

    Public baseLineDirectoryLocation As String = String.Empty

    Public BOMCount As Integer = 0

    '===================

    Public mtcAssemblyList_BEC As List(Of MTC_Assembly) = New List(Of MTC_Assembly)()

    Public mtcBaseLineList_BEC As List(Of MTC_BaseLine) = New List(Of MTC_BaseLine)()

    Public mtcPartList_BEC As List(Of MTC_Part) = New List(Of MTC_Part)()

    Public mtcElectricalPartList_BEC As List(Of MTC_Electrical) = New List(Of MTC_Electrical)()
    '

    Public mtcSheetMetalList_BEC As List(Of MTC_SheetMetal) = New List(Of MTC_SheetMetal)()


    Public mtrAssemblyList_BEC As List(Of MTR_Assembly) = New List(Of MTR_Assembly)()

    Public mtrPartList_BEC As List(Of MTR_Part) = New List(Of MTR_Part)()

    Public mtrSheetMetalList_BEC As List(Of MTR_SheetMetal) = New List(Of MTR_SheetMetal)()

    '====================

    'Sheet Metal
    Public routingSequenceSheetMetalList_BEC As List(Of RoutingSequence_SheetMetal) = New List(Of RoutingSequence_SheetMetal)()

    Public routingSequenceSheetMetalList_DGS As List(Of RoutingSequence_SheetMetal) = New List(Of RoutingSequence_SheetMetal)()

    'Structure
    Public routingSequenceStructureList_BEC As List(Of RoutingSequence_Structure) = New List(Of RoutingSequence_Structure)()

    Public routingSequenceStructureList_DGS As List(Of RoutingSequence_Structure) = New List(Of RoutingSequence_Structure)()

    'Misc
    Public routingSequenceMiscList_BEC As List(Of RoutingSequence_Structure) = New List(Of RoutingSequence_Structure)()

    Public routingSequenceMiscList_DGS As List(Of RoutingSequence_Structure) = New List(Of RoutingSequence_Structure)()


    ' Assembly
    Public routingSequenceAssemblyList_BEC As List(Of RoutingSequence_Assembly) = New List(Of RoutingSequence_Assembly)()

    Public routingSequenceAssemblyList_DGS As List(Of RoutingSequence_Assembly) = New List(Of RoutingSequence_Assembly)()


    Public mtcAssemblyList_DGS As List(Of MTC_Assembly) = New List(Of MTC_Assembly)()

    Public mtcBaseLineList_DGS As List(Of MTC_BaseLine) = New List(Of MTC_BaseLine)()

    Public mtcPartList_DGS As List(Of MTC_Part) = New List(Of MTC_Part)()

    Public mtcElectricalPartList_DGS As List(Of MTC_Electrical) = New List(Of MTC_Electrical)()

    Public mtcSheetMetalList_DGS As List(Of MTC_SheetMetal) = New List(Of MTC_SheetMetal)()


    Public mtrAssemblyList_DGS As List(Of MTR_Assembly) = New List(Of MTR_Assembly)()

    Public mtrPartList_DGS As List(Of MTR_Part) = New List(Of MTR_Part)()

    Public mtrSheetMetalList_DGS As List(Of MTR_SheetMetal) = New List(Of MTR_SheetMetal)()

    '====================

    '17th Sep 2024
    Public partPath As String = String.Empty

    Public sheetMetalPath As String = String.Empty

    Public dtCurrentPartData As System.Data.DataTable = New DataTable("Part Data")

    Public dtCurrentSheetMetalData As System.Data.DataTable = New DataTable("SheetMetal Data")

End Class
