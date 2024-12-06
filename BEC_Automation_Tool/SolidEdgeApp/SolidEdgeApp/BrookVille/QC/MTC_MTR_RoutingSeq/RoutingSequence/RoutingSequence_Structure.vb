Public Class RoutingSequence_Structure

    Public assemblyPath As String = String.Empty

    Public assemblyName As String = String.Empty

    Public partNumber As String = String.Empty

    Public material As String = String.Empty

    Public materialSpec As String = String.Empty

    Public materialUsed As String = String.Empty

    Public holeFeature As String = String.Empty

    Public holeType As String = String.Empty

    Public holeQty As String = String.Empty

    Public m2mfSource As String = String.Empty

    Public isHoleFeature As String = "False"

    Public holeFit As String = String.Empty

    Public PMI As String = String.Empty

    Public projectName As String = String.Empty

    Public materialDescription As String = String.Empty

    Public filePath As String = String.Empty

    Public massItem As String = String.Empty

    Public m2mflocation As String = String.Empty

    Public m2mFbin As String = String.Empty

    ''' <summary>
    ''' In Some sheetmetal part there is no any sheetmetal feature,
    ''' So in this case document is not converted in sheet metal. So isValidPart is set to False for this kind of part
    ''' </summary>
    Public isValidPart As Boolean = True

    Public quantity As String = String.Empty

    Public category As String = String.Empty

End Class