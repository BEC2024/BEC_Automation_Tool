Public Class DocumentModel
    Public itemnumber As String = String.Empty
    Public fileNameWithoutExt As String = String.Empty
    Public fileName As String = String.Empty
    Public revisionNumber_Prop As String = String.Empty
    Public author As String = String.Empty
    Public documentno As String = String.Empty
    Public materialused As String = String.Empty
    Public matlspec As String = String.Empty
    Public fullpath As String = String.Empty
    Public lastsaved As String = String.Empty
    Public density As String = String.Empty
    Public projectname As String = String.Empty
    Public isadjustable As String = String.Empty
    Public properties As String = "Yes"
    Public customprop As String = String.Empty
    Public columname As String = String.Empty
    Public Documenttype As String = String.Empty
    Public adjustable As Boolean = False
    Public hardwarepart As Boolean
    Public isflatpattern As String = String.Empty
    Public iscutout As String = String.Empty
    Public revisionNumber_FileName As String = String.Empty
    Public authorcheck As String = String.Empty
    Public UomProperty As String = String.Empty
    Public Title As String = String.Empty
    Public comments As String = String.Empty
    Public category As String = String.Empty
    Public isElectrical As Boolean = False
    Public keywords As String = String.Empty
    Public SEfeatures As String = String.Empty
    Public partlistcount As Integer = 0
    Public interferencereport As String = String.Empty
    Public checkPartFeature As String = String.Empty
    Public checkAssemblyFeature As String = String.Empty
    Public gageeexcelfile As String = String.Empty
    Public issupress As String = "No"
    Public sketchisfullydefined As String = String.Empty
    Public allinterpartcopycheck As String = String.Empty
    Public statusfile As Boolean = False
    Public interference As Boolean = False
    Public ispartfound As Boolean = False
    Public ECO As String = String.Empty
    Public partCopiesDetected As String = "No"
    Public interPartCopiesDetected As String = "No"
    Public documentLinkBroken As String = "No"
    Public isBaseline As Boolean = False
    Public interPartLink As String = String.Empty
    Public isBrookVilleProject_Baseline As String = "No"
    Public isThreadExists As String = "No"
    Public isGeometryBroken As String = "Yes"
    Public isValidBaseLineDirectoryPath As Boolean = False

    Public modifiedDate As String = String.Empty
    Public isValidPart As Boolean = True

    Public material As String = String.Empty
    Public materialThickness As String = String.Empty
    Public bendRadius As String = String.Empty
    Public flat_Pattern_Model_CutSizeX As String = String.Empty
    Public flat_Pattern_Model_CutSizeY As String = String.Empty
    Public partNumber As String = String.Empty


    Public holeQty As String = "0"
    Public isHoleFeatureExists As String = "False"

    Public bendQty As String = "0"

    Public louverExists As String = "False"
    Public hemExists As String = "False"
    Public beadExists As String = "False"
    Public gussetExists As String = "False"

    Public hem_Bead_GussetExists As String = "False"

    Public holeFit As String = String.Empty

    Public massItem As String = String.Empty

    'Public materialDescription As String = String.Empty

    'Public filePath As String = String.Empty

    Public lastAuthor As String = String.Empty

    Public qAQC As String = String.Empty

    Public quantity As String = String.Empty

End Class