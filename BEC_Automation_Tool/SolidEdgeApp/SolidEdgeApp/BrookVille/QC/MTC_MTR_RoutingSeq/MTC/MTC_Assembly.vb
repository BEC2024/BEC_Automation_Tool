Public Class MTC_Assembly

    Public assemblyPath As String = String.Empty

    ''' <summary>
    ''' '0. Assembly Name
    ''' </summary>
    Public assemblyName As String = String.Empty

    ''' <summary>
    ''' 1. Eco Number
    ''' </summary>
    Public ecoNumber As String = String.Empty


    Public isPartFound As Boolean = False

    ''' <summary>
    ''' '2. Part Number
    ''' </summary>
    Public partNumber As String = String.Empty

    ''' <summary>
    ''' '3. Revision Level
    ''' </summary>
    Public revisionLevel As String = String.Empty

    ''' <summary>
    ''' '4. Author
    ''' </summary>
    Public author As String = String.Empty

    Public projectName As String = String.Empty

    ''' <summary>
    ''' '5. Project Name
    ''' </summary>
    Public projectNameExist As String = String.Empty


    ''' <summary>
    ''' '6. Revision Number Correct
    ''' </summary>
    Public revisionNumberCorrect As String = String.Empty


    Public documentNumber As String = String.Empty

    ''' <summary>
    ''' '7. Document Number correct
    ''' </summary>
    Public documentNumberCorrect As String = String.Empty



    ''' <summary>
    ''' '8. Is the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name)
    ''' </summary>
    Public authorExists As String = String.Empty

    ''' <summary>
    ''' '9. Do all technically unused properties have a "dash" populated?
    ''' </summary>
    Public isDashPopulated As String = String.Empty


    Public title As String = String.Empty

    ''' <summary>
    ''' '10. Does the Model Title MATCH the M2M Item Master (INV) Description Field?
    ''' </summary>
    Public isTitleMatch_ItemMaster As String = String.Empty


    Public UomProperty As String = String.Empty

    ''' <summary>
    ''' '11. Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
    ''' </summary>
    Public isUOMMatch_M2M As String = String.Empty

    ''' <summary>
    '''  '12. Perform a Parts List Report. How many total BOM items are in this assembly/weldment? (Not qty of all parts, only line items that will show up on draft PL) *
    ''' </summary>
    Public partListCount As String = String.Empty

    ''' <summary>
    ''' '13. Is interfernces found in assembly?
    ''' </summary>
    Public isInterferenceFound As String = String.Empty

    ''' <summary>
    ''' '14. inter-part copies detected
    ''' </summary>
    Public isInterPartCopiesDetected As String = String.Empty

    ''' <summary>
    ''' '15. part copies detected
    ''' </summary>
    Public isPartCopiesDetected As String = String.Empty

    ''' <summary>
    ''' '16. broken  file Path detected
    ''' </summary>
    Public isBrokenFilePathDetected As String = String.Empty

    ''' <summary>
    ''' 17. What is the last modified date?
    ''' </summary>
    Public modifiedDate As String = String.Empty


    ''' <summary>
    ''' '18. isadjustable We are not able to find the adjustable of assembly..
    ''' </summary>
    Public isAdjustable As String = String.Empty


    Public lastAuthor As String = String.Empty

End Class