Public Class MTC_Part

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

    Public materialSpec As String = String.Empty

    ''' <summary>
    ''' '12. Is the mat'l spec field populated? (indicated by RED in metadata audit file)
    ''' </summary>
    Public isMaterialSpecExists As String = String.Empty

    Public materialUsed As String = String.Empty

    ''' <summary>
    ''' '13. Is the material used field populated? *
    ''' </summary>
    Public isMaterialUsedExists As String = String.Empty

    ''' <summary>
    ''' '14. Did I use the 
    ''' tool for ALL holes requiring fasteners? (include clearance holes for hardware, tapped holes, And Slots) *
    ''' </summary>
    Public isHoleToolUsed As String = String.Empty

    ''' <summary>
    ''' 15. Is sketch fully constrain?
    ''' </summary>
    Public isSketchFullyConstraint As String = String.Empty

    ''' <summary>
    ''' 16. Have any suppressed (unused) features been removed from the model Pathfinder? *
    ''' </summary>
    Public haveSuppressedFeatureRemoved As String = String.Empty

    ''' <summary>
    ''' 17. Is the part "Adjustable"? (part should NOT be adjustable) *
    ''' </summary>
    Public isAdjustable As String = String.Empty


    ''' <summary>
    ''' 18. Does the M2M Source reflect part creation (Make child of stk/pur parent, Project  procurement strategy, material types Like composite, etc.)? *
    ''' </summary>
    Public m2mSource As String = String.Empty

    ''' <summary>
    ''' 19. What is the last modified date?
    ''' </summary>
    Public modifiedDate As String = String.Empty

    ''' <summary>
    ''' In Some  part there which has sheet metal feature, 
    ''' So in this case document is not converted in part. So isValidPart is set to False for this kind of part
    ''' </summary>
    Public isValidPart As Boolean = True

End Class