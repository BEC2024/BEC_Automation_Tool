Public Class MTC_SheetMetal

    Public assemblyPath As String = String.Empty

    ''' <summary>
    ''' '0. Assembly Name
    ''' </summary>
    Public assemblyName As String = String.Empty


    ''' <summary>
    ''' 1. Eco Number
    ''' </summary>
    Public ecoNumber As String = String.Empty

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
    ''' 14. Is the bend radius of the part equal to or above the ASTM minimum? *
    ''' </summary>
    Public gageExcelFile As String = String.Empty

    ''' <summary>
    '''  '15. . Is the flat pattern turned ON? (Applicable only to parts needing dxf or cut length) *
    ''' </summary>
    Public isFlatPatternActive As String = String.Empty


    ''' <summary>
    ''' '16. Did I use the 
    ''' tool for ALL holes requiring fasteners? (include clearance holes for        hardware, tapped holes, And Slots) *
    ''' </summary>
    Public holeToolsUsed As String = String.Empty


    ''' <summary>
    ''' '17. Is the part "Adjustable"? (part should NOT be adjustable) *
    ''' </summary>
    Public isAdjustatble As String = String.Empty


    ''' <summary>
    ''' 18. Does the M2M Source reflect part creation (Make child of stk/pur parent, Project        procurement strategy, material types Like composite, etc.)? *
    ''' </summary>
    Public m2mSource As String = String.Empty

    ''' <summary>
    ''' 19. What is the last modified date?
    ''' </summary>
    Public modifiedDate As String = String.Empty


    ''' <summary>
    ''' In Some sheetmetal part there is no any sheetmetal feature, 
    ''' So in this case document is not converted in sheet metal. So isValidPart is set to False for this kind of part
    ''' </summary>
    Public isValidPart As Boolean = True


    Public material As String = String.Empty
    Public materialThickness As String = String.Empty
    Public bendRadius As String = String.Empty
    Public flat_Pattern_Model_CutSizeX As String = String.Empty
    Public flat_Pattern_Model_CutSizeY As String = String.Empty



    Public holeQty As String = "0"
    Public holeFeatureExists As String = "False"

    Public bendQty As String = "0"

    Public louverExists As String = "False"
    Public hemExists As String = "False"
    Public beadExists As String = "False"
    Public gussetExists As String = "False"

    Public hem_Bead_GussetExists As String = "False"

    Public holeFit As String = String.Empty

    Public filePath As String = String.Empty
    Public materialDesc As String = String.Empty


End Class
