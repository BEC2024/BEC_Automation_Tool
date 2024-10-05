Public Class MTC_BaseLine

    Public assemblyPath As String = String.Empty

    Public isValidBaseLineDirectoryPath As Boolean = False

    ''' <summary>
    ''' '0. Assembly Name
    ''' </summary>
    Public assemblyName As String = String.Empty

    ''' <summary>
    ''' 1. Is the part number match with M2M? *
    ''' </summary>
    Public isPartFound As String = String.Empty

    ''' <summary>
    ''' '2. What is the revision level? *
    ''' </summary>
    Public revisionLevel As String = String.Empty

    ''' <summary>
    ''' '3. Who is the Author of the file? (For BEC authors, type 3 capitalized initials of author)
    ''' </summary>
    Public author As String = String.Empty

    ''' <summary>
    ''' '4. What type of component?
    ''' </summary>
    Public category As String = String.Empty

    ''' <summary>
    ''' '5. Virtual thread applied for Fasteners?
    ''' </summary>
    Public isVirtualThreadExists As String = String.Empty

    ''' <summary>
    ''' '6. Are ALL features fully defined in the model Pathfinder? * (Sketch fully constraint?)
    ''' </summary>
    Public isSketchFullyDefined As String = String.Empty

    ''' <summary>
    ''' '7. Any suppressed feature found?
    ''' </summary>
    Public isSuppressFeatureFound As String = String.Empty


    Public materialSpec As String = String.Empty

    ''' <summary>
    ''' '8. Is the "Mat'l Spec" field populated? (PURCHASED for library components) *
    ''' </summary>
    Public isMaterialSpecExists As String = String.Empty


    Public materialUsed As String = String.Empty


    ''' <summary>
    ''' '9. Is the "Material Used" field populated? (PURCHASED for library components) *
    ''' </summary>
    Public isMaterialUsedExists As String = String.Empty

    Public title As String = String.Empty

    ''' <summary>
    ''' '10. Is the "Title" field populated with the correct part description? (This should MATCH the M2M Item Master (INV) Description field) *
    ''' </summary>
    Public isMaterialDesc_Title As String = String.Empty

    ''' <summary>
    '''  '11. Is the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) *
    ''' </summary>
    Public isAuthorExists As String = String.Empty

    '12. Is the "Keywords" field populated with the FULL M2M Item Master Description as shown in the Comments field? *

    Public comment As String = String.Empty

    ''' <summary>
    ''' '13. Is the "Comments" field populated with the Vendor name and Vendor part number? (It should appear as VENDOR NAME = VENDOR PART NUMBER) *
    ''' </summary>
    Public isCommentExist As String = String.Empty


    Public documentNumber As String = String.Empty

    ''' <summary>
    ''' '14. Is the "Document Number" field populated with the correct part number? (This should  MATCH the M2M Item Master Part Number field) *
    ''' </summary>
    Public isCorrectDocumentNumber As String = String.Empty

    ''' <summary>
    '''  '15. Is the "Revision" field populated with the correct revision number? *
    ''' </summary>
    Public isCorrectRevisionNumber As String = String.Empty

    ''' <summary>
    ''' '16. "Is the ""Project"" field populated with the correct Project Name? (The Project should be BROOKVILLE EQUIPMENT CORP) *"
    ''' </summary>
    Public isCorrectProjectName As String = String.Empty

    Public documentType As String = String.Empty

    Public hardwarePart As String = String.Empty

    ''' <summary>
    ''' 17. Is the "Hardware Part" box checked for hardware components? (This should be for ALL  nuts, bolts, washers, screws, etc.) *
    ''' </summary>
    Public isHardwarePartBoxChecked As String = String.Empty

    ''' <summary>
    ''' 18. Do all other unused property fields have a "dash" (-) populated? *
    ''' </summary>
    Public isDashPopulated As String = String.Empty

    Public UomProperty As String = String.Empty

    ''' <summary>
    ''' 19. Does the unit of measures (UOM) match M2M "For Inventory UOM"? (i.e. EA, sqft, in, etc.) *
    ''' </summary>
    Public isUOMMatch_M2M As String = String.Empty

    ''' <summary>
    ''' 20. Is the M2M Source marked stock/purchased? *
    ''' </summary>
    Public isM2MSourceStocked As String = String.Empty

    ''' <summary>
    ''' '21. Did I use the hole tool for ALL holes requiring fasteners and they match the Vendor supplied Data?(This include clearances holes For hardware, tapped holes, And Slots) *
    ''' </summary>
    Public isHoleToolUsed As String = String.Empty

    '22. Are all inter-part copies, part copies and included geometry broken? *

    ''' <summary>
    ''' '23. inter-part copies detected?
    ''' </summary>
    Public isInterPartCopiesDetected As String = String.Empty

    ''' <summary>
    ''' '24. part copies detected?
    ''' </summary>
    Public isPartCopiesDetected As String = String.Empty

    ''' <summary>
    ''' 25. broken  file Path detected?
    ''' </summary>
    Public isBrokenFilePathDetected As String = String.Empty

    ''' <summary>
    ''' '26. Is the model a SE Adjustable Part? (Models should NOT be adjustable) *
    ''' </summary>
    Public isAdjustable As String = String.Empty

    ''' <summary>
    ''' 27. Does this model have any SE Simplified Features? (NO model should contain SE Simplified Features) *
    ''' </summary>
    Public hasSESimplifiedFeature As String = String.Empty

    ''' <summary>
    ''' '28. Do all hardware AND library components in this asm model have an SE Status of Baselined?(If Available Or Released status hardware/library component exists, select "NO" ) *
    ''' </summary>
    Public hasSEStatusBaseLined As String = String.Empty


    ''' <summary>
    ''' 29. What is the last modified date?
    ''' </summary>
    Public modifiedDate As String = String.Empty

End Class