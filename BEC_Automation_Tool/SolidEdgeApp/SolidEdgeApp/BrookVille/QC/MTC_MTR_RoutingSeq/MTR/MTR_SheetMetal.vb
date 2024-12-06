Public Class MTR_SheetMetal

    ''' <summary>
    ''' 0. Assembly Name
    ''' </summary>
    Public assemblyName As String = String.Empty

    ''' <summary>
    ''' 1. Verify that ALL features have been fully constrained
    ''' </summary>
    Public isFeatureFullyConstrained As String = String.Empty

    ''' <summary>
    ''' 2. Verify that ALL suppressed and unused features have been removed
    ''' </summary>
    Public verifySuppressFeature As String = String.Empty

    ''' <summary>
    ''' 3. Verify that the part model is NOT adjustable
    ''' </summary>
    Public isAdjustable As String = String.Empty

    ''' <summary>
    ''' 4. Verify that the inter-part copies are broken when released
    ''' </summary>
    Public isInterPartCopiesDetected As String = String.Empty

    ''' <summary>
    ''' 5. Verify that the part copies are broken when released
    ''' </summary>
    Public isPartCopiesDetected As String = String.Empty

    ''' <summary>
    ''' 6. Verify the all categories below have been populated with relative  Information Or, at a minimum, a dash (Summary, project, custom)
    ''' </summary>
    Public isValidAllCategories As String = String.Empty

    ''' <summary>
    '''  7. Verify that the weight and mass has been applied
    ''' </summary>
    Public verifyWeightMass As String = String.Empty

    ''' <summary>
    ''' 8. Verify that the “Update on File Save” is UNCHECKED
    ''' </summary>
    Public verifyUpdateOnFileSave As String = String.Empty


    ''' <summary>
    ''' 9. Verify that the included geometry is broken when released
    ''' </summary>
    Public isGeometryBroken As String = String.Empty


    ''' <summary>
    ''' 10. Who is the author of file
    ''' </summary>
    Public author As String = String.Empty



    ''' <summary>
    ''' 11. What is the modified date?
    ''' </summary>
    Public modifiedDate As String = String.Empty


    ''' <summary>
    ''' In Some sheetmetal part there is no any sheetmetal feature, 
    ''' So in this case document is not converted in sheet metal. 
    ''' So isValidPart is set to False for this kind of part
    ''' </summary>
    Public isValidPart As Boolean = True

End Class