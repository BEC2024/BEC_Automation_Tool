Public Class MTR_Assembly

    ''' <summary>
    ''' 0. Assembly Name
    ''' </summary>
    Public assemblyName As String = String.Empty

    ''' <summary>
    ''' 1. Verify that there are NO adjustable parts present in the assembly model
    ''' </summary>
    Public isAdjustable As String = String.Empty

    ''' <summary>
    ''' 2. Verify that the inter-part copies are broken when released
    ''' </summary>
    Public isInterPartCopiesDetected As String = String.Empty

    ''' <summary>
    ''' 3. Verify that the part copies are broken when released
    ''' </summary>
    Public isPartCopiesDetected As String = String.Empty

    ''' <summary>
    ''' 4. Verify the all categories below have been populated with relative  Information Or, at a minimum, a dash (Summary, project, custom)
    ''' </summary>
    Public isValidAllCategories As String = String.Empty

    ''' <summary>
    ''' 5. Do assembly features exist within the assembly model?  If present, can they be removed?
    ''' </summary>
    Public isAssemblyFeatureExist As String = String.Empty

    ''' <summary>
    ''' 6. Verify that mating parts have been checked for interferences
    ''' </summary>
    Public isMatingPartInterferenceChecked As String = String.Empty

    ''' <summary>
    ''' 7. Verify that there are no interferences with objects in the environment
    ''' </summary>
    Public verifyInterference As String = String.Empty

    ''' <summary>
    ''' 8. Verify that the “Update on File Save” is UNCHECKED
    ''' </summary>
    Public verifyUpdateOnFileSave As String = String.Empty

    ''' <summary>
    ''' 9. Verify that the included geometry is broken when released
    ''' </summary>
    Public isGeometryBroken As String = String.Empty

    ''' <summary>
    ''' 10. Who is the author of file?
    ''' </summary>
    Public author As String = String.Empty


    ''' <summary>
    ''' 11. What is the modified date?
    ''' </summary>
    Public modifiedDate As String = String.Empty

End Class