Public Class MTR_Part


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
    ''' 4. Verify the all categories below have been populated with relative  Information Or, at a minimum, a dash (Summary, project, custom)
    ''' </summary>
    Public isValidAllCategories As String = String.Empty

    ''' <summary>
    ''' 5. If the model is a fastener then the HARDWARE PART box should be checked
    ''' </summary>
    Public verifyFastenerHardwarePart As String = String.Empty

    ''' <summary>
    '''  6. Verify that the weight and mass has been applied
    ''' </summary>
    Public verifyWeightMass As String = String.Empty

    ''' <summary>
    ''' 7. Verify that the “Update on File Save” is UNCHECKED
    ''' </summary>
    Public verifyUpdateOnFileSave As String = String.Empty


    ''' <summary>
    ''' 8. Who is the author of file
    ''' </summary>
    Public author As String = String.Empty



    ''' <summary>
    ''' 9. What is the modified date?
    ''' </summary>
    Public modifiedDate As String = String.Empty



End Class
