Public Class MTC_Electrical
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


    ''' <summary>
    ''' 11. What is the last modified date?
    ''' </summary>
    Public modifiedDate As String = String.Empty
End Class
