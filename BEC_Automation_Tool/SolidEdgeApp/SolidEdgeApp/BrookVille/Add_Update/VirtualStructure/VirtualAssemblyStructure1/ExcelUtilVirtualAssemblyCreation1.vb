Imports Microsoft.Win32
Imports NLog

Public Class ExcelUtilVirtualAssemblyCreation1

    Public Enum MSApplications
        WORD
        ACCESS
        EXCEL
    End Enum
    Public Shared log As Logger = LogManager.GetCurrentClassLogger()
    Public Shared sheetName As String = "PUMP AND GEN" ' "3D MODEL STRUCTURE" 

    Public Shared Function isInstalled(ByVal App As MSApplications) As Boolean
        Dim strSubKey As String = String.Empty
        Select Case App
            Case MSApplications.ACCESS
                strSubKey = "Access.Application"

            Case MSApplications.EXCEL
                strSubKey = "Excel.Application"

            Case MSApplications.WORD
                strSubKey = "Word.Application"
        End Select

        Dim objKey As RegistryKey = Registry.ClassesRoot
        Dim objSubKey As RegistryKey = objKey.OpenSubKey(strSubKey)
        If objSubKey Is Nothing Then
            Return False
        Else
            Return True
        End If
        objKey.Close()
    End Function

    Public Shared Function ReadVirtualAssemblyCreationExcel2(ByRef excelFilePath As String) As List(Of VirtualAssemblyClass)

        Dim dictData As Dictionary(Of String, DataSet) = New Dictionary(Of String, DataSet)()
        Dim lstVirtualAssembly As List(Of VirtualAssemblyClass) = New List(Of VirtualAssemblyClass)()
        If Not ExcelUtil.IsInstalled(ExcelUtil.MSApplications.EXCEL) Then
            Return lstVirtualAssembly
        End If

        Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
        xlsWB = xlsApp.Workbooks.Open(excelFilePath, Nothing, False)

        Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim cn As Integer = 1
        Dim dsData As DataSet = New DataSet("Data")
        Dim objExcelUtil As ExcelUtil = New ExcelUtil()

        Dim rCnt As Integer = 2
        Dim mainasmlist As List(Of String) = New List(Of String)

        Try
            xlsSheet1 = xlsWB.Worksheets(sheetName)
            xlsCell1 = xlsSheet1.UsedRange

            Dim mainAssemblyList As List(Of String) = New List(Of String)()

            'Read Columns
            Dim dict As Dictionary(Of String, List(Of String)) = New Dictionary(Of String, List(Of String))

            For col = 2 To xlsCell1.Columns.Count
                Dim mainAssembly As String = xlsCell1.Cells(rCnt, col).value.ToString()
                If Not mainAssemblyList.Contains(mainAssembly) Then
                    mainAssemblyList.Add(mainAssembly)
                End If
            Next

            For Each assemblyName As String In mainAssemblyList

                Dim virtualAssemblyClassObj As VirtualAssemblyClass = New VirtualAssemblyClass()
                virtualAssemblyClassObj.mainAssemblyName = assemblyName

                Dim dicSubAssemblyDetails As Dictionary(Of String, List(Of String)) = New Dictionary(Of String, List(Of String))()

                For col = 2 To xlsCell1.Columns.Count

                    Dim childAsseList As List(Of String) = New List(Of String)()

                    If assemblyName = xlsCell1.Cells(rCnt, col).value.ToString() Then

                        Dim subAssemName As String = xlsCell1.Cells(3, col).value

                        For index = 5 To xlsCell1.Rows.Count
                            Dim childAssemName As String = xlsCell1.Cells(index, col).value
                            If childAssemName = String.Empty Then
                                Continue For
                            End If
                            childAsseList.Add(childAssemName)
                        Next

                        dicSubAssemblyDetails.Add(subAssemName, childAsseList)
                        virtualAssemblyClassObj.dict1 = dicSubAssemblyDetails

                    End If

                Next

                virtualAssemblyClassObj.dicSubAssemblyDetails.Add(assemblyName, virtualAssemblyClassObj)
                lstVirtualAssembly.Add(virtualAssemblyClassObj)
            Next

            'Read Data

            ' dsData = ReadRawMaterials(xlsSheet1, xlsCell1, xlsWB, sheetName)
            xlsApp.DisplayAlerts = False
            xlsWB.Save()
            xlsWB.Close()
            xlsApp.DisplayAlerts = True
            xlsApp.Quit()
        Catch ex As Exception
            CustomLogUtil.Log("While Reading Excel", ex.Message, ex.StackTrace)
            log.Error($"While Reading Excel{ex.Message}{ex.StackTrace}")
        End Try
        Return lstVirtualAssembly
    End Function

    Public Shared Function ReadVirtualAssemblyCreationExcel(ByRef excelFilePath As String, ByRef topLevelAssembly As String) As Dictionary(Of String, Dictionary(Of String, List(Of String)))
        Dim dicMainAssemblyDetails As Dictionary(Of String, Dictionary(Of String, List(Of String))) =
                New Dictionary(Of String, Dictionary(Of String, List(Of String)))()

        If Not ExcelUtil.IsInstalled(ExcelUtil.MSApplications.EXCEL) Then
            Return dicMainAssemblyDetails
        End If

        Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
        xlsWB = xlsApp.Workbooks.Open(excelFilePath, Nothing, False)

        Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim cn As Integer = 1
        Dim dsData As DataSet = New DataSet("Data")
        Dim objExcelUtil As ExcelUtil = New ExcelUtil()
        Dim rCnt As Integer = 2
        Dim mainasmlist As List(Of String) = New List(Of String)

        Try
            xlsSheet1 = xlsWB.Worksheets(sheetName)
            xlsCell1 = xlsSheet1.UsedRange

            topLevelAssembly = xlsCell1.Cells(1, 2).value.ToString()

            'MsgBox(topLevelAssembly)
            'Add columns
            Dim mainAssemblyList As List(Of String) = New List(Of String)()
            For col = 2 To xlsCell1.Columns.Count
                Dim mainAssembly As String = String.Empty
                If Not xlsCell1.Cells(rCnt, col).value.ToString() Is Nothing Then
                    mainAssembly = xlsCell1.Cells(rCnt, col).value.ToString().Replace("/", "")
                End If

                If Not mainAssemblyList.Contains(mainAssembly) Then
                    mainAssemblyList.Add(mainAssembly)
                End If
            Next

            ' Read sub assembly and part documents
            For Each assemblyName As String In mainAssemblyList
                Dim dicSubAssemblyDetails As Dictionary(Of String, List(Of String)) = New Dictionary(Of String, List(Of String))()

                If dicMainAssemblyDetails.ContainsKey(assemblyName) Then
                    dicSubAssemblyDetails = dicMainAssemblyDetails(assemblyName)
                End If

                For col = 2 To xlsCell1.Columns.Count

                    If assemblyName = xlsCell1.Cells(rCnt, col).value.ToString() Then

                        Dim childAsseList As List(Of String) = New List(Of String)()
                        Dim subAssemName As String = String.Empty

                        If Not xlsCell1.Cells(3, col).value Is Nothing Then
                            subAssemName = xlsCell1.Cells(3, col).value.Replace("/", "")
                        End If


                        If dicSubAssemblyDetails.ContainsKey(subAssemName) Then
                            childAsseList = dicSubAssemblyDetails(subAssemName)
                        End If

                        For index = 5 To xlsCell1.Rows.Count
                            Dim childAssemName As String = String.Empty
                            If Not xlsCell1.Cells(index, col).value Is Nothing Then
                                childAssemName = xlsCell1.Cells(index, col).value.Replace("/", "")
                            End If

                            If childAssemName = String.Empty Then
                                Continue For
                            End If
                            childAsseList.Add(childAssemName)
                        Next

                        If Not dicSubAssemblyDetails.ContainsKey(subAssemName) Then
                            dicSubAssemblyDetails.Add(subAssemName, childAsseList)
                        Else
                            dicSubAssemblyDetails(subAssemName) = childAsseList
                        End If

                    End If
                Next

                If Not dicMainAssemblyDetails.ContainsKey(assemblyName) Then
                    dicMainAssemblyDetails.Add(assemblyName, dicSubAssemblyDetails)
                Else
                    dicMainAssemblyDetails(assemblyName) = dicSubAssemblyDetails
                End If

            Next

            xlsApp.DisplayAlerts = False
            xlsWB.Save()
            xlsWB.Close()
            xlsApp.DisplayAlerts = True
            xlsApp.Quit()


        Catch ex As Exception
            CustomLogUtil.Log("While Reading Excel", ex.Message, ex.StackTrace)
            log.Error($"While Reading Excel{ex.Message}{ex.StackTrace}")
        End Try
        Return dicMainAssemblyDetails
    End Function

    Public Shared Function ReadVirtualAssemblyCreationExcelNew(ByRef excelFilePath As String, ByRef topLevelAssembly As String) As Dictionary(Of String, Dictionary(Of String, List(Of String)))
        Dim dicMainAssemblyDetails As Dictionary(Of String, Dictionary(Of String, List(Of String))) =
                New Dictionary(Of String, Dictionary(Of String, List(Of String)))()

        If Not ExcelUtil.IsInstalled(ExcelUtil.MSApplications.EXCEL) Then
            Return dicMainAssemblyDetails
        End If

        Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
        xlsWB = xlsApp.Workbooks.Open(excelFilePath, Nothing, False)

        Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim cn As Integer = 1
        Dim dsData As DataSet = New DataSet("Data")
        Dim objExcelUtil As ExcelUtil = New ExcelUtil()
        Dim rCnt As Integer = 2
        Dim mainasmlist As List(Of String) = New List(Of String)
        Dim refAssName As String = String.Empty
        Dim subAssemName1 As String = String.Empty
        Dim TopLevelAssemblyTitle = "/"
        Try
            xlsSheet1 = xlsWB.Worksheets(sheetName)
            xlsCell1 = xlsSheet1.UsedRange
            'MsgBox(xlsCell1.Cells(2, 2).value.ToString())
            topLevelAssembly = (xlsCell1.Cells(2, 3).value.ToString())
            TopLevelAssemblyTitle += (xlsCell1.Cells(2, 6).value.ToString())
            topLevelAssembly += TopLevelAssemblyTitle


            'MsgBox(topLevelAssembly)
            'Add columns
            Dim mainAssemblyList As List(Of String) = New List(Of String)()
            For col = 2 To xlsCell1.Columns.Count
                Dim mainAssembly As String = String.Empty
                'If Not xlsCell1.Cells(4, col).value.ToString() Is Nothing Then
                '    mainAssembly = xlsCell1.Cells(4, col).value.ToString().Replace("/", "")
                'End If
                If Not xlsCell1.Cells(3, col).value.ToString() Is Nothing Then
                    mainAssembly = xlsCell1.Cells(3, col).value.ToString().Replace("/", "")
                End If
                If Not mainAssemblyList.Contains(mainAssembly) Then
                    mainAssemblyList.Add(mainAssembly)
                End If
            Next


            '---------------------------------------------------
            ' Read sub assembly and part documents
            For Each assemblyName As String In mainAssemblyList

                Dim dicSubAssemblyDetails As Dictionary(Of String, List(Of String)) = New Dictionary(Of String, List(Of String))()
                Dim dicRefAssemblyDetails As Dictionary(Of String, Dictionary(Of String, List(Of String))) =
                New Dictionary(Of String, Dictionary(Of String, List(Of String)))()
                If dicRefAssemblyDetails.ContainsKey(assemblyName) Then
                    dicSubAssemblyDetails = dicRefAssemblyDetails(assemblyName)
                End If

                For col = 2 To xlsCell1.Columns.Count

                    'If assemblyName = xlsCell1.Cells(4, col).value.ToString() Then
                    If assemblyName = xlsCell1.Cells(3, col).value.ToString() Then
                        Dim childAsseList As List(Of String) = New List(Of String)()
                        Dim childAsseList1 As List(Of String) = New List(Of String)()
                        Dim subAssemName As String = String.Empty
                        refAssName = topLevelAssembly + " - Ref"
                        'If Not xlsCell1.Cells(3, col).value Is Nothing Then
                        '    subAssemName = xlsCell1.Cells(3, col).value
                        '    subAssemName1 = subAssemName
                        'End If
                        If Not xlsCell1.Cells(4, col).value Is Nothing Then
                            subAssemName = xlsCell1.Cells(4, col).value
                            subAssemName1 = subAssemName
                        End If

                        If dicSubAssemblyDetails.ContainsKey(subAssemName) Then 'refAssName
                            childAsseList = dicSubAssemblyDetails(subAssemName) 'refAssName
                        End If

                        For index = 5 To xlsCell1.Rows.Count
                            Dim skipname = "GROUPED IN SYSTEM"
                            Dim childAssemName As String = String.Empty
                            If Not xlsCell1.Cells(index, col).value Is Nothing Then
                                If Not xlsCell1.Cells(index, col).value.ToString.ToUpper = skipname.ToUpper Then
                                    childAssemName = xlsCell1.Cells(index, col).value.Replace("/", "")
                                End If
                            End If

                            If childAssemName = String.Empty Then
                                Continue For
                            End If
                            childAsseList.Add(childAssemName)
                        Next

                        'new child assembly
                        For col1 = 1 To xlsCell1.Columns.Count
                            If Not xlsCell1.Cells(4, col1).value Is Nothing Then
                                If Not childAsseList1.Contains(xlsCell1.Cells(4, col1).value) Then
                                    childAsseList1.Add(xlsCell1.Cells(4, col1).value)
                                End If
                            End If
                        Next


                        If Not dicSubAssemblyDetails.ContainsKey(subAssemName) Then
                            dicSubAssemblyDetails.Add(subAssemName, childAsseList)
                            dicSubAssemblyDetails.Add(refAssName, childAsseList1)
                        Else
                            dicSubAssemblyDetails(subAssemName) = childAsseList
                            dicSubAssemblyDetails(refAssName) = childAsseList1
                        End If

                    End If
                Next

                If Not dicMainAssemblyDetails.ContainsKey(assemblyName) Then
                    dicMainAssemblyDetails.Add(assemblyName, dicSubAssemblyDetails)
                Else
                    dicMainAssemblyDetails(assemblyName) = dicSubAssemblyDetails
                End If

            Next

            xlsApp.DisplayAlerts = False
            xlsWB.Save()
            xlsWB.Close()
            xlsApp.DisplayAlerts = True
            xlsApp.Quit()


        Catch ex As Exception
            CustomLogUtil.Log("While Reading Excel", ex.Message, ex.StackTrace)
            log.Error($"While Reading Excel{ex.Message}{ex.StackTrace}")
        End Try
        Return dicMainAssemblyDetails
    End Function
    'Public Shared Function ReadVirtualAssemblyCreationExcelNew(ByRef excelFilePath As String, ByRef topLevelAssembly As String) As Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, List(Of String))))
    '    Dim dicMainAssemblyDetails As Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, List(Of String)))) =
    '            New Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, List(Of String))))()


    '    If Not ExcelUtil.IsInstalled(ExcelUtil.MSApplications.EXCEL) Then
    '        Return dicMainAssemblyDetails
    '    End If

    '    Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
    '    Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
    '    xlsWB = xlsApp.Workbooks.Open(excelFilePath, Nothing, False)

    '    Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
    '    Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing
    '    Dim cn As Integer = 1
    '    Dim dsData As DataSet = New DataSet("Data")
    '    Dim objExcelUtil As ExcelUtil = New ExcelUtil()
    '    Dim rCnt As Integer = 2
    '    Dim mainasmlist As List(Of String) = New List(Of String)
    '    Dim refAssName As String = String.Empty
    '    Dim subAssemName1 As String = String.Empty
    '    Try
    '        xlsSheet1 = xlsWB.Worksheets(sheetName)
    '        xlsCell1 = xlsSheet1.UsedRange
    '        'MsgBox(xlsCell1.Cells(2, 2).value.ToString())
    '        topLevelAssembly = (xlsCell1.Cells(2, 2).value.ToString())



    '        'MsgBox(topLevelAssembly)
    '        'Add columns
    '        Dim mainAssemblyList As List(Of String) = New List(Of String)()
    '        For col = 2 To xlsCell1.Columns.Count
    '            Dim mainAssembly As String = String.Empty
    '            If Not xlsCell1.Cells(4, col).value.ToString() Is Nothing Then
    '                mainAssembly = xlsCell1.Cells(4, col).value.ToString().Replace("/", "")
    '            End If

    '            If Not mainAssemblyList.Contains(mainAssembly) Then
    '                mainAssemblyList.Add(mainAssembly)
    '            End If
    '        Next


    '        '---------------------------------------------------
    '        ' Read sub assembly and part documents
    '        For Each assemblyName As String In mainAssemblyList

    '            Dim dicSubAssemblyDetails As Dictionary(Of String, List(Of String)) = New Dictionary(Of String, List(Of String))()
    '            Dim dicRefAssemblyDetails As Dictionary(Of String, Dictionary(Of String, List(Of String))) =
    '            New Dictionary(Of String, Dictionary(Of String, List(Of String)))()
    '            If dicRefAssemblyDetails.ContainsKey(assemblyName) Then
    '                dicSubAssemblyDetails = dicRefAssemblyDetails(assemblyName)
    '            End If

    '            For col = 2 To xlsCell1.Columns.Count

    '                If assemblyName = xlsCell1.Cells(4, col).value.ToString() Then

    '                    Dim childAsseList As List(Of String) = New List(Of String)()
    '                    Dim subAssemName As String = String.Empty
    '                    refAssName = topLevelAssembly + " - Ref"
    '                    If Not xlsCell1.Cells(3, col).value Is Nothing Then
    '                        subAssemName = xlsCell1.Cells(3, col).value
    '                        subAssemName1 = subAssemName
    '                    End If


    '                    If dicSubAssemblyDetails.ContainsKey(subAssemName) Then 'refAssName
    '                        childAsseList = dicSubAssemblyDetails(subAssemName) 'refAssName
    '                    End If

    '                    For index = 5 To xlsCell1.Rows.Count
    '                        Dim skipname = "GROUPED IN SYSTEM"
    '                        Dim childAssemName As String = String.Empty
    '                        If Not xlsCell1.Cells(index, col).value Is Nothing Then
    '                            childAssemName = xlsCell1.Cells(index, col).value.Replace("/", "")
    '                            If xlsCell1.Cells(index, col).value.ToString.ToUpper = skipname.ToUpper Then
    '                                childAssemName = String.Empty
    '                            End If
    '                        End If

    '                        If childAssemName = String.Empty Then
    '                            Continue For
    '                        End If
    '                        childAsseList.Add(childAssemName)
    '                    Next

    '                    If Not dicSubAssemblyDetails.ContainsKey(subAssemName) Then 'refAssName
    '                        dicSubAssemblyDetails.Add(subAssemName, childAsseList) 'refAssName
    '                    Else
    '                        dicSubAssemblyDetails(subAssemName) = childAsseList 'refAssName
    '                    End If

    '                End If
    '            Next
    '            '-----------------------------------------------------------------------------------------------------------------


    '            If Not dicRefAssemblyDetails.ContainsKey(refAssName) Then
    '                dicRefAssemblyDetails.Add(refAssName, dicSubAssemblyDetails)
    '            Else
    '                dicRefAssemblyDetails(refAssName) = dicSubAssemblyDetails
    '            End If
    '            '-----------------------------------------------------------------------------------------------------------------

    '            If Not dicMainAssemblyDetails.ContainsKey(assemblyName) Then

    '                dicMainAssemblyDetails.Add(assemblyName, dicRefAssemblyDetails)


    '            Else
    '                dicMainAssemblyDetails(assemblyName) = dicRefAssemblyDetails
    '            End If

    '        Next

    '        xlsApp.DisplayAlerts = False
    '        xlsWB.Save()
    '        xlsWB.Close()
    '        xlsApp.DisplayAlerts = True
    '        xlsApp.Quit()


    '    Catch ex As Exception
    '        CustomLogUtil.Log("While Reading Excel", ex.Message, ex.StackTrace)
    '        log.Error($"While Reading Excel{ex.Message}{ex.StackTrace}")
    '    End Try
    '    Return dicMainAssemblyDetails
    'End Function
End Class