Imports System.IO
Imports System.Runtime.InteropServices
Imports ExcelDataReader
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Win32
Imports NLog

Public Class ExcelUtil

#Region "Wait"

    Dim waitFormObj As Wait
    Public Shared log As Logger = LogManager.GetCurrentClassLogger()

    Public Enum MSApplications
        WORD
        ACCESS
        EXCEL
    End Enum

    Public Shared Function IsInstalled(ByVal App As MSApplications) As Boolean
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

    Public Sub WaitStartSave()
        '==Processing==
        Dim waitThread As System.Threading.Thread
        waitThread = New System.Threading.Thread(AddressOf LaunchWaitSave)
        waitThread.Start()
        Threading.Thread.Sleep(1000)
        waitFormObj.SetWaitMessage("Reading is in progress..")

        waitFormObj.SetProgressInformationVisibility(False)
        waitFormObj.SetProgressInformationMessage("")

        waitFormObj.SetProgressCountVisibility(False)
        waitFormObj.SetProgressCountMessage("0/0")
        '
    End Sub

    Public Sub LaunchWaitSave()
        waitFormObj = New Wait()
        waitFormObj.ShowDialog()
    End Sub

    Public Sub WaitEndSave()
        If waitFormObj IsNot Nothing Then
            waitFormObj.dispose2()
            waitFormObj = Nothing
        End If
    End Sub

#End Region

    Public Enum ExcelSheetName
        BECMaterials
    End Enum

    Public Enum ExcelSheetName2
        Template
    End Enum

    Public Enum ExcelSheetNameBOM
        Sheet1
    End Enum

    Public Enum ExcelSheetNameIntereference
        Sheet1
    End Enum

    Public Enum ExcelSheetNameHardware
        Sheet1
    End Enum

    'temp29Jan2021
    Public Enum ExcelSheetColumns
        Category
        Type
        Size
        Thickness
        Grade
        Material_Used
        Material_Specification
        BEC_Material
        Gage_Name
        Gage_Name_Original
        Bend_Radius
        Gage_Table
        Template
        Image
        Bend_Type
    End Enum

    Public Enum ExcelSheetColumns2
        Category
        Type
        Template
    End Enum

    Public Enum ExcelSheetColumnsBOM
        Category
        Material_Used
        Type
        BEC_Material
        Size
        Stock_Clearance
        Thickness
        Linear_Length
    End Enum

    Public Enum ExcelSheetColumnsIntereference
        Material
        Type
    End Enum

    Public Enum ExcelSheetColumnsHardware
        Hole_Size
        Hole_Standard
        Hole_Type
        Hardware_Type
        Size
        Length
        File_Name
    End Enum

    Public Enum ExcelSheetColumnsCreatePartStructure
        Category
        Type
        Size
        Grade
        Height
        Width
        Thickness
        Material_Used
        Material_Specification
        BEC_Material
        Template
        Description
        Gage_Table
        Gage_Name
        Bend_Radius
        Bend_Type
        Diameter
        Gap
        Linear_Length
        Priority
    End Enum

    Public Enum ExcelSheetNameCreatePart

        ' Part
        SheetMetal

        [Structure]

    End Enum

    Public Enum ExcelSheetNameM2M

        ' Part
        '  SheetMetal
        Sheet1

    End Enum

    Public Enum ExcelMtcReview
        Fac
        fpartno
        fdescript
        frev
        fmeasure
        fsource
        VendorName
        fKeywords
        flocation
        Fbin
    End Enum

    Public Shared Function Test(ByRef excelFilePath As String) As Dictionary(Of String, DataSet)

        Dim dictData As New Dictionary(Of String, DataSet)()
        'temp14Feb18
        'If Not CopyExcel() Then
        '    Return dictData
        'End If

        If Not ExcelUtil.IsInstalled(ExcelUtil.MSApplications.EXCEL) Then
            'Exit Function
            Return dictData
        End If

        Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
        xlsWB = xlsApp.Workbooks.Open(excelFilePath, Nothing, False)

        Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim cn As Integer = 1
        Dim dsData As New DataSet("Data")
        Dim objExcelUtil As New ExcelUtil()

        'objExcelUtil.waitStartSave()
        Try
            Dim sheetNames As Array
            sheetNames = System.Enum.GetNames(GetType(ExcelSheetName))
            Dim sheetName As String
            For Each sheetName In sheetNames

                'Check sheet exist in excel
                Dim sheet_found As Boolean = False
                For Each xs As Microsoft.Office.Interop.Excel.Worksheet In xlsWB.Sheets
                    If xs.Name = sheetName Then
                        sheet_found = True
                        Exit For
                    End If
                Next
                If sheet_found Then
                    dsData = ReadRawMaterials(xlsSheet1, xlsCell1, xlsWB, sheetName)
                    dictData.Add(sheetName, dsData)
                End If
                cn += 1
            Next

            xlsApp.DisplayAlerts = False
            xlsWB.Save()
            xlsWB.Close()
            xlsApp.DisplayAlerts = True
            xlsApp.Quit()

            'objExcelUtil.WaitEndSave()
        Catch ex As Exception

        End Try
        Return dictData
    End Function

    Public Shared Function MtcM2mFile(ByRef excelFilePath As String) As Dictionary(Of String, DataSet)

        Dim dictData As New Dictionary(Of String, DataSet)()

        If Not ExcelUtil.IsInstalled(ExcelUtil.MSApplications.EXCEL) Then
            'Exit Function
            Return dictData
        End If

        Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
        xlsWB = xlsApp.Workbooks.Open(excelFilePath, Nothing, False)

        Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim cn As Integer = 1
        Dim dsData As New DataSet("Data")
        Dim objExcelUtil As New ExcelUtil()

        'objExcelUtil.waitStartSave()
        Try
            Dim sheetNames As Array
            sheetNames = System.Enum.GetNames(GetType(ExcelSheetNameM2M))
            Dim sheetName As String
            For Each sheetName In sheetNames

                'Check sheet exist in excel
                Dim sheet_found As Boolean = False
                For Each xs As Microsoft.Office.Interop.Excel.Worksheet In xlsWB.Sheets
                    If xs.Name = sheetName Then
                        sheet_found = True
                        Exit For
                    End If
                Next

                If sheet_found Then
                    dsData = ReadM2Mfile(xlsSheet1, xlsCell1, xlsWB, sheetName)
                    dictData.Add(sheetName, dsData)
                End If
                cn += 1
            Next

            xlsApp.DisplayAlerts = False
            xlsWB.Save()
            xlsWB.Close()
            xlsApp.DisplayAlerts = True
            xlsApp.Quit()

            'objExcelUtil.WaitEndSave()
        Catch ex As Exception
            'MsgBox("While Reading MTC M2M File", ex.Message, ex.StackTrace)
        End Try
        Return dictData
    End Function

    Public Shared Function TestCreatePart(ByRef excelFilePath As String) As Dictionary(Of String, DataSet)

        Dim dictData As New Dictionary(Of String, DataSet)()
        Dim dsStructure As New DataSet("Structure")
        Dim dsSheetMetal As New DataSet("SheetMetal")
        Dim dsData As New DataSet("Data")
        Dim objExcelUtil As New ExcelUtil()
        FastExcelReadMaterials(excelFilePath, dsStructure, dsSheetMetal)
        dictData.Add("Structure", dsStructure)
        dictData.Add("SheetMetal", dsSheetMetal)
        'If Not ExcelUtil.IsInstalled(ExcelUtil.MSApplications.EXCEL) Then
        '    'Exit Function
        '    Return dictData
        'End If

        'Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
        'Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
        'xlsWB = xlsApp.Workbooks.Open(excelFilePath, Nothing, False)

        'Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        'Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing
        'Dim cn As Integer = 1
        ''Dim dsData As New DataSet("Data")
        ''Dim objExcelUtil As New ExcelUtil()

        ''Dim dsStructure As New DataSet("Structure")
        ''Dim dsSheetMetal As New DataSet("SheetMetal")

        ''objExcelUtil.waitStartSave()

        'Try
        '    Dim sheetNames As Array
        'sheetNames = System.Enum.GetNames(GetType(ExcelSheetNameCreatePart))

        '    Dim sheetName As String
        '    For Each sheetName In sheetNames

        '        'Check sheet exist in excel
        '        Dim sheet_found As Boolean = False
        '        For Each xs As Microsoft.Office.Interop.Excel.Worksheet In xlsWB.Sheets
        '            If xs.Name = sheetName Then
        '                sheet_found = True
        '                If xs.Name = "Structure" Then

        '                    dsStructure = ReadRawMaterialsCreatepart_Structure(xlsSheet1, xlsCell1, xlsWB, sheetName)
        '                    dictData.Add(sheetName, dsStructure)
        '                End If

        '                If xs.Name = "SheetMetal" Then
        '                    dsSheetMetal = ReadRawMaterialsCreatepart_Structure(xlsSheet1, xlsCell1, xlsWB, sheetName)
        '                    dictData.Add(sheetName, dsSheetMetal)
        '                End If
        '                'Exit For
        '            End If
        '        Next

        '        'If sheet_found Then
        '        '    dsData = ReadRawMaterialsCreatepart(xlsSheet1, xlsCell1, xlsWB, sheetName)
        '        '    dictData.Add(sheetName, dsData)
        '        'End If

        '        cn += 1
        '    Next

        '    xlsApp.DisplayAlerts = False
        '    xlsWB.Save()
        '    xlsWB.Close()
        '    xlsApp.DisplayAlerts = True
        '    xlsApp.Quit()

        '    'objExcelUtil.WaitEndSave()
        'Catch ex As Exception

        'End Try
        Return dictData
    End Function

    Public Shared Function TestHardware(ByRef excelFilePath As String) As Dictionary(Of String, DataSet)

        Dim dictData As New Dictionary(Of String, DataSet)()

        If Not ExcelUtil.IsInstalled(ExcelUtil.MSApplications.EXCEL) Then
            'Exit Function
            Return dictData
        End If

        Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
        xlsWB = xlsApp.Workbooks.Open(excelFilePath, Nothing, False)

        Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim cn As Integer = 1
        Dim dsData As New DataSet("Data")
        Dim objExcelUtil As New ExcelUtil()

        'objExcelUtil.waitStartSave()
        Try
            Dim sheetNames As Array
            sheetNames = System.Enum.GetNames(GetType(ExcelSheetNameHardware))
            Dim sheetName As String
            For Each sheetName In sheetNames

                'Check sheet exist in excel
                Dim sheet_found As Boolean = False
                For Each xs As Microsoft.Office.Interop.Excel.Worksheet In xlsWB.Sheets
                    If xs.Name = sheetName Then
                        sheet_found = True
                        Exit For
                    End If
                Next
                If sheet_found Then
                    dsData = ReadRawMaterialsHardware(xlsSheet1, xlsCell1, xlsWB, sheetName)
                    dictData.Add(sheetName, dsData)
                End If
                cn += 1
            Next

            xlsApp.DisplayAlerts = False
            xlsWB.Save()
            xlsWB.Close()
            xlsApp.DisplayAlerts = True
            xlsApp.Quit()

            'objExcelUtil.WaitEndSave()
        Catch ex As Exception

        End Try
        Return dictData
    End Function

    Public Shared Function TestIntereference(ByRef excelFilePath As String) As Dictionary(Of String, DataSet)

        Dim dictData As New Dictionary(Of String, DataSet)()

        If Not ExcelUtil.IsInstalled(ExcelUtil.MSApplications.EXCEL) Then
            'Exit Function
            Return dictData
        End If

        Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
        xlsWB = xlsApp.Workbooks.Open(excelFilePath, Nothing, False)

        Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim cn As Integer = 1
        Dim dsData As New DataSet("Data")
        Dim objExcelUtil As New ExcelUtil()

        'objExcelUtil.waitStartSave()
        Try
            Dim sheetNames As Array
            sheetNames = System.Enum.GetNames(GetType(ExcelSheetNameIntereference))
            Dim sheetName As String
            For Each sheetName In sheetNames

                'Check sheet exist in excel
                Dim sheet_found As Boolean = False
                For Each xs As Microsoft.Office.Interop.Excel.Worksheet In xlsWB.Sheets
                    If xs.Name = sheetName Then
                        sheet_found = True
                        Exit For
                    End If
                Next
                If sheet_found Then
                    dsData = ReadRawMaterialsIntereference(xlsSheet1, xlsCell1, xlsWB, sheetName)
                    dictData.Add(sheetName, dsData)
                End If
                cn += 1
            Next

            xlsApp.DisplayAlerts = False
            xlsWB.Save()
            xlsWB.Close()
            xlsApp.DisplayAlerts = True
            xlsApp.Quit()

            'objExcelUtil.WaitEndSave()
        Catch ex As Exception
            CustomLogUtil.Log("While Reading The Excel For Interference..", ex.Message, ex.StackTrace)
        End Try
        Return dictData
    End Function

    Public Shared Function TestBOM(ByRef excelFilePath As String) As Dictionary(Of String, DataSet)

        Dim dictData As New Dictionary(Of String, DataSet)()

        If Not ExcelUtil.IsInstalled(ExcelUtil.MSApplications.EXCEL) Then
            'Exit Function
            Return dictData
        End If

        Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
        xlsWB = xlsApp.Workbooks.Open(excelFilePath, Nothing, False)

        Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim cn As Integer = 1
        Dim dsData As New DataSet("Data")
        Dim objExcelUtil As New ExcelUtil()

        'objExcelUtil.waitStartSave()
        Try
            Dim sheetNames As Array
            sheetNames = System.Enum.GetNames(GetType(ExcelSheetNameBOM))
            Dim sheetName As String
            For Each sheetName In sheetNames

                'Check sheet exist in excel
                Dim sheet_found As Boolean = False
                For Each xs As Microsoft.Office.Interop.Excel.Worksheet In xlsWB.Sheets
                    If xs.Name = sheetName Then
                        sheet_found = True
                        Exit For
                    End If
                Next
                If sheet_found Then
                    dsData = ReadRawMaterialsBOM(xlsSheet1, xlsCell1, xlsWB, sheetName)
                    dictData.Add(sheetName, dsData)
                End If
                cn += 1
            Next

            xlsApp.DisplayAlerts = False
            xlsWB.Save()
            xlsWB.Close()
            xlsApp.DisplayAlerts = True
            xlsApp.Quit()

            'objExcelUtil.WaitEndSave()
        Catch ex As Exception
            CustomLogUtil.Log("While Reading TestBoom Data", ex.Message, ex.StackTrace)
        End Try
        Return dictData
    End Function

    Public Shared Function Test3(ByRef excelFilePath As String) As Dictionary(Of String, DataSet)

        Dim dictData As New Dictionary(Of String, DataSet)()
        'temp14Feb18
        'If Not CopyExcel() Then
        '    Return dictData
        'End If

        If Not ExcelUtil.IsInstalled(ExcelUtil.MSApplications.EXCEL) Then
            'Exit Function
            Return dictData
        End If

        Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
        xlsWB = xlsApp.Workbooks.Open(excelFilePath, Nothing, False)

        Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim cn As Integer = 1
        Dim dsData As New DataSet("Data")
        Dim objExcelUtil As New ExcelUtil()

        'objExcelUtil.waitStartSave()
        Try
            Dim sheetNames As Array
            sheetNames = System.Enum.GetNames(GetType(ExcelSheetName2))
            Dim sheetName As String
            For Each sheetName In sheetNames

                'Check sheet exist in excel
                Dim sheet_found As Boolean = False
                For Each xs As Microsoft.Office.Interop.Excel.Worksheet In xlsWB.Sheets
                    If xs.Name = sheetName Then
                        sheet_found = True
                        Exit For
                    End If
                Next
                If sheet_found Then
                    dsData = ReadRawMaterials31(xlsSheet1, xlsCell1, xlsWB, sheetName)
                    dictData.Add(sheetName, dsData)
                End If
                cn += 1
            Next

            xlsApp.DisplayAlerts = False
            xlsWB.Save()
            xlsWB.Close()
            xlsApp.DisplayAlerts = True
            xlsApp.Quit()

            'objExcelUtil.WaitEndSave()
        Catch ex As Exception

        End Try
        Return dictData
    End Function

    Public Shared Function ReadRawMaterials3(ByVal excelFilePath As String) As Dictionary(Of String, DataSet)

        Dim dictData As New Dictionary(Of String, DataSet)()
        Dim objExcelUtil As New ExcelUtil()
        objExcelUtil.WaitStartSave()
        Try
            Dim stopwatch As Stopwatch = Stopwatch.StartNew()

            dictData = Test3(excelFilePath)

            stopwatch.[Stop]()
            Console.WriteLine(stopwatch.ElapsedMilliseconds)

            'If Not CopyExcel() Then
            '    'Exit Function
            '    Return dictData
            'End If

            'If Not ExcelUtil.isInstalled(ExcelUtil.MSApplications.EXCEL) Then
            '    'Exit Function
            '    Return dictData
            'End If

            'Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
            'Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
            'xlsWB = xlsApp.Workbooks.Open(excelFilePath, Nothing, False)
            'Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            'Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing
            'Dim cn As Integer = 1
            'Dim dsData As DataSet = New DataSet("Data")

            'Dim objExcelUtil As ExcelUtil = New ExcelUtil()
            'objExcelUtil.waitStartSave()
            'Try
            '    Dim sheetNames As Array
            '    sheetNames = System.Enum.GetNames(GetType(excelSheetName))
            '    Dim sheetName As String
            '    For Each sheetName In sheetNames

            '        'Check sheet exist in excel
            '        Dim sheet_found As Boolean = False
            '        For Each xs As Microsoft.Office.Interop.Excel.Worksheet In xlsWB.Sheets
            '            If xs.Name = sheetName Then
            '                sheet_found = True
            '                Exit For
            '            End If
            '        Next
            '        If sheet_found Then
            '            'xlsSheet1 = Nothing
            '            'xlsCell1 = Nothing

            '            dsData = ReadRawMaterials(xlsSheet1, xlsCell1, xlsWB, sheetName)
            '            dictData.Add(sheetName, dsData)
            '        End If
            '        cn = cn + 1
            '    Next

            '    xlsWB.Save()
            '    xlsWB.Close()
            '    xlsApp.Quit()

            '    objExcelUtil.WaitEndSave()
        Catch ex As Exception
            objExcelUtil.WaitEndSave()
            MessageBox.Show("Error in reading excel : ", "Read excel", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CustomLogUtil.Log("Error in reading excel", ex.Message, ex.StackTrace)
        Finally
            objExcelUtil.WaitEndSave()
            'releaseObject(xlsSheet1)
            'releaseObject(xlsWB)
            'releaseObject(xlsCell1)
            'releaseObject(xlsApp)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try
        GlobalEntity.dictRawMaterials = dictData
        Return dictData
    End Function

    'temp04Dec2019
    Public Shared Function ReadRawMaterials2(ByVal excelFilePath As String) As Dictionary(Of String, DataSet)

        Dim dictData As New Dictionary(Of String, DataSet)()
        Dim objExcelUtil As New ExcelUtil()
        objExcelUtil.WaitStartSave()
        Try
            Dim stopwatch As Stopwatch = Stopwatch.StartNew()

            dictData = Test(excelFilePath)

            stopwatch.[Stop]()
            Console.WriteLine(stopwatch.ElapsedMilliseconds)

            'If Not CopyExcel() Then
            '    'Exit Function
            '    Return dictData
            'End If

            'If Not ExcelUtil.isInstalled(ExcelUtil.MSApplications.EXCEL) Then
            '    'Exit Function
            '    Return dictData
            'End If

            'Dim xlsApp As New Microsoft.Office.Interop.Excel.Application
            'Dim xlsWB As Microsoft.Office.Interop.Excel.Workbook
            'xlsWB = xlsApp.Workbooks.Open(excelFilePath, Nothing, False)
            'Dim xlsSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            'Dim xlsCell1 As Microsoft.Office.Interop.Excel.Range = Nothing
            'Dim cn As Integer = 1
            'Dim dsData As DataSet = New DataSet("Data")

            'Dim objExcelUtil As ExcelUtil = New ExcelUtil()
            'objExcelUtil.waitStartSave()
            'Try
            '    Dim sheetNames As Array
            '    sheetNames = System.Enum.GetNames(GetType(excelSheetName))
            '    Dim sheetName As String
            '    For Each sheetName In sheetNames

            '        'Check sheet exist in excel
            '        Dim sheet_found As Boolean = False
            '        For Each xs As Microsoft.Office.Interop.Excel.Worksheet In xlsWB.Sheets
            '            If xs.Name = sheetName Then
            '                sheet_found = True
            '                Exit For
            '            End If
            '        Next
            '        If sheet_found Then
            '            'xlsSheet1 = Nothing
            '            'xlsCell1 = Nothing

            '            dsData = ReadRawMaterials(xlsSheet1, xlsCell1, xlsWB, sheetName)
            '            dictData.Add(sheetName, dsData)
            '        End If
            '        cn = cn + 1
            '    Next

            '    xlsWB.Save()
            '    xlsWB.Close()
            '    xlsApp.Quit()

            '    objExcelUtil.WaitEndSave()
        Catch ex As Exception
            objExcelUtil.WaitEndSave()
            MessageBox.Show("Error in reading excel : ", "Read excel", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CustomLogUtil.Log("Error While reading excel", ex.Message, ex.StackTrace)
        Finally
            objExcelUtil.WaitEndSave()
            'releaseObject(xlsSheet1)
            'releaseObject(xlsWB)
            'releaseObject(xlsCell1)
            'releaseObject(xlsApp)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try
        GlobalEntity.dictRawMaterials = dictData
        Return dictData
    End Function

    Public Shared Function ReadRawMaterials2BOM(ByVal excelFilePath As String) As Dictionary(Of String, DataSet)

        Dim dictData As New Dictionary(Of String, DataSet)()
        Dim objExcelUtil As New ExcelUtil()
        objExcelUtil.WaitStartSave()
        Try
            Dim stopwatch As Stopwatch = Stopwatch.StartNew()

            dictData = TestBOM(excelFilePath)

            stopwatch.[Stop]()
            Console.WriteLine(stopwatch.ElapsedMilliseconds)
        Catch ex As Exception
            objExcelUtil.WaitEndSave()
            MessageBox.Show("Error in reading excel : ", "Read excel", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CustomLogUtil.Log("While Reading Raw material BOM Excel Data", ex.Message, ex.StackTrace)
        Finally
            objExcelUtil.WaitEndSave()

            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try
        GlobalEntity.dictRawMaterials = dictData
        Return dictData
    End Function

    Public Shared Function ReadRawMaterials2Interference(ByVal excelFilePath As String) As Dictionary(Of String, DataSet)

        Dim dictData As New Dictionary(Of String, DataSet)()
        Dim objExcelUtil As New ExcelUtil()
        objExcelUtil.WaitStartSave()
        Try
            Dim stopwatch As Stopwatch = Stopwatch.StartNew()

            dictData = TestIntereference(excelFilePath)

            stopwatch.[Stop]()
            Console.WriteLine(stopwatch.ElapsedMilliseconds)
        Catch ex As Exception
            objExcelUtil.WaitEndSave()
            MessageBox.Show("Error in reading excel : ", "Read excel", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            objExcelUtil.WaitEndSave()

            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try
        GlobalEntity.dictRawMaterials = dictData
        Return dictData
    End Function

    Public Shared Function ReadRawMaterials2_Hardware(ByVal excelFilePath As String) As Dictionary(Of String, DataSet)

        Dim dictData As New Dictionary(Of String, DataSet)()
        Dim objExcelUtil As New ExcelUtil()
        objExcelUtil.WaitStartSave()
        Try
            Dim stopwatch As Stopwatch = Stopwatch.StartNew()

            dictData = TestHardware(excelFilePath)

            stopwatch.[Stop]()
            Console.WriteLine(stopwatch.ElapsedMilliseconds)
        Catch ex As Exception
            objExcelUtil.WaitEndSave()
            MessageBox.Show("Error in reading excel : ", "Read excel", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            objExcelUtil.WaitEndSave()

            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try
        GlobalEntity.dictRawMaterials = dictData
        Return dictData
    End Function

    Public Shared Sub AddColumns(ByRef dt As System.Data.DataTable)
        Dim sheetColumnsNames As Array
        sheetColumnsNames = System.Enum.GetNames(GetType(ExcelSheetColumns))
        Dim columnName As String
        For Each columnName In sheetColumnsNames
            dt.Columns.Add(columnName, GetType(String))
        Next
    End Sub

    Public Shared Function ReadM2Mfile_CSV(ByVal csvFilePath As String) As Dictionary(Of String, DataSet)

        Dim dictData As New Dictionary(Of String, DataSet)()
        Dim objExcelUtil As New ExcelUtil()
        Try
            objExcelUtil.WaitStartSave()
            Dim dtM2MData As New Data.DataTable(ExcelSheetNameM2M.Sheet1.ToString())

            'Add columns of M2M review
            Dim sheetColumnsNames As Array
            sheetColumnsNames = System.Enum.GetNames(GetType(ExcelMtcReview))
            Dim columnName As String
            For Each columnName In sheetColumnsNames
                dtM2MData.Columns.Add(columnName, GetType(String))
            Next

            'Add the datatable in dataset
            Dim dsM2MData As New DataSet(ExcelSheetNameM2M.Sheet1.ToString())

            Using MyReader As New Microsoft.VisualBasic.
                           FileIO.TextFieldParser(csvFilePath)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(",")
                Dim currentRow As String()
                Dim lineCount As Integer = 1
                While Not MyReader.EndOfData
                    Try
                        currentRow = MyReader.ReadFields()
                        'Same columns for all types
                        Dim materialType As String = currentRow(2).ToString()
                        'If there is not any material type then skip
                        If materialType = String.Empty Then
                            Continue While
                        End If

                        If Not lineCount = 1 Then
                            Dim dr As DataRow = dtM2MData.NewRow()
                            AddRow(currentRow, dr, dtM2MData)
                        End If

                        lineCount += 1
                    Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
                    End Try
                End While
                Console.WriteLine($"**************Line Count {lineCount}")
            End Using
            dsM2MData.Tables.Add(dtM2MData)
            dictData.Add(ExcelSheetNameM2M.Sheet1.ToString(), dsM2MData)
        Catch ex As Exception
            objExcelUtil.WaitEndSave()
            Dim errMsg As String = "Error in reading csv : " + vbNewLine + vbNewLine + ex.Message + vbNewLine + ex.StackTrace
            MessageBox.Show(errMsg, "Read CSV", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CustomLogUtil.Log("While Reading M2M CSV File", ex.Message, ex.StackTrace)
        Finally
            objExcelUtil.WaitEndSave()
        End Try
        'SwAddin.dictRawMaterials = dictData
        Return dictData
    End Function

    Public Shared Sub AddRow(ByRef currentRow As String(), ByRef dr As DataRow, ByRef dt As System.Data.DataTable)

        dr(0) = currentRow(0).ToString()
        dr(1) = currentRow(1).ToString()
        dr(2) = currentRow(2).ToString()
        dr(3) = currentRow(3).ToString()
        dr(4) = currentRow(4).ToString()
        dr(5) = currentRow(5).ToString()
        dr(6) = currentRow(6).ToString()
        dr(7) = currentRow(7).ToString()
        dr(8) = currentRow(8).ToString()

        dt.Rows.Add(dr)
    End Sub

    Public Shared Function ReadM2Mfile(ByVal excelFilePath As String) As Dictionary(Of String, DataSet)

        Dim dictData As New Dictionary(Of String, DataSet)()
        Dim objExcelUtil As New ExcelUtil()
        objExcelUtil.WaitStartSave()
        Try
            Dim stopwatch As Stopwatch = Stopwatch.StartNew()

            dictData = MtcM2mFile(excelFilePath)

            stopwatch.[Stop]()
            Console.WriteLine(stopwatch.ElapsedMilliseconds)
        Catch ex As Exception
            objExcelUtil.WaitEndSave()
            MessageBox.Show("Error in reading excel : ", "Read excel", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CustomLogUtil.Log("Error In Reading Excel", ex.Message, ex.StackTrace)
        Finally
            objExcelUtil.WaitEndSave()

            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try
        GlobalEntity.dictRawMaterials = dictData
        Return dictData
    End Function

    Public Shared Function ReadMaterials(ByVal excelFilePath As String) As Dictionary(Of String, DataSet)

        Dim dictData As New Dictionary(Of String, DataSet)()
        Dim objExcelUtil As New ExcelUtil()
        objExcelUtil.WaitStartSave()
        Try
            Dim stopwatch As Stopwatch = Stopwatch.StartNew()

            dictData = TestCreatePart(excelFilePath)

            stopwatch.[Stop]()
            Console.WriteLine(stopwatch.ElapsedMilliseconds)
        Catch ex As Exception
            objExcelUtil.WaitEndSave()
            MessageBox.Show("Error in reading excel : ", "Read excel", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CustomLogUtil.Log("Error in reading BEC Material Excel", ex.Message, ex.StackTrace)
        Finally
            objExcelUtil.WaitEndSave()

            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try
        GlobalEntity.dictRawMaterials = dictData
        Return dictData
    End Function

    Public Shared Function GetColumnCount(ByVal columnName As String, ByRef xlsCell1 As Range) As Integer
        Dim colCount As Integer = 0
        Try
            For cCnt As Integer = 1 To xlsCell1.Columns.Count


                If xlsCell1.Cells(1, cCnt).value IsNot Nothing AndAlso xlsCell1.Cells(1, cCnt).value.ToString().Trim().ToUpper() = columnName.Trim().ToUpper() Then
                    Debug.Print(xlsCell1.Cells(1, cCnt).value.ToString())
                    colCount = cCnt
                    Exit For
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("Error in fetching column index of " + columnName, "Get column index", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return colCount
    End Function

    Public Shared Function ReadRawMaterials(ByRef xlsSheet1 As Worksheet, ByRef xlsCell1 As Range, ByRef xlsWB As Workbook, sheetName As String) As DataSet
        Dim ds As New DataSet("Data")
        'Dim cn As Integer = 0

        Dim dtExcelSheet As New Data.DataTable(sheetName)

        Dim sheetColumnsNames As Array
        sheetColumnsNames = System.Enum.GetNames(GetType(ExcelSheetColumns))
        Dim columnName As String
        For Each columnName In sheetColumnsNames
            dtExcelSheet.Columns.Add(columnName, GetType(String))
        Next

        Try
            xlsSheet1 = xlsWB.Worksheets(sheetName)
            xlsCell1 = xlsSheet1.UsedRange

            Dim cCntSW_Category As Integer = GetColumnCount(ExcelSheetColumns.Category.ToString, xlsCell1)
            Dim cCntSW_Type As Integer = GetColumnCount(ExcelSheetColumns.Type.ToString, xlsCell1)
            Dim cCntSW_Size As Integer = GetColumnCount(ExcelSheetColumns.Size.ToString, xlsCell1)
            Dim cCntSW_Thickness As Integer = GetColumnCount(ExcelSheetColumns.Thickness.ToString, xlsCell1)
            Dim cCnt_Grade As Integer = GetColumnCount(ExcelSheetColumns.Grade.ToString, xlsCell1)
            Dim cCntSW_MaterialUsed As Integer = GetColumnCount(ExcelSheetColumns.Material_Used.ToString, xlsCell1)

            Dim cCntSW_MaterialSpecification As Integer = GetColumnCount(ExcelSheetColumns.Material_Specification.ToString, xlsCell1)
            Dim cCntSW_BECMaterial As Integer = GetColumnCount(ExcelSheetColumns.BEC_Material.ToString, xlsCell1)
            Dim cCntSW_Gage_Name As Integer = GetColumnCount(ExcelSheetColumns.Gage_Name.ToString, xlsCell1)
            Dim cCntSW_Gage_Name_Original As Integer = GetColumnCount(ExcelSheetColumns.Gage_Name_Original.ToString, xlsCell1)

            Dim cCntSW_Bend_Radius As Integer = GetColumnCount(ExcelSheetColumns.Bend_Radius.ToString, xlsCell1)
            Dim cCntSW_Gage_Table As Integer = GetColumnCount(ExcelSheetColumns.Gage_Table.ToString, xlsCell1)

            Dim cCntSW_Template As Integer = GetColumnCount(ExcelSheetColumns.Template.ToString, xlsCell1)
            Dim cCntSW_Image As Integer = GetColumnCount(ExcelSheetColumns.Image.ToString, xlsCell1)
            Dim cCntSW_Bend_Type As Integer = GetColumnCount(ExcelSheetColumns.Bend_Type.ToString, xlsCell1)

            'temp30April2019
            'Here if value in excel is zero, then code consider it as NOTHING.
            'But we need zero value for creating code.
            'So we have added try catch to get the proper value
            For rCnt As Integer = 2 To xlsCell1.Rows.Count
                'cn = rCnt

                'SW_Template_Name
                Dim dr As DataRow = dtExcelSheet.NewRow()

                dr(ExcelSheetColumns.Category.ToString) = xlsCell1.Cells(rCnt, cCntSW_Category).value.ToString()

                dr(ExcelSheetColumns.Type.ToString) = xlsCell1.Cells(rCnt, cCntSW_Type).value.ToString()
                dr(ExcelSheetColumns.Size.ToString) = xlsCell1.Cells(rCnt, cCntSW_Size).value.ToString()

                If xlsCell1.Cells(rCnt, cCntSW_Thickness).value IsNot Nothing Then
                    dr(ExcelSheetColumns.Thickness.ToString) = xlsCell1.Cells(rCnt, cCntSW_Thickness).value.ToString()
                End If
                dr(ExcelSheetColumns.Grade.ToString) = xlsCell1.Cells(rCnt, cCnt_Grade).value.ToString()
                dr(ExcelSheetColumns.Material_Used.ToString) = xlsCell1.Cells(rCnt, cCntSW_MaterialUsed).value.ToString()
                dr(ExcelSheetColumns.Material_Specification.ToString) = xlsCell1.Cells(rCnt, cCntSW_MaterialSpecification).value.ToString()
                dr(ExcelSheetColumns.BEC_Material.ToString) = xlsCell1.Cells(rCnt, cCntSW_BECMaterial).value.ToString()
                dr(ExcelSheetColumns.Gage_Name.ToString) = xlsCell1.Cells(rCnt, cCntSW_Gage_Name).value.ToString()
                dr(ExcelSheetColumns.Gage_Name_Original.ToString) = xlsCell1.Cells(rCnt, cCntSW_Gage_Name_Original).value.ToString()
                dr(ExcelSheetColumns.Bend_Radius.ToString) = xlsCell1.Cells(rCnt, cCntSW_Bend_Radius).value.ToString()
                dr(ExcelSheetColumns.Gage_Table.ToString) = xlsCell1.Cells(rCnt, cCntSW_Gage_Table).value.ToString()
                dr(ExcelSheetColumns.Template.ToString) = xlsCell1.Cells(rCnt, cCntSW_Template).value.ToString()
                dr(ExcelSheetColumns.Image.ToString) = xlsCell1.Cells(rCnt, cCntSW_Image).value.ToString()
                dr(ExcelSheetColumns.Bend_Type.ToString) = xlsCell1.Cells(rCnt, cCntSW_Bend_Type).value.ToString()

                If Not (dr(ExcelSheetColumns.Type.ToString) = String.Empty _
                        AndAlso dr(ExcelSheetColumns.Size.ToString) = String.Empty _
                    AndAlso dr(ExcelSheetColumns.Thickness.ToString) = String.Empty _
                    AndAlso dr(ExcelSheetColumns.Grade.ToString) = String.Empty _
                    AndAlso dr(ExcelSheetColumns.Material_Used.ToString) = String.Empty _
                    AndAlso dr(ExcelSheetColumns.Material_Specification.ToString) = String.Empty _
                    AndAlso dr(ExcelSheetColumns.BEC_Material.ToString) = String.Empty _
                    AndAlso dr(ExcelSheetColumns.Gage_Name.ToString) = String.Empty _
                    AndAlso dr(ExcelSheetColumns.Bend_Radius.ToString) = String.Empty _
                    AndAlso dr(ExcelSheetColumns.Gage_Table.ToString) = String.Empty _
                    AndAlso dr(ExcelSheetColumns.Category.ToString) = String.Empty _
                    AndAlso dr(ExcelSheetColumns.Template.ToString) = String.Empty _
                    AndAlso dr(ExcelSheetColumns.Image.ToString) = String.Empty _
                    AndAlso dr(ExcelSheetColumns.Bend_Type.ToString) = String.Empty _
                    AndAlso dr(ExcelSheetColumns.Gage_Name_Original.ToString) = String.Empty
                   ) Then
                    dtExcelSheet.Rows.Add(dr)
                End If

            Next
            ds.Tables.Add(dtExcelSheet)
        Catch ex As Exception
            MessageBox.Show("Error in fetching excel details: ", "Fetch excel details", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CustomLogUtil.Log("Error in fetching excel details: ", ex.Message, ex.StackTrace)
        End Try

        Return ds
    End Function

    Public Shared Function ReadRawMaterialsHardware(ByRef xlsSheet1 As Worksheet, ByRef xlsCell1 As Range, ByRef xlsWB As Workbook, sheetName As String) As DataSet
        Dim ds As New DataSet("Data")
        'Dim cn As Integer = 0

        Dim dtExcelSheet As New Data.DataTable(sheetName)

        Dim sheetColumnsNames As Array
        sheetColumnsNames = System.Enum.GetNames(GetType(ExcelSheetColumnsHardware))
        Dim columnName As String
        For Each columnName In sheetColumnsNames
            dtExcelSheet.Columns.Add(columnName, GetType(String))
        Next

        Try
            xlsSheet1 = xlsWB.Worksheets(sheetName)
            xlsCell1 = xlsSheet1.UsedRange

            Dim cCntSW_Hole_Size As Integer = GetColumnCount(ExcelSheetColumnsHardware.Hole_Size.ToString, xlsCell1)
            Dim cCntSW_Hole_Standard As Integer = GetColumnCount(ExcelSheetColumnsHardware.Hole_Standard.ToString, xlsCell1)
            Dim cCntSW_Hole_Type As Integer = GetColumnCount(ExcelSheetColumnsHardware.Hole_Type.ToString, xlsCell1)
            Dim cCntSW_Hardware_Type As Integer = GetColumnCount(ExcelSheetColumnsHardware.Hardware_Type.ToString, xlsCell1)
            Dim cCntSW_Size As Integer = GetColumnCount(ExcelSheetColumnsHardware.Size.ToString, xlsCell1)
            Dim cCntSW_Length As Integer = GetColumnCount(ExcelSheetColumnsHardware.Length.ToString, xlsCell1)
            Dim cCntSW_File_Name As Integer = GetColumnCount(ExcelSheetColumnsHardware.File_Name.ToString, xlsCell1)

            'temp30April2019
            'Here if value in excel is zero, then code consider it as NOTHING.
            'But we need zero value for creating code.
            'So we have added try catch to get the proper value
            For rCnt As Integer = 2 To xlsCell1.Rows.Count
                'cn = rCnt

                'SW_Template_Name
                Dim dr As DataRow = dtExcelSheet.NewRow()
                If xlsCell1.Cells(rCnt, cCntSW_Hole_Size).value Is Nothing Then
                    Exit For
                End If
                dr(ExcelSheetColumnsHardware.Hole_Size.ToString) = xlsCell1.Cells(rCnt, cCntSW_Hole_Size).value.ToString()
                dr(ExcelSheetColumnsHardware.Hole_Standard.ToString) = xlsCell1.Cells(rCnt, cCntSW_Hole_Standard).value.ToString()
                dr(ExcelSheetColumnsHardware.Hole_Type.ToString) = xlsCell1.Cells(rCnt, cCntSW_Hole_Type).value.ToString()
                dr(ExcelSheetColumnsHardware.Hardware_Type.ToString) = xlsCell1.Cells(rCnt, cCntSW_Hardware_Type).value.ToString()
                dr(ExcelSheetColumnsHardware.Size.ToString) = xlsCell1.Cells(rCnt, cCntSW_Size).value.ToString()
                dr(ExcelSheetColumnsHardware.Length.ToString) = xlsCell1.Cells(rCnt, cCntSW_Length).value.ToString()
                dr(ExcelSheetColumnsHardware.File_Name.ToString) = xlsCell1.Cells(rCnt, cCntSW_File_Name).value.ToString()
                dtExcelSheet.Rows.Add(dr)

                'If Not (dr(excelSheetColumnsBOM.Type.ToString) = String.Empty _
                '        AndAlso dr(excelSheetColumnsBOM.Size.ToString) = String.Empty _
                '    AndAlso dr(excelSheetColumnsBOM.Material_Used.ToString) = String.Empty _
                '    AndAlso dr(excelSheetColumnsBOM.BEC_Material.ToString) = String.Empty) Then
                '    'AndAlso dr(excelSheetColumnsBOM.Stock_Clearance.ToString) = String.Empty
                '    dtExcelSheet.Rows.Add(dr)
                'End If

            Next
            ds.Tables.Add(dtExcelSheet)
        Catch ex As Exception
            MessageBox.Show("Error in fetching excel details: ", "Fetch excel details", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CustomLogUtil.Log("Error in fetching excel details", ex.Message, ex.StackTrace)
        End Try

        Return ds
    End Function

    Public Shared Function ReadM2Mfile(ByRef xlsSheet1 As Worksheet, ByRef xlsCell1 As Range, ByRef xlsWB As Workbook, sheetName As String) As DataSet
        Dim ds As New DataSet("Data")
        'Dim cn As Integer = 0

        Dim dtExcelSheet As New Data.DataTable(sheetName)

        Dim sheetColumnsNames As Array
        sheetColumnsNames = System.Enum.GetNames(GetType(ExcelMtcReview))
        Dim columnName As String
        For Each columnName In sheetColumnsNames
            dtExcelSheet.Columns.Add(columnName, GetType(String))
        Next

        Try
            xlsSheet1 = xlsWB.Worksheets(sheetName)
            xlsCell1 = xlsSheet1.UsedRange

            Dim partno As Integer = GetColumnCount(ExcelMtcReview.fpartno.ToString, xlsCell1)
            Dim descript As Integer = GetColumnCount(ExcelMtcReview.fdescript.ToString, xlsCell1)
            Dim rev As Integer = GetColumnCount(ExcelMtcReview.frev.ToString, xlsCell1)
            Dim measure As Integer = GetColumnCount(ExcelMtcReview.fmeasure.ToString, xlsCell1)
            Dim source As Integer = GetColumnCount(ExcelMtcReview.fsource.ToString, xlsCell1)
            Dim vendorname As Integer = GetColumnCount(ExcelMtcReview.VendorName.ToString, xlsCell1)
            Dim keywords As Integer = GetColumnCount(ExcelMtcReview.fKeywords.ToString, xlsCell1)

            'temp30April2019
            'Here if value in excel is zero, then code consider it as NOTHING.
            'But we need zero value for creating code.
            'So we have added try catch to get the proper value
            For rCnt As Integer = 2 To xlsCell1.Rows.Count
                'cn = rCnt

                'SW_Template_Name
                Dim dr As DataRow = dtExcelSheet.NewRow()
                If xlsCell1.Cells(rCnt, partno).value Is Nothing Then
                    Exit For
                End If

                dr(ExcelMtcReview.fpartno.ToString) = xlsCell1.Cells(rCnt, partno).value
                dr(ExcelMtcReview.fdescript.ToString) = xlsCell1.Cells(rCnt, descript).value
                dr(ExcelMtcReview.frev.ToString) = xlsCell1.Cells(rCnt, rev).value
                dr(ExcelMtcReview.fmeasure.ToString) = xlsCell1.Cells(rCnt, measure).value
                dr(ExcelMtcReview.fsource.ToString) = xlsCell1.Cells(rCnt, source).value
                dr(ExcelMtcReview.VendorName.ToString) = "" 'xlsCell1.Cells(rCnt, vendorname).value
                dr(ExcelMtcReview.fKeywords.ToString) = xlsCell1.Cells(rCnt, keywords).value

                dtExcelSheet.Rows.Add(dr)

            Next
            ds.Tables.Add(dtExcelSheet)
        Catch ex As Exception
            MessageBox.Show("Error in fetching excel details: ", "Fetch excel details", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CustomLogUtil.Log("Error in fetching excel details", ex.Message, ex.StackTrace)
        End Try

        Return ds
    End Function
    Public Shared Sub FastExcelReadMaterials(ByRef excelFilePath As String, ByRef dsStructure As DataSet, ByRef dsSheetMetal As DataSet)
        Dim ds As New DataSet("Data")

        Using stream As FileStream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read)
            Using reader As IExcelDataReader = ExcelReaderFactory.CreateOpenXmlReader(stream)
                Dim conf = New ExcelDataSetConfiguration With
                        {
                            .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration With
                            {
                                .UseHeaderRow = True
                            }
                        }
                ds = reader.AsDataSet(conf)

                dsStructure = ds.Copy()
                dsStructure.Tables.Remove("SheetMetal")
                'dsStructure.Tables.Remove("BEC DATA")
                'dsStructure.Tables.Remove("CODES")
                dsSheetMetal = ds.Copy()
                dsSheetMetal.Tables.Remove("Structure")
                'dsSheetMetal.Tables.Remove("BEC DATA")
                'dsSheetMetal.Tables.Remove("CODES")

            End Using
        End Using

        'Dim listOfColumns As New List(Of String)
        'For i = 0 To dsStructure.Tables(0).Columns.Count - 1
        '    listOfColumns.Add(dsStructure.Tables(0).Rows(0)(i).ToString())
        'Next
        'dsStructure.Tables(0).Rows(0).Delete()
        'For i = 0 To listOfColumns.Count - 1
        '    dsStructure.Tables(0).Columns(i).ColumnName = listOfColumns.Item(i)
        'Next
        'listOfColumns.Clear()
        'For i = 0 To dsSheetMetal.Tables(0).Columns.Count - 1
        '    listOfColumns.Add(dsSheetMetal.Tables(0).Rows(0)(i).ToString())
        'Next
        'dsSheetMetal.Tables(0).Rows(0).Delete()
        'For i = 0 To listOfColumns.Count - 1
        '    dsSheetMetal.Tables(0).Columns(i).ColumnName = listOfColumns.Item(i)
        'Next

    End Sub
    Public Shared Function ReadRawMaterialsCreatepart_Structure(ByRef xlsSheet1 As Worksheet, ByRef xlsCell1 As Range, ByRef xlsWB As Workbook, sheetName As String) As DataSet

        Dim ds As New DataSet("Data")
        'Dim cn As Integer = 0

        Dim dtExcelSheet As New Data.DataTable(sheetName)

        Dim sheetColumnsNames As Array
        sheetColumnsNames = System.Enum.GetNames(GetType(ExcelSheetColumnsCreatePartStructure))
        Dim columnName As String
        For Each columnName In sheetColumnsNames
            Dim dataColumn = dtExcelSheet.Columns.Add(columnName, GetType(String))
        Next

        Try
            xlsSheet1 = xlsWB.Worksheets(sheetName)
            xlsCell1 = xlsSheet1.UsedRange

            Dim Type As Integer = GetColumnCount(ExcelSheetColumnsCreatePartStructure.Type.ToString, xlsCell1)
            Dim Size As Integer = GetColumnCount(ExcelSheetColumnsCreatePartStructure.Size.ToString, xlsCell1)
            Dim Grade As Integer = GetColumnCount(ExcelSheetColumnsCreatePartStructure.Grade.ToString, xlsCell1)
            Dim Length As Integer = GetColumnCount(ExcelSheetColumnsCreatePartStructure.Height.ToString, xlsCell1)
            Dim Width As Integer = GetColumnCount(ExcelSheetColumnsCreatePartStructure.Width.ToString, xlsCell1)
            Dim Thickness As Integer = GetColumnCount(ExcelSheetColumnsCreatePartStructure.Thickness.ToString, xlsCell1)
            Dim Materialused As Integer = GetColumnCount(ExcelSheetColumnsCreatePartStructure.Material_Used.ToString, xlsCell1)
            Dim MaterialSpecification As Integer = GetColumnCount(ExcelSheetColumnsCreatePartStructure.Material_Specification.ToString, xlsCell1)
            Dim BecMaterial As Integer = GetColumnCount(ExcelSheetColumnsCreatePartStructure.BEC_Material.ToString, xlsCell1)
            Dim Template As Integer = GetColumnCount(ExcelSheetColumnsCreatePartStructure.Template.ToString, xlsCell1)
            Dim Description As Integer = GetColumnCount(ExcelSheetColumnsCreatePartStructure.Description.ToString, xlsCell1)
            Dim Diameter As Integer = GetColumnCount(ExcelSheetColumnsCreatePartStructure.Diameter.ToString, xlsCell1)
            Dim Gap As Integer = GetColumnCount(ExcelSheetColumnsCreatePartStructure.Gap.ToString, xlsCell1)
            Dim Gage_Table As Integer = 0
            Dim Gage_Name As Integer = 0
            Dim Bend_Radius As Integer = 0
            Dim Bend_Type As Integer = 0
            Dim Category As Integer = GetColumnCount(ExcelSheetColumnsCreatePartStructure.Category.ToString, xlsCell1)
            Dim LinearLength As Integer = GetColumnCount(ExcelSheetColumnsCreatePartStructure.Linear_Length.ToString, xlsCell1)


            If sheetName = "SheetMetal" Then
                Gage_Table = GetColumnCount(ExcelSheetColumnsCreatePartStructure.Gage_Table.ToString, xlsCell1)
                Gage_Name = GetColumnCount(ExcelSheetColumnsCreatePartStructure.Gage_Name.ToString, xlsCell1)
                Bend_Radius = GetColumnCount(ExcelSheetColumnsCreatePartStructure.Bend_Radius.ToString, xlsCell1)
                Bend_Type = GetColumnCount(ExcelSheetColumnsCreatePartStructure.Bend_Type.ToString, xlsCell1)

            End If
            'Dim Gage_Table As Integer = GetColumnCount(excelSheetColumnsCreatePart.Gage_Table.ToString, xlsCell1)

            'temp30April2019
            'Here if value in excel is zero, then code consider it as NOTHING.
            'But we need zero value for creating code.
            'So we have added try catch to get the proper value

            For rCnt As Integer = 2 To xlsCell1.Rows.Count
                'cn = rCnt

                'SW_Template_Name
                Dim dr As DataRow = dtExcelSheet.NewRow()
                If xlsCell1.Cells(rCnt, Type).value Is Nothing Then
                    Exit For
                End If
                'If xlsCell1.Cells(rCnt, BecMaterial).value Is Nothing Then
                '    dr(excelSheetColumnsCreatePart.BEC_Material.ToString) = xlsCell1.Cells(rCnt, BecMaterial).value
                'End If


                dr(ExcelSheetColumnsCreatePartStructure.Category.ToString) = xlsCell1.Cells(rCnt, Category).value
                dr(ExcelSheetColumnsCreatePartStructure.Type.ToString) = xlsCell1.Cells(rCnt, Type).value

                dr(ExcelSheetColumnsCreatePartStructure.Grade.ToString) = xlsCell1.Cells(rCnt, Grade).value.ToString()

                dr(ExcelSheetColumnsCreatePartStructure.Thickness.ToString) = xlsCell1.Cells(rCnt, Thickness).value
                dr(ExcelSheetColumnsCreatePartStructure.Material_Used.ToString) = xlsCell1.Cells(rCnt, Materialused).value
                dr(ExcelSheetColumnsCreatePartStructure.Material_Specification.ToString) = xlsCell1.Cells(rCnt, MaterialSpecification).value
                dr(ExcelSheetColumnsCreatePartStructure.BEC_Material.ToString) = xlsCell1.Cells(rCnt, BecMaterial).value

                If sheetName = "SheetMetal" Then
                    dr(ExcelSheetColumnsCreatePartStructure.Category.ToString) = "SheetMetal"
                    dr(ExcelSheetColumnsCreatePartStructure.Gage_Table.ToString) = xlsCell1.Cells(rCnt, Gage_Table).value
                    dr(ExcelSheetColumnsCreatePartStructure.Gage_Name.ToString) = xlsCell1.Cells(rCnt, Gage_Name).value
                    dr(ExcelSheetColumnsCreatePartStructure.Bend_Radius.ToString) = xlsCell1.Cells(rCnt, Bend_Radius).value
                    dr(ExcelSheetColumnsCreatePartStructure.Bend_Type.ToString) = xlsCell1.Cells(rCnt, Bend_Type).value
                    dr(ExcelSheetColumnsCreatePartStructure.Size.ToString) = xlsCell1.Cells(rCnt, Size).value
                Else
                    dr(ExcelSheetColumnsCreatePartStructure.Category.ToString) = "Structure"
                    dr(ExcelSheetColumnsCreatePartStructure.Size.ToString) = xlsCell1.Cells(rCnt, Size).value
                    'dr(ExcelSheetColumnsCreatePartStructure.Grade.ToString) = xlsCell1.Cells(rCnt, Grade).value
                    dr(ExcelSheetColumnsCreatePartStructure.Height.ToString) = xlsCell1.Cells(rCnt, Length).value
                    dr(ExcelSheetColumnsCreatePartStructure.Width.ToString) = xlsCell1.Cells(rCnt, Width).value
                    dr(ExcelSheetColumnsCreatePartStructure.Template.ToString) = xlsCell1.Cells(rCnt, Template).value
                    dr(ExcelSheetColumnsCreatePartStructure.Description.ToString) = xlsCell1.Cells(rCnt, Description).value
                    dr(ExcelSheetColumnsCreatePartStructure.Gap.ToString) = xlsCell1.Cells(rCnt, Gap).value
                    dr(ExcelSheetColumnsCreatePartStructure.Diameter.ToString) = xlsCell1.Cells(rCnt, Diameter).value
                    dr(ExcelSheetColumnsCreatePartStructure.Template.ToString) = xlsCell1.Cells(rCnt, Template).value
                    dr(ExcelSheetColumnsCreatePartStructure.Linear_Length.ToString) = xlsCell1.Cells(rCnt, LinearLength).value
                End If
                '
                dtExcelSheet.Rows.Add(dr)

                'If Not (dr(excelSheetColumnsBOM.Type.ToString) = String.Empty _
                '        AndAlso dr(excelSheetColumnsBOM.Size.ToString) = String.Empty _
                '    AndAlso dr(excelSheetColumnsBOM.Material_Used.ToString) = String.Empty _
                '    AndAlso dr(excelSheetColumnsBOM.BEC_Material.ToString) = String.Empty) Then
                '    'AndAlso dr(excelSheetColumnsBOM.Stock_Clearance.ToString) = String.Empty
                '    dtExcelSheet.Rows.Add(dr)
                'End If

            Next
            ds.Tables.Add(dtExcelSheet)
        Catch ex As Exception
            MessageBox.Show("Error in fetching excel details: ", "Fetch excel details", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CustomLogUtil.Log($"While Reading Raw Materials Sheet : {sheetName}", ex.Message, ex.StackTrace)
        End Try

        Return ds
    End Function

    Public Shared Function ReadRawMaterialsIntereference(ByRef xlsSheet1 As Worksheet, ByRef xlsCell1 As Range, ByRef xlsWB As Workbook, sheetName As String) As DataSet

        Dim ds As New DataSet("Data")
        'Dim cn As Integer = 0

        Dim dtExcelSheet As New Data.DataTable(sheetName)

        Dim sheetColumnsNames As Array
        sheetColumnsNames = System.Enum.GetNames(GetType(ExcelSheetColumnsIntereference))
        Dim columnName As String

        'myDataColumn1 = New DataColumn()
        Dim myDataColumn1 As New DataColumn With {
            .ColumnName = "Select",
            .DefaultValue = "0",
            .DataType = System.Type.GetType("System.Boolean")
        }
        dtExcelSheet.Columns.Add(myDataColumn1)

        For Each columnName In sheetColumnsNames
            dtExcelSheet.Columns.Add(columnName, GetType(String))
        Next

        'Dim column As DataColumn = New DataColumn()
        'column.DataType = System.Type.[GetType]("System.Int32")
        'column.Caption = "Sr"
        'column.ColumnName = "Sr"
        'column.AutoIncrement = True
        'column.AutoIncrementSeed = 1
        'column.AutoIncrementStep = 1
        'dtExcelSheet.Columns.Add(column)

        Try
            xlsSheet1 = xlsWB.Worksheets(sheetName)
            xlsCell1 = xlsSheet1.UsedRange

            Dim cCntSW_Material As Integer = GetColumnCount(ExcelSheetColumnsIntereference.Material.ToString, xlsCell1)
            Dim cCntSW_Type As Integer = GetColumnCount(ExcelSheetColumnsIntereference.Type.ToString, xlsCell1)

            'temp30April2019
            'Here if value in excel is zero, then code consider it as NOTHING.
            'But we need zero value for creating code.
            'So we have added try catch to get the proper value

            For rCnt As Integer = 2 To xlsCell1.Rows.Count
                'cn = rCnt

                'SW_Template_Name
                Dim dr As DataRow = dtExcelSheet.NewRow()
                If xlsCell1.Cells(rCnt, cCntSW_Material).value Is Nothing Then
                    Exit For
                End If
                dr(ExcelSheetColumnsIntereference.Material.ToString) = xlsCell1.Cells(rCnt, cCntSW_Material).value.ToString()
                dr(ExcelSheetColumnsIntereference.Type.ToString) = xlsCell1.Cells(rCnt, cCntSW_Type).value.ToString()

                If Not (dr(ExcelSheetColumnsIntereference.Material.ToString) = String.Empty _
                        AndAlso dr(ExcelSheetColumnsIntereference.Type.ToString) = String.Empty) Then
                    dtExcelSheet.Rows.Add(dr)
                End If

            Next
            ds.Tables.Add(dtExcelSheet)
        Catch ex As Exception
            MessageBox.Show("Error in fetching excel details: ", "Fetch excel details", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CustomLogUtil.Log("While Reading Raw Materials Intereference", ex.Message, ex.StackTrace)
        End Try

        Return ds
    End Function

    Public Shared Function ReadRawMaterialsBOM(ByRef xlsSheet1 As Worksheet, ByRef xlsCell1 As Range, ByRef xlsWB As Workbook, sheetName As String) As DataSet
        Dim ds As New DataSet("Data")
        'Dim cn As Integer = 0

        Dim dtExcelSheet As New Data.DataTable(sheetName)

        Dim sheetColumnsNames As Array
        sheetColumnsNames = System.Enum.GetNames(GetType(ExcelSheetColumnsBOM))
        Dim columnName As String
        For Each columnName In sheetColumnsNames
            dtExcelSheet.Columns.Add(columnName, GetType(String))
        Next

        Try
            xlsSheet1 = xlsWB.Worksheets(sheetName)
            xlsCell1 = xlsSheet1.UsedRange

            Dim cCntSW_Material_Used As Integer = GetColumnCount(ExcelSheetColumnsBOM.Material_Used.ToString, xlsCell1)
            Dim cCntSW_Type As Integer = GetColumnCount(ExcelSheetColumnsBOM.Type.ToString, xlsCell1)
            Dim cCntSW_BECMaterial As Integer = GetColumnCount(ExcelSheetColumnsBOM.BEC_Material.ToString, xlsCell1)
            Dim cCntSW_Size As Integer = GetColumnCount(ExcelSheetColumnsBOM.Size.ToString, xlsCell1)
            Dim cCntSW_Stock_Clearance As Integer = GetColumnCount(ExcelSheetColumnsBOM.Stock_Clearance.ToString, xlsCell1)

            'temp30April2019
            'Here if value in excel is zero, then code consider it as NOTHING.
            'But we need zero value for creating code.
            'So we have added try catch to get the proper value
            For rCnt As Integer = 2 To xlsCell1.Rows.Count
                'cn = rCnt

                'SW_Template_Name
                Dim dr As DataRow = dtExcelSheet.NewRow()
                If xlsCell1.Cells(rCnt, cCntSW_Material_Used).value Is Nothing Then
                    Exit For
                End If
                dr(ExcelSheetColumnsBOM.Material_Used.ToString) = xlsCell1.Cells(rCnt, cCntSW_Material_Used).value.ToString()

                dr(ExcelSheetColumnsBOM.Type.ToString) = xlsCell1.Cells(rCnt, cCntSW_Type).value.ToString()
                dr(ExcelSheetColumnsBOM.Size.ToString) = xlsCell1.Cells(rCnt, cCntSW_Size).value.ToString()
                dr(ExcelSheetColumnsBOM.Stock_Clearance.ToString) = xlsCell1.Cells(rCnt, cCntSW_Stock_Clearance).value.ToString()

                If xlsCell1.Cells(rCnt, cCntSW_BECMaterial).value IsNot Nothing Then
                    dr(ExcelSheetColumnsBOM.BEC_Material.ToString) = xlsCell1.Cells(rCnt, cCntSW_BECMaterial).value.ToString()
                End If

                If Not (dr(ExcelSheetColumnsBOM.Type.ToString) = String.Empty _
                        AndAlso dr(ExcelSheetColumnsBOM.Size.ToString) = String.Empty _
                    AndAlso dr(ExcelSheetColumnsBOM.Material_Used.ToString) = String.Empty _
                    AndAlso dr(ExcelSheetColumnsBOM.BEC_Material.ToString) = String.Empty) Then
                    'AndAlso dr(excelSheetColumnsBOM.Stock_Clearance.ToString) = String.Empty
                    dtExcelSheet.Rows.Add(dr)
                End If

            Next
            ds.Tables.Add(dtExcelSheet)
        Catch ex As Exception
            MessageBox.Show("Error in fetching excel details: ", "Fetch excel details", MessageBoxButtons.OK, MessageBoxIcon.Error)
            log.Error($"While Fetchng data for Raw Material Boom from excel {ex.Message} {ex.StackTrace}")
            CustomLogUtil.Log("While Fetchng data for Raw Material Boom from excel", ex.Message, ex.StackTrace)
        End Try

        Return ds
    End Function

    Public Shared Function ReadRawMaterials31(ByRef xlsSheet1 As Worksheet, ByRef xlsCell1 As Range, ByRef xlsWB As Workbook, sheetName As String) As DataSet
        Dim ds As New DataSet("Data")
        'Dim cn As Integer = 0
        Dim dtExcelSheet As New Data.DataTable(sheetName)

        Dim sheetColumnsNames As Array
        sheetColumnsNames = System.Enum.GetNames(GetType(ExcelSheetColumns2))
        Dim columnName As String
        For Each columnName In sheetColumnsNames
            dtExcelSheet.Columns.Add(columnName, GetType(String))
        Next

        Try
            xlsSheet1 = xlsWB.Worksheets(sheetName)
            xlsCell1 = xlsSheet1.UsedRange

            Dim cCntSW_Category As Integer = GetColumnCount(ExcelSheetColumns2.Category.ToString, xlsCell1)
            Dim cCntSW_Type As Integer = GetColumnCount(ExcelSheetColumns2.Type.ToString, xlsCell1)
            Dim cCntSW_Template As Integer = GetColumnCount(ExcelSheetColumns2.Template.ToString, xlsCell1)

            'temp30April2019
            'Here if value in excel is zero, then code consider it as NOTHING.
            'But we need zero value for creating code.
            'So we have added try catch to get the proper value
            For rCnt As Integer = 2 To xlsCell1.Rows.Count
                'cn = rCnt

                'SW_Template_Name
                Dim dr As DataRow = dtExcelSheet.NewRow()

                dr(ExcelSheetColumns2.Category.ToString) = xlsCell1.Cells(rCnt, cCntSW_Category).value.ToString()
                dr(ExcelSheetColumns2.Type.ToString) = xlsCell1.Cells(rCnt, cCntSW_Type).value.ToString()
                dr(ExcelSheetColumns2.Template.ToString) = xlsCell1.Cells(rCnt, cCntSW_Template).value.ToString()

                If Not (dr(ExcelSheetColumns2.Category.ToString) = String.Empty _
                        AndAlso dr(ExcelSheetColumns2.Type.ToString) = String.Empty _
                    AndAlso dr(ExcelSheetColumns2.Template.ToString) = String.Empty) Then
                    dtExcelSheet.Rows.Add(dr)
                End If

            Next
            ds.Tables.Add(dtExcelSheet)
        Catch ex As Exception
            MessageBox.Show("Error in fetching excel details: ", "Fetch excel details", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CustomLogUtil.Log("Error in fetching excel details:", ex.Message, ex.StackTrace)

        End Try

        Return ds
    End Function

    Public Shared Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers() 'temp14Dec18
        End Try
    End Sub

    Private Function IsXlsxVersion(ByVal version As String) As Boolean

        Try
            If version >= 12 Then
                Return True
            End If
        Catch ex As Exception
        End Try

        Return False

    End Function
    Public Function SaveStdPartsAndMiscExcelReport_RAW(ByVal excelPath As String, ByRef dtExportReport As System.Data.DataTable) As Boolean
        Dim success As Boolean = False

        Try
            Dim xlApp As Microsoft.Office.Interop.Excel.Application = Nothing
            Dim xlWorkBooks As Workbooks = Nothing
            Dim xlWorkBook As Workbook = Nothing
            Dim xlSheet4 As Worksheet = Nothing
            Dim xlCells As Range = Nothing

            xlApp = New Microsoft.Office.Interop.Excel.Application With {
                .Visible = False,
                .DisplayAlerts = False
            }
            If IsXlsxVersion(xlApp.Version) Then
                excelPath += ".xlsx"
            Else
                excelPath += ".xls"
            End If
            If IO.File.Exists(excelPath) Then
                xlWorkBook = xlApp.Workbooks.Open(excelPath)
            Else
                xlWorkBooks = xlApp.Workbooks
                xlWorkBook = xlWorkBooks.Add()
            End If



            'xlWorkBook = xlWorkBooks.Add()
            Dim rowCount As Integer = 2



            Dim xlSheet3 As Worksheet = Nothing
            Dim xlCells3 As Range = Nothing
            xlSheet3 = xlWorkBook.Sheets.Add
            Dim sheetName = String.Empty
            If dtExportReport.TableName = "Misc" Then
                sheetName = "Misc"
            Else
                sheetName = "Std Parts - Hardware"
            End If
            xlSheet3.Name = sheetName
            xlCells3 = xlSheet3.Cells

            'Add columns

#Region "Add columns"

            'old add columna Name
            'Dim colCount As Integer = 1
            'For Each col As DataColumn In dtExportReport.Columns
            '    xlCells3(1, colCount) = col.ColumnName.ToString
            '    colCount = colCount + 1
            'Next

            'xlCells3(1, 1) = "Sr"
            'xlCells3(1, 2) = "Category" 'new 
            'xlCells3(1, 3) = "BEC Number"
            'xlCells3(1, 4) = "Material"
            'xlCells3(1, 5) = "Size"
            'xlCells3(1, 6) = "Description"
            'xlCells3(1, 7) = "Description2"
            'xlCells3(1, 8) = "Document Number"
            'xlCells3(1, 9) = "Quantity"
            'xlCells3(1, 10) = "Length  (inch)"
            'xlCells3(1, 11) = "Width (inch)"
            'xlCells3(1, 12) = "Stock Allowance"
            'xlCells3(1, 13) = "Total_Length (inch)"
            'xlCells3(1, 14) = "Total_Width (inch)"
            'xlCells3(1, 15) = "Area (Sq inch)"
            'xlCells3(1, 16) = "Order Area/Length"
            'xlCells3(1, 17) = "Order Area/Length(Ft)" 'new
            'xlCells3(1, 18) = "Standard Thickness/Length" 'new
            'xlCells3(1, 19) = "Title" 'new
            'xlCells3(1, 20) = "Information" 'new
            For Col = 0 To dtExportReport.Columns.Count - 1
                xlCells3(1, Col + 1) = dtExportReport.Columns(Col).ColumnName()
            Next


#End Region

            Dim i As Integer = 0
            Dim stratrows As Integer = 2
            Dim endrows As Integer = 2

            Dim startCol As String = "A"
            'Dim endCol As String = "O"
            Dim endCol As String
            If dtExportReport.TableName.Contains("Misc") Then
                endCol = "H" '"F"
            Else
                endCol = "F"
            End If


            xlSheet3.Range($"{startCol}1:{endCol}1").Font.Color = System.Drawing.Color.White
            xlSheet3.Range($"{startCol}1:{endCol}1").Interior.Color = System.Drawing.Color.Gray
            rowCount = 2

            Dim excelRangeTopRow As Range = xlCells3.Range($"{startCol}{1}", $"{endCol}{1}")
            Dim bordersTopRow As Borders = excelRangeTopRow.Borders
            bordersTopRow.LineStyle = XlLineStyle.xlContinuous
            bordersTopRow.Weight = 2D

            'Freeze top row
            Try
                Dim topRowRange As Range = xlSheet3.Range($"{startCol}2", $"{endCol}2")
                topRowRange.Select()
                xlApp.ActiveWindow.FreezePanes = True
            Catch ex As Exception
            End Try

            'Dim materialUsedList As List(Of String) = From row In dtExportReport.AsEnumerable()
            '                                          Select row.Field(Of String)(AssemblyBomForm.AssemblyColumns.Material_Used.ToString()) Distinct.ToList()
            'Dim materialUsedList As List(Of String) = From row In dtExportReport.AsEnumerable()
            '                                          Select row.Field(Of String)("BEC Material") Distinct.ToList()
            Dim abc As Boolean = True
            Dim currentColour As String = "W"

            'Set alternate colour

            For Each dr As DataRow In dtExportReport.Rows

                Dim excelRange1 As Range = xlCells3.Range($"{startCol}{rowCount}", $"G{rowCount}")
                excelRange1.Cells.NumberFormat = "@"

                For index = 1 To dtExportReport.Columns.Count
                    Dim value As String = dr(index - 1)
                    xlCells3(rowCount, index) = If(value = "", "Not Found", value)

                Next



                'Dim matUsed As String = dr("BEC Material")
                'If Not matUsed = String.Empty Then

                '    If currentColour = "W" Then
                '        currentColour = "B"
                '    Else
                '        currentColour = "W"
                '    End If

                'End If

                If currentColour = "W" Then
                    xlSheet3.Range($"{startCol}{rowCount}:{endCol}{rowCount}").Interior.Color = System.Drawing.Color.GhostWhite
                    xlSheet3.Range($"{startCol}{rowCount}:{endCol}{rowCount}").Font.Color = System.Drawing.Color.DarkBlue
                Else
                    xlSheet3.Range($"{startCol}{rowCount}:{endCol}{rowCount}").Interior.Color = System.Drawing.Color.FromArgb(252, 228, 214)
                End If

#Region "Old Code"

                'If MaterialUsed = String.Empty Then
                '    MaterialUsed = dtExportReport.Rows(i)("Material_Used").ToString().Trim()
                '    If rowCount <> 2 Then
                '        stratrows = rowCount - 1

                '    End If

                'End If

                'If MaterialUsed <> dtExportReport.Rows(i)("Material_Used").ToString().Trim() Or dtExportReport.Rows(i)("Material_Used").ToString().Trim() = String.Empty Then

                '    'endrows = rowCount - 1
                '    'Dim pl As String = $"N{stratrows}:N{endrows}"
                '    'xlCells3.Range(pl).Merge()

                '    'xlCells3.Range($"N{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                '    'xlCells3.Range($"N{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                '    'MaterialUsed = String.Empty

                'End If

                'If dtExportReport.Rows.Count = i + 1 Then

                '    'endrows = rowCount
                '    'Dim pl As String = $"N{stratrows}:N{endrows}"
                '    'xlCells3.Range(pl).Merge()
                '    'xlCells3.Range($"N{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                '    'xlCells3.Range($"N{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                '    'Dim pl1 As String = $"B{stratrows}:B{endrows}"
                '    'xlCells3.Range(pl1).Merge()
                '    ''xlSheet3.Range(pl1).MergeCells = True
                '    ''xlCells3.Range($"B{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                '    ''xlCells3.Range($"B{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                '    'Dim pl2 As String = $"C{stratrows}:C{endrows}"
                '    ''xlSheet3.Range(pl2).MergeCells = True

                '    'xlCells3.Range(pl2).Merge()
                '    ''xlCells3.Range($"C{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                '    ''xlCells3.Range($"C{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                'End If

#End Region

                Try
                    Dim excelRange As Range = xlCells3.Range($"{startCol}{rowCount}", $"{endCol}{rowCount}")
                    Dim borders As Borders = excelRange.Borders
                    borders.LineStyle = XlLineStyle.xlContinuous
                    borders.Weight = 2D
                Catch ex As Exception
                End Try

                Try
                    Dim excelRange As Range = xlCells3.Range($"{endCol}{rowCount}", $"{endCol}{rowCount}")
                    excelRange.Cells.Font.Bold = True

                    Dim excelRangeTotalW As Range = xlCells3.Range($"M{rowCount}", $"M{rowCount}")
                    excelRangeTotalW.Cells.Font.Bold = True

                    Dim excelRangeTotalL As Range = xlCells3.Range($"L{rowCount}", $"L{rowCount}")
                    excelRangeTotalL.Cells.Font.Bold = True
                Catch ex As Exception

                End Try

                i += 1
                rowCount += 1
            Next

            rowCount = 2
            stratrows = 2
            endrows = 2

            Try
                For Each dr As DataRow In dtExportReport.Rows

                    ''Dim matUsed As String = dr(AssemblyBomForm.AssemblyColumns.Material_Used.ToString())
                    'Dim matUsed As String = dr("BEC Material")
                    'If matUsed = "SMHRCH3X4.1#A588" Then
                    '    Debug.Print("aaaa")
                    'End If

                    'If Not matUsed = String.Empty And rowCount > 2 Then
                    If rowCount > 2 Then
                        endrows = rowCount

                        endrows = rowCount - 1
                        Debug.Print($"Start {stratrows} End {endrows}")

                        'Dim pl0 As String = $"P{stratrows}:P{endrows}"
                        'xlCells3.Range(pl0).Merge()
                        'xlCells3.Range($"P{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                        'xlCells3.Range($"P{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                        'Dim pl As String = $"O{stratrows}:O{endrows}"
                        'xlCells3.Range(pl).Merge()
                        'xlCells3.Range($"O{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                        'xlCells3.Range($"O{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                        'Dim pl1 As String = $"N{stratrows}:N{endrows}"
                        'xlCells3.Range(pl1).Merge()
                        'xlCells3.Range($"N{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                        'xlCells3.Range($"N{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                        Dim pl2 As String = $"B{stratrows}:B{endrows}"
                        xlCells3.Range(pl2).Merge()
                        xlCells3.Range($"B{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                        xlCells3.Range($"B{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                        Dim pl3 As String = $"C{stratrows}:C{endrows}"
                        'xlCells3.Range(pl3).Merge()
                        xlCells3.Range($"C{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                        xlCells3.Range($"C{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                        Dim pl4 As String = $"D{stratrows}:D{endrows}"
                        'xlCells3.Range(pl4).Merge()
                        xlCells3.Range($"D{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                        xlCells3.Range($"D{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                        Dim pl5 As String = $"E{stratrows}:E{endrows}"
                        'xlCells3.Range(pl5).Merge()
                        xlCells3.Range($"E{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                        xlCells3.Range($"E{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                        stratrows = rowCount
                        endrows = rowCount
                    Else
                        endrows += 1
                    End If

                    rowCount += 1
                Next

                'Dim pl6 As String = $"E{stratrows}:E{endrows}"
                'xlCells3.Range(pl6).Merge()
                'xlCells3.Range($"E{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                'xlCells3.Range($"E{endrows}").HorizontalAlignment = Excel.Constants.xlCenter
                'Dim pl1 As String = $"O{stratrows}:O{endrows}"
                'xlCells3.Range(pl1).Merge()
                'xlCells3.Range($"O{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                'xlCells3.Range($"O{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                Dim pl21 As String = $"B{stratrows}:B{endrows}"
                'xlCells3.Range(pl21).Merge()
                xlCells3.Range($"B{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                xlCells3.Range($"B{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                Dim pl31 As String = $"C{stratrows}:C{endrows}"
                'xlCells3.Range(pl31).Merge()
                xlCells3.Range($"C{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                xlCells3.Range($"C{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                Dim pl41 As String = $"D{stratrows}:D{endrows}"
                'xlCells3.Range(pl41).Merge()
                xlCells3.Range($"D{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                xlCells3.Range($"D{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                Dim pl51 As String = $"E{stratrows}:E{endrows}"
                'xlCells3.Range(pl51).Merge()
                xlCells3.Range($"E{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                xlCells3.Range($"E{endrows}").HorizontalAlignment = Excel.Constants.xlCenter
            Catch ex As Exception
            End Try

            'AutoFit
            xlSheet3.Columns.AutoFit()

            'If IO.File.Exists(excelPath) Then
            '    xlWorkBook.Save()
            'Else
            '    xlWorkBook.SaveCopyAs(excelPath)
            'End If

            'For index = 1 To dtExportReport.Rows.Count
            '    For j = 1 To dtExportReport.Columns.Count


            '        Dim value As String = xlCells3(index, j).Value.ToString
            '        If value = "NF" Then
            '            xlCells3(index, j).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
            '            If IO.File.Exists(excelPath) Then
            '                xlWorkBook.Save()
            '            Else
            '                xlWorkBook.SaveCopyAs(excelPath)
            '            End If
            '        End If
            '    Next
            'Next


            'Save and close excel

            If IO.File.Exists(excelPath) Then
                xlWorkBook.Save()
            Else
                xlWorkBook.SaveCopyAs(excelPath)
            End If

            Dim Rcount As Integer = xlSheet3.Rows.Count
            Dim Ccount As Integer = xlSheet3.Columns.Count

            xlWorkBook.Close(False)
            xlApp.UserControl = True
            xlApp.Quit()

            'Dispose objects
            If xlCells IsNot Nothing Then
                Marshal.FinalReleaseComObject(xlCells)
                xlCells = Nothing
            End If

            If xlSheet4 IsNot Nothing Then
                Marshal.FinalReleaseComObject(xlSheet4)
                xlSheet4 = Nothing
            End If

            If xlWorkBook IsNot Nothing Then
                Marshal.FinalReleaseComObject(xlWorkBook)
                xlWorkBook = Nothing
            End If

            If xlWorkBooks IsNot Nothing Then
                Marshal.FinalReleaseComObject(xlWorkBooks)
                xlWorkBooks = Nothing
            End If

            If xlApp IsNot Nothing Then
                Marshal.FinalReleaseComObject(xlApp)
                xlApp = Nothing
            End If

            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            success = True
        Catch ex As Exception
            MessageBox.Show("Unable to save excel report (Check whether Excel is installed): ", "Error")
            log.Error($"While Creating Raw Material Estimation Report (Check whether Excel is installed):{ex.Message}{ex.StackTrace}")
            CustomLogUtil.Log("While Creating Raw Material Estimation Report (Check whether Excel is installed)", ex.Message, ex.StackTrace)
        End Try
        Return success
    End Function
    Public Function SaveSPSExcelReport_RAW(ByVal excelPath As String, ByRef dtExportReport As System.Data.DataTable) As Boolean
        Dim success As Boolean = False

        Try
            Dim xlApp As Microsoft.Office.Interop.Excel.Application = Nothing
            Dim xlWorkBooks As Workbooks = Nothing
            Dim xlWorkBook As Workbook = Nothing
            Dim xlSheet4 As Worksheet = Nothing
            Dim xlCells As Range = Nothing

            xlApp = New Microsoft.Office.Interop.Excel.Application With {
                .Visible = False,
                .DisplayAlerts = False
            }
            If IsXlsxVersion(xlApp.Version) Then
                excelPath += ".xlsx"
            Else
                excelPath += ".xls"
            End If
            If IO.File.Exists(excelPath) Then
                xlWorkBook = xlApp.Workbooks.Open(excelPath)
            Else
                xlWorkBooks = xlApp.Workbooks
                xlWorkBook = xlWorkBooks.Add()
            End If



            Dim rowCount As Integer = 2



            Dim xlSheet3 As Worksheet = Nothing
            Dim xlCells3 As Range = Nothing
            xlSheet3 = xlWorkBook.Sheets.Add
            xlSheet3.Name = "Sheet-Plate-Structure" '"Export Report"
            xlCells3 = xlSheet3.Cells

            'Add columns

#Region "Add columns"

            'old add columna Name
            'Dim colCount As Integer = 1
            'For Each col As DataColumn In dtExportReport.Columns
            '    xlCells3(1, colCount) = col.ColumnName.ToString
            '    colCount = colCount + 1
            'Next

            'xlCells3(1, 1) = "Sr"
            'xlCells3(1, 2) = "Category" 'new 
            'xlCells3(1, 3) = "BEC Number"
            'xlCells3(1, 4) = "Material"
            'xlCells3(1, 5) = "Size"
            'xlCells3(1, 6) = "Description"
            'xlCells3(1, 7) = "Description2"
            'xlCells3(1, 8) = "Document Number"
            'xlCells3(1, 9) = "Quantity"
            'xlCells3(1, 10) = "Length  (inch)"
            'xlCells3(1, 11) = "Width (inch)"
            'xlCells3(1, 12) = "Stock Allowance"
            'xlCells3(1, 13) = "Total_Length (inch)"
            'xlCells3(1, 14) = "Total_Width (inch)"
            'xlCells3(1, 15) = "Area (Sq inch)"
            'xlCells3(1, 16) = "Order Area/Length"
            'xlCells3(1, 17) = "Order Area/Length(Ft)" 'new
            'xlCells3(1, 18) = "Standard Thickness/Length" 'new
            'xlCells3(1, 19) = "Title" 'new
            'xlCells3(1, 20) = "Information" 'new
            For Col = 0 To dtExportReport.Columns.Count - 1
                xlCells3(1, Col + 1) = dtExportReport.Columns(Col).ColumnName()
            Next


#End Region

            Dim i As Integer = 0
            Dim stratrows As Integer = 2
            Dim endrows As Integer = 2

            Dim startCol As String = "A"
            'Dim endCol As String = "O"
            Dim endCol As String = "P"

            xlSheet3.Range($"{startCol}1:{endCol}1").Font.Color = System.Drawing.Color.White
            xlSheet3.Range($"{startCol}1:{endCol}1").Interior.Color = System.Drawing.Color.Gray
            rowCount = 2

            Dim excelRangeTopRow As Range = xlCells3.Range($"{startCol}{1}", $"{endCol}{1}")
            Dim bordersTopRow As Borders = excelRangeTopRow.Borders
            bordersTopRow.LineStyle = XlLineStyle.xlContinuous
            bordersTopRow.Weight = 2D

            'Freeze top row
            Try
                Dim topRowRange As Range = xlSheet3.Range($"{startCol}2", $"{endCol}2")
                topRowRange.Select()
                xlApp.ActiveWindow.FreezePanes = True
            Catch ex As Exception
            End Try

            'Dim materialUsedList As List(Of String) = From row In dtExportReport.AsEnumerable()
            '                                          Select row.Field(Of String)(AssemblyBomForm.AssemblyColumns.Material_Used.ToString()) Distinct.ToList()
            Dim materialUsedList As List(Of String) = From row In dtExportReport.AsEnumerable()
                                                      Select row.Field(Of String)("BEC Material") Distinct.ToList()
            Dim abc As Boolean = True
            Dim currentColour As String = "W"

            'Set alternate colour

            For Each dr As DataRow In dtExportReport.Rows

                Dim excelRange1 As Range = xlCells3.Range($"{startCol}{rowCount}", $"G{rowCount}")
                excelRange1.Cells.NumberFormat = "@"
                Try

                    For index = 1 To dtExportReport.Columns.Count

                        Dim value As String = dr(index - 1).ToString
                        xlCells3(rowCount, index) = If(value = "", "Not Found", value)

                    Next
                Catch ex As Exception

                End Try




                Dim matUsed As String = dr("BEC Material")
                If Not matUsed = String.Empty Then

                    If currentColour = "W" Then
                        currentColour = "B"
                    Else
                        currentColour = "W"
                    End If

                End If

                If currentColour = "W" Then
                    xlSheet3.Range($"{startCol}{rowCount}:{endCol}{rowCount}").Interior.Color = System.Drawing.Color.GhostWhite
                    xlSheet3.Range($"{startCol}{rowCount}:{endCol}{rowCount}").Font.Color = System.Drawing.Color.DarkBlue
                Else
                    xlSheet3.Range($"{startCol}{rowCount}:{endCol}{rowCount}").Interior.Color = System.Drawing.Color.FromArgb(252, 228, 214)
                End If

#Region "Old Code"

                'If MaterialUsed = String.Empty Then
                '    MaterialUsed = dtExportReport.Rows(i)("Material_Used").ToString().Trim()
                '    If rowCount <> 2 Then
                '        stratrows = rowCount - 1

                '    End If

                'End If

                'If MaterialUsed <> dtExportReport.Rows(i)("Material_Used").ToString().Trim() Or dtExportReport.Rows(i)("Material_Used").ToString().Trim() = String.Empty Then

                '    'endrows = rowCount - 1
                '    'Dim pl As String = $"N{stratrows}:N{endrows}"
                '    'xlCells3.Range(pl).Merge()

                '    'xlCells3.Range($"N{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                '    'xlCells3.Range($"N{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                '    'MaterialUsed = String.Empty

                'End If

                'If dtExportReport.Rows.Count = i + 1 Then

                '    'endrows = rowCount
                '    'Dim pl As String = $"N{stratrows}:N{endrows}"
                '    'xlCells3.Range(pl).Merge()
                '    'xlCells3.Range($"N{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                '    'xlCells3.Range($"N{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                '    'Dim pl1 As String = $"B{stratrows}:B{endrows}"
                '    'xlCells3.Range(pl1).Merge()
                '    ''xlSheet3.Range(pl1).MergeCells = True
                '    ''xlCells3.Range($"B{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                '    ''xlCells3.Range($"B{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                '    'Dim pl2 As String = $"C{stratrows}:C{endrows}"
                '    ''xlSheet3.Range(pl2).MergeCells = True

                '    'xlCells3.Range(pl2).Merge()
                '    ''xlCells3.Range($"C{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                '    ''xlCells3.Range($"C{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                'End If

#End Region

                Try
                    Dim excelRange As Range = xlCells3.Range($"{startCol}{rowCount}", $"{endCol}{rowCount}")
                    Dim borders As Borders = excelRange.Borders
                    borders.LineStyle = XlLineStyle.xlContinuous
                    borders.Weight = 2D
                Catch ex As Exception
                End Try

                Try
                    Dim excelRange As Range = xlCells3.Range($"{endCol}{rowCount}", $"{endCol}{rowCount}")
                    excelRange.Cells.Font.Bold = True

                    Dim excelRangeTotalW As Range = xlCells3.Range($"M{rowCount}", $"M{rowCount}")
                    excelRangeTotalW.Cells.Font.Bold = True

                    Dim excelRangeTotalL As Range = xlCells3.Range($"L{rowCount}", $"L{rowCount}")
                    excelRangeTotalL.Cells.Font.Bold = True
                Catch ex As Exception

                End Try

                i += 1
                rowCount += 1
            Next

            rowCount = 2
            stratrows = 2
            endrows = 2

            Try
                For Each dr As DataRow In dtExportReport.Rows

                    'Dim matUsed As String = dr(AssemblyBomForm.AssemblyColumns.Material_Used.ToString())
                    Dim matUsed As String = dr("BEC Material")
                    If matUsed = "SMHRCH3X4.1#A588" Then
                        Debug.Print("aaaa")
                    End If

                    If Not matUsed = String.Empty And rowCount > 2 Then

                        endrows = rowCount

                        endrows = rowCount - 1
                        Debug.Print($"Start {stratrows} End {endrows}")

                        Dim pl0 As String = $"P{stratrows}:P{endrows}"
                        xlCells3.Range(pl0).Merge()
                        xlCells3.Range($"P{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                        xlCells3.Range($"P{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                        Dim pl As String = $"O{stratrows}:O{endrows}"
                        xlCells3.Range(pl).Merge()
                        xlCells3.Range($"O{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                        xlCells3.Range($"O{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                        'Dim pl1 As String = $"N{stratrows}:N{endrows}"
                        'xlCells3.Range(pl1).Merge()
                        'xlCells3.Range($"N{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                        'xlCells3.Range($"N{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                        Dim pl2 As String = $"B{stratrows}:B{endrows}"
                        xlCells3.Range(pl2).Merge()
                        xlCells3.Range($"B{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                        xlCells3.Range($"B{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                        Dim pl3 As String = $"C{stratrows}:C{endrows}"
                        xlCells3.Range(pl3).Merge()
                        xlCells3.Range($"C{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                        xlCells3.Range($"C{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                        Dim pl4 As String = $"D{stratrows}:D{endrows}"
                        xlCells3.Range(pl4).Merge()
                        xlCells3.Range($"D{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                        xlCells3.Range($"D{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                        Dim pl5 As String = $"E{stratrows}:E{endrows}"
                        xlCells3.Range(pl5).Merge()
                        xlCells3.Range($"E{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                        xlCells3.Range($"E{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                        stratrows = rowCount
                        endrows = rowCount
                    Else
                        endrows += 1
                    End If

                    rowCount += 1
                Next

                'Dim pl6 As String = $"E{stratrows}:E{endrows}"
                'xlCells3.Range(pl6).Merge()
                'xlCells3.Range($"E{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                'xlCells3.Range($"E{endrows}").HorizontalAlignment = Excel.Constants.xlCenter
                Dim pl01 As String = $"P{stratrows}:P{endrows}"
                xlCells3.Range(pl01).Merge()
                xlCells3.Range($"O{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                xlCells3.Range($"O{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                Dim pl1 As String = $"O{stratrows}:O{endrows}"
                xlCells3.Range(pl1).Merge()
                xlCells3.Range($"O{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                xlCells3.Range($"O{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                Dim pl21 As String = $"B{stratrows}:B{endrows}"
                xlCells3.Range(pl21).Merge()
                xlCells3.Range($"B{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                xlCells3.Range($"B{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                Dim pl31 As String = $"C{stratrows}:C{endrows}"
                xlCells3.Range(pl31).Merge()
                xlCells3.Range($"C{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                xlCells3.Range($"C{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                Dim pl41 As String = $"D{stratrows}:D{endrows}"
                xlCells3.Range(pl41).Merge()
                xlCells3.Range($"D{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                xlCells3.Range($"D{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                Dim pl51 As String = $"E{stratrows}:E{endrows}"
                xlCells3.Range(pl51).Merge()
                xlCells3.Range($"E{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                xlCells3.Range($"E{endrows}").HorizontalAlignment = Excel.Constants.xlCenter
            Catch ex As Exception
            End Try

            'AutoFit
            xlSheet3.Columns.AutoFit()
            'If IO.File.Exists(excelPath) Then
            '    xlWorkBook.Save()
            'Else
            '    xlWorkBook.SaveCopyAs(excelPath)
            'End If
            'For index = 1 To dtExportReport.Rows.Count
            '    For j = 0 To dtExportReport.Columns.Count


            '        Dim value As String = xlCells3(i, index).Value.ToString
            '        If value = "NF" Then
            '            xlSheet3.Range(xlCells3(i, index)).Interior.Color = System.Drawing.Color.Red
            '            If IO.File.Exists(excelPath) Then
            '                xlWorkBook.Save()
            '            Else
            '                xlWorkBook.SaveCopyAs(excelPath)
            '            End If
            '        End If
            '    Next
            'Next


            'Save and close excel
            If IO.File.Exists(excelPath) Then
                xlWorkBook.Save()
            Else
                xlWorkBook.SaveCopyAs(excelPath)
            End If
            xlWorkBook.Close(False)
            xlApp.UserControl = True
            xlApp.Quit()

            'Dispose objects
            If xlCells IsNot Nothing Then
                Marshal.FinalReleaseComObject(xlCells)
                xlCells = Nothing
            End If

            If xlSheet4 IsNot Nothing Then
                Marshal.FinalReleaseComObject(xlSheet4)
                xlSheet4 = Nothing
            End If

            If xlWorkBook IsNot Nothing Then
                Marshal.FinalReleaseComObject(xlWorkBook)
                xlWorkBook = Nothing
            End If

            If xlWorkBooks IsNot Nothing Then
                Marshal.FinalReleaseComObject(xlWorkBooks)
                xlWorkBooks = Nothing
            End If

            If xlApp IsNot Nothing Then
                Marshal.FinalReleaseComObject(xlApp)
                xlApp = Nothing
            End If

            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            success = True
        Catch ex As Exception
            MessageBox.Show("Unable to save excel report (Check whether Excel is installed): ", "Error")
            log.Error($"While Creating Raw Material Estimation Report (Check whether Excel is installed):{ex.Message}{ex.StackTrace}")
            CustomLogUtil.Log("While Creating Raw Material Estimation Report (Check whether Excel is installed)", ex.Message, ex.StackTrace)
        End Try
        Return success
    End Function
    Public Function SaveExcelReport(ByVal excelPath As String, ByRef dtExportReport As System.Data.DataTable) As Boolean
        Dim success As Boolean = False
        dtExportReport.Columns.RemoveAt(1)
        dtExportReport.Columns.RemoveAt(2)
        Try
            Dim xlApp As Microsoft.Office.Interop.Excel.Application = Nothing
            Dim xlWorkBooks As Workbooks = Nothing
            Dim xlWorkBook As Workbook = Nothing
            Dim xlSheet4 As Worksheet = Nothing
            Dim xlCells As Range = Nothing

            xlApp = New Microsoft.Office.Interop.Excel.Application With {
                .Visible = False,
                .DisplayAlerts = False
            }

            xlWorkBooks = xlApp.Workbooks
            xlWorkBook = xlWorkBooks.Add()
            Dim rowCount As Integer = 2

            If IsXlsxVersion(xlApp.Version) Then
                excelPath += ".xlsx"
            Else
                excelPath += ".xls"
            End If

            Dim xlSheet3 As Worksheet = Nothing
            Dim xlCells3 As Range = Nothing
            xlSheet3 = xlWorkBook.Sheets.Add
            xlSheet3.Name = "Export Report"
            xlCells3 = xlSheet3.Cells

            'Add columns

#Region "Add columns"

            'old add columna Name
            'Dim colCount As Integer = 1
            'For Each col As DataColumn In dtExportReport.Columns
            '    xlCells3(1, colCount) = col.ColumnName.ToString
            '    colCount = colCount + 1
            'Next

            'xlCells3(1, 1) = "Sr"
            'xlCells3(1, 2) = "Category" 'new 
            'xlCells3(1, 3) = "BEC Number"
            'xlCells3(1, 4) = "Material"
            'xlCells3(1, 5) = "Size"
            'xlCells3(1, 6) = "Description"
            'xlCells3(1, 7) = "Description2"
            'xlCells3(1, 8) = "Document Number"
            'xlCells3(1, 9) = "Quantity"
            'xlCells3(1, 10) = "Length  (inch)"
            'xlCells3(1, 11) = "Width (inch)"
            'xlCells3(1, 12) = "Stock Allowance"
            'xlCells3(1, 13) = "Total_Length (inch)"
            'xlCells3(1, 14) = "Total_Width (inch)"
            'xlCells3(1, 15) = "Area (Sq inch)"
            'xlCells3(1, 16) = "Order Area/Length"
            'xlCells3(1, 17) = "Order Area/Length(Ft)" 'new
            'xlCells3(1, 18) = "Standard Thickness/Length" 'new
            'xlCells3(1, 19) = "Title" 'new
            'xlCells3(1, 20) = "Information" 'new

            xlCells3(1, 1) = "Sr"
            xlCells3(1, 2) = "Category" 'new 
            xlCells3(1, 3) = "BEC Number"
            xlCells3(1, 4) = "Material"
            'xlCells3(1, 5) = "Size"
            xlCells3(1, 6) = "Description"
            'xlCells3(1, 7) = "Description2"
            xlCells3(1, 8) = "Part Number" '"Document Number"
            xlCells3(1, 9) = "Quantity"
            xlCells3(1, 10) = "Length  (inch)"
            xlCells3(1, 11) = "Width (inch)"
            xlCells3(1, 12) = "Stock Allowance"
            xlCells3(1, 13) = "Total_Length (inch)"
            xlCells3(1, 14) = "Total_Width (inch)"
            xlCells3(1, 15) = "Area (Sq inch)"
            xlCells3(1, 16) = "Order Area/Length"
            xlCells3(1, 17) = "Order Area/Length(Ft)" 'new
            xlCells3(1, 18) = "Standard Thickness/Length" 'new
            xlCells3(1, 19) = "Title" 'new
            xlCells3(1, 20) = "Information" 'new
#End Region

            Dim i As Integer = 0
            Dim stratrows As Integer = 2
            Dim endrows As Integer = 2

            Dim startCol As String = "A"
            Dim endCol As String = "O"

            xlSheet3.Range($"{startCol}1:{endCol}1").Font.Color = System.Drawing.Color.White
            xlSheet3.Range($"{startCol}1:{endCol}1").Interior.Color = System.Drawing.Color.Gray
            rowCount = 2

            Dim excelRangeTopRow As Range = xlCells3.Range($"{startCol}{1}", $"{endCol}{1}")
            Dim bordersTopRow As Borders = excelRangeTopRow.Borders
            bordersTopRow.LineStyle = XlLineStyle.xlContinuous
            bordersTopRow.Weight = 2D

            'Freeze top row
            Try
                Dim topRowRange As Range = xlSheet3.Range($"{startCol}2", $"{endCol}2")
                topRowRange.Select()
                xlApp.ActiveWindow.FreezePanes = True
            Catch ex As Exception
            End Try

            Dim materialUsedList As List(Of String) = From row In dtExportReport.AsEnumerable()
                                                      Select row.Field(Of String)(AssemblyBomForm.AssemblyColumns.Material_Used.ToString()) Distinct.ToList()

            Dim abc As Boolean = True
            Dim currentColour As String = "W"

            'Set alternate colour

            For Each dr As DataRow In dtExportReport.Rows

                Dim excelRange1 As Range = xlCells3.Range($"{startCol}{rowCount}", $"G{rowCount}")
                excelRange1.Cells.NumberFormat = "@"

                For index = 1 To dtExportReport.Columns.Count
                    xlCells3(rowCount, index) = dr(index - 1)
                Next

                Dim matUsed As String = dr(AssemblyBomForm.AssemblyColumns.Material_Used.ToString())
                If Not matUsed = String.Empty Then

                    If currentColour = "W" Then
                        currentColour = "B"
                    Else
                        currentColour = "W"
                    End If

                End If

                If currentColour = "W" Then
                    xlSheet3.Range($"{startCol}{rowCount}:{endCol}{rowCount}").Interior.Color = System.Drawing.Color.GhostWhite
                    xlSheet3.Range($"{startCol}{rowCount}:{endCol}{rowCount}").Font.Color = System.Drawing.Color.DarkBlue
                Else
                    xlSheet3.Range($"{startCol}{rowCount}:{endCol}{rowCount}").Interior.Color = System.Drawing.Color.FromArgb(252, 228, 214)
                End If

#Region "Old Code"

                'If MaterialUsed = String.Empty Then
                '    MaterialUsed = dtExportReport.Rows(i)("Material_Used").ToString().Trim()
                '    If rowCount <> 2 Then
                '        stratrows = rowCount - 1

                '    End If

                'End If

                'If MaterialUsed <> dtExportReport.Rows(i)("Material_Used").ToString().Trim() Or dtExportReport.Rows(i)("Material_Used").ToString().Trim() = String.Empty Then

                '    'endrows = rowCount - 1
                '    'Dim pl As String = $"N{stratrows}:N{endrows}"
                '    'xlCells3.Range(pl).Merge()

                '    'xlCells3.Range($"N{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                '    'xlCells3.Range($"N{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                '    'MaterialUsed = String.Empty

                'End If

                'If dtExportReport.Rows.Count = i + 1 Then

                '    'endrows = rowCount
                '    'Dim pl As String = $"N{stratrows}:N{endrows}"
                '    'xlCells3.Range(pl).Merge()
                '    'xlCells3.Range($"N{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                '    'xlCells3.Range($"N{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                '    'Dim pl1 As String = $"B{stratrows}:B{endrows}"
                '    'xlCells3.Range(pl1).Merge()
                '    ''xlSheet3.Range(pl1).MergeCells = True
                '    ''xlCells3.Range($"B{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                '    ''xlCells3.Range($"B{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                '    'Dim pl2 As String = $"C{stratrows}:C{endrows}"
                '    ''xlSheet3.Range(pl2).MergeCells = True

                '    'xlCells3.Range(pl2).Merge()
                '    ''xlCells3.Range($"C{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                '    ''xlCells3.Range($"C{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                'End If

#End Region

                Try
                    Dim excelRange As Range = xlCells3.Range($"{startCol}{rowCount}", $"{endCol}{rowCount}")
                    Dim borders As Borders = excelRange.Borders
                    borders.LineStyle = XlLineStyle.xlContinuous
                    borders.Weight = 2D
                Catch ex As Exception
                End Try

                Try
                    Dim excelRange As Range = xlCells3.Range($"{endCol}{rowCount}", $"{endCol}{rowCount}")
                    excelRange.Cells.Font.Bold = True

                    Dim excelRangeTotalW As Range = xlCells3.Range($"M{rowCount}", $"M{rowCount}")
                    excelRangeTotalW.Cells.Font.Bold = True

                    Dim excelRangeTotalL As Range = xlCells3.Range($"L{rowCount}", $"L{rowCount}")
                    excelRangeTotalL.Cells.Font.Bold = True
                Catch ex As Exception

                End Try

                i += 1
                rowCount += 1
            Next

            rowCount = 2
            stratrows = 2
            endrows = 2

            Try
                For Each dr As DataRow In dtExportReport.Rows

                    Dim matUsed As String = dr(AssemblyBomForm.AssemblyColumns.Material_Used.ToString())

                    If matUsed = "SMHRCH3X4.1#A588" Then
                        Debug.Print("aaaa")
                    End If

                    If Not matUsed = String.Empty And rowCount > 2 Then

                        endrows = rowCount

                        endrows = rowCount - 1
                        Debug.Print($"Start {stratrows} End {endrows}")

                        Dim pl As String = $"O{stratrows}:O{endrows}"
                        xlCells3.Range(pl).Merge()
                        xlCells3.Range($"O{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                        xlCells3.Range($"O{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                        'Dim pl1 As String = $"N{stratrows}:N{endrows}"
                        'xlCells3.Range(pl1).Merge()
                        'xlCells3.Range($"N{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                        'xlCells3.Range($"N{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                        Dim pl2 As String = $"B{stratrows}:B{endrows}"
                        xlCells3.Range(pl2).Merge()
                        xlCells3.Range($"B{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                        xlCells3.Range($"B{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                        Dim pl3 As String = $"C{stratrows}:C{endrows}"
                        xlCells3.Range(pl3).Merge()
                        xlCells3.Range($"C{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                        xlCells3.Range($"C{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                        Dim pl4 As String = $"D{stratrows}:D{endrows}"
                        xlCells3.Range(pl4).Merge()
                        xlCells3.Range($"D{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                        xlCells3.Range($"D{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                        Dim pl5 As String = $"E{stratrows}:E{endrows}"
                        xlCells3.Range(pl5).Merge()
                        xlCells3.Range($"E{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                        xlCells3.Range($"E{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                        stratrows = rowCount
                        endrows = rowCount
                    Else
                        endrows += 1
                    End If

                    rowCount += 1
                Next

                'Dim pl6 As String = $"E{stratrows}:E{endrows}"
                'xlCells3.Range(pl6).Merge()
                'xlCells3.Range($"E{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                'xlCells3.Range($"E{endrows}").HorizontalAlignment = Excel.Constants.xlCenter
                Dim pl1 As String = $"O{stratrows}:O{endrows}"
                xlCells3.Range(pl1).Merge()
                xlCells3.Range($"O{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                xlCells3.Range($"O{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                Dim pl21 As String = $"B{stratrows}:B{endrows}"
                xlCells3.Range(pl21).Merge()
                xlCells3.Range($"B{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                xlCells3.Range($"B{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                Dim pl31 As String = $"C{stratrows}:C{endrows}"
                xlCells3.Range(pl31).Merge()
                xlCells3.Range($"C{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                xlCells3.Range($"C{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                Dim pl41 As String = $"D{stratrows}:D{endrows}"
                xlCells3.Range(pl41).Merge()
                xlCells3.Range($"D{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                xlCells3.Range($"D{endrows}").HorizontalAlignment = Excel.Constants.xlCenter

                Dim pl51 As String = $"E{stratrows}:E{endrows}"
                xlCells3.Range(pl51).Merge()
                xlCells3.Range($"E{endrows}").VerticalAlignment = Excel.Constants.xlCenter
                xlCells3.Range($"E{endrows}").HorizontalAlignment = Excel.Constants.xlCenter
            Catch ex As Exception
            End Try

            'AutoFit
            xlSheet3.Columns.AutoFit()

            'Save and close excel
            xlWorkBook.SaveCopyAs(excelPath)
            xlWorkBook.Close(False)
            xlApp.UserControl = True
            xlApp.Quit()

            'Dispose objects
            If xlCells IsNot Nothing Then
                Marshal.FinalReleaseComObject(xlCells)
                xlCells = Nothing
            End If

            If xlSheet4 IsNot Nothing Then
                Marshal.FinalReleaseComObject(xlSheet4)
                xlSheet4 = Nothing
            End If

            If xlWorkBook IsNot Nothing Then
                Marshal.FinalReleaseComObject(xlWorkBook)
                xlWorkBook = Nothing
            End If

            If xlWorkBooks IsNot Nothing Then
                Marshal.FinalReleaseComObject(xlWorkBooks)
                xlWorkBooks = Nothing
            End If

            If xlApp IsNot Nothing Then
                Marshal.FinalReleaseComObject(xlApp)
                xlApp = Nothing
            End If

            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            success = True
        Catch ex As Exception
            MessageBox.Show("Unable to save excel report (Check whether Excel is installed): ", "Error")
            log.Error($"While Creating Raw Material Estimation Report (Check whether Excel is installed):{ex.Message}{ex.StackTrace}")
            CustomLogUtil.Log("While Creating Raw Material Estimation Report (Check whether Excel is installed)", ex.Message, ex.StackTrace)
        End Try
        Return success
    End Function

    Private Function GetSeparateSheetColumns() As List(Of String)
        Dim separateSheetReportList As New List(Of String) From {
            "Document name",
            "Document type",
            "IVC type",
            "Material",
            "Family",
            "Material description",
            "Is material mismatched",
            "Is separate sheet exist",
            "Is DXF sheet exists",
            "DXF sheets name",
            "Document path",
            "Separate sheet name",
            "View names"
        }
        Return separateSheetReportList
    End Function

End Class