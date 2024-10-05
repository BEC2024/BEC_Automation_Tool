Imports Microsoft.Office.Interop.Excel

Public Class MTC_MTR_ExcelUtil

    Dim excelcol As Integer = 3
    Dim columnChar As Char = "Z"
    Dim color1 As System.Drawing.Color = Color.FromArgb(226, 239, 218)
    Dim color2 As System.Drawing.Color = Color.FromArgb(252, 228, 214)
    Dim colorRed As System.Drawing.Color = Color.FromArgb(255, 199, 206)


    'temp 20Feb2024
    ''' <summary>
    '''
    ''' </summary>
    ''' <param name="mtcMtrModelObj"></param>
    ''' 
    Public Sub SaveAsMTCExcel(ByVal mtcMtrModelObj As MTC_MTR_Model, MTC_Author_Name As String)
        '9th Sep 2024 'Added try..catch
        Try
            Dim excelAsmCol As Integer = 4
            Dim excelPartCol As Integer = 4
            Dim excelSheetMetalCol As Integer = 4
            Dim excelBaseLineCol As Integer = 4
            Dim excelElectricalCol As Integer = 4

            Dim asmExcelRowCnt As Integer = 23
            Dim partExcelRowCnt As Integer = 22
            Dim sheetMetalExcelRowCnt As Integer = 24
            Dim baseLineExcelRowCnt As Integer = 42
            Dim electricalExcelRowCnt As Integer = 11

            Dim xlApp As Application
            Dim xlWorkBook As Workbook

            Dim xlWorkSheetAssembly As Worksheet
            Dim xlWorkSheetPart As Worksheet
            Dim xlWorkSheetSheetMetal As Worksheet
            Dim xlWorkSheetBaseline As Worksheet
            Dim xlWorkSheetElectrical As Worksheet

            xlApp = New Application
            'Dim aPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location
            'Dim dirPath As String = IO.Path.GetDirectoryName(aPath)
            'Dim mtcExcelPath As String = IO.Path.Combine(dirPath, $"MTC_BEC.xlsx")
            Dim mtcExcelPath As String = Config.configObj.MTCExcelPath

            xlWorkBook = xlApp.Workbooks.Open(mtcExcelPath)

            xlWorkSheetAssembly = xlWorkBook.Worksheets("Assembly")
            xlWorkSheetPart = xlWorkBook.Worksheets("Part")
            xlWorkSheetSheetMetal = xlWorkBook.Worksheets("Sheetmetal")
            xlWorkSheetBaseline = xlWorkBook.Worksheets("Baseline")
            xlWorkSheetElectrical = xlWorkBook.Worksheets("Electrical")

            '2nd Sep 2024
            'SetHorizontalAlignment(xlWorkSheetAssembly)
            'SetHorizontalAlignment(xlWorkSheetPart)
            'SetHorizontalAlignment(xlWorkSheetSheetMetal)
            'SetHorizontalAlignment(xlWorkSheetBaseline)
            'SetHorizontalAlignment(xlWorkSheetElectrical)

            '2nd Sep 2024
            ' Set horizontal and vertical alignment to center for all worksheets
            CenterTextInWorksheet(xlWorkSheetAssembly)
            CenterTextInWorksheet(xlWorkSheetPart)
            CenterTextInWorksheet(xlWorkSheetSheetMetal)
            CenterTextInWorksheet(xlWorkSheetBaseline)
            CenterTextInWorksheet(xlWorkSheetElectrical)

            SetAssemblyWorkSheet_MTC(xlWorkSheetAssembly, mtcMtrModelObj.mtcAssemblyList_BEC, mtcMtrModelObj.mtcAssemblyList_DGS, MTC_Author_Name)
            SetPartWorkSheet_MTC(xlWorkSheetPart, mtcMtrModelObj.mtcPartList_BEC, mtcMtrModelObj.mtcPartList_DGS, MTC_Author_Name)
            SetSheetMetalWorkSheet_MTC(xlWorkSheetSheetMetal, mtcMtrModelObj.mtcSheetMetalList_BEC, mtcMtrModelObj.mtcSheetMetalList_DGS, MTC_Author_Name)
            SetBaseLineWorkSheet_MTC(xlWorkSheetBaseline, mtcMtrModelObj.mtcBaseLineList_BEC, mtcMtrModelObj.mtcBaseLineList_DGS, MTC_Author_Name)
            SetElectricalPartWorkSheet_MTC(xlWorkSheetElectrical, mtcMtrModelObj.mtcElectricalPartList_BEC, mtcMtrModelObj.mtcElectricalPartList_DGS, MTC_Author_Name)

            xlWorkSheetAssembly.Columns.AutoFit()
            xlWorkSheetPart.Columns.AutoFit()
            xlWorkSheetSheetMetal.Columns.AutoFit()
            xlWorkSheetBaseline.Columns.AutoFit()
            xlWorkSheetElectrical.Columns.AutoFit()

            '2nd Sep 2024
            ' Set "Assembly" sheet as the default sheet that opens
            xlWorkSheetAssembly.Activate()

            Dim Asmname As String = IO.Path.GetFileNameWithoutExtension(mtcMtrModelObj.assemblyPath)
            Dim fileName As String = $"{Asmname}_MTC_Report_{System.DateTime.Now.ToString("MMM_d_yyyy_HH_mm")}"
            Dim newname As String = IO.Path.Combine(mtcMtrModelObj.export_MTR_Report_DirectoryLocation, $"{fileName}.xlsx")
            xlWorkBook.SaveAs(newname)
            xlWorkBook.Close()
            xlApp.Quit()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub



    '17th Sep 2024

    ''' <summary>
    '''
    ''' </summary>
    ''' <param name="mtcMtrModelObj"></param>
    ''' 
    Public Sub SaveAsMTCExcelPartDoc(ByVal mtcMtrModelObj As MTC_MTR_Model, MTC_Author_Name As String)

        Try

            Dim excelPartCol As Integer = 4

            Dim partExcelRowCnt As Integer = 22

            Dim xlApp As Application
            Dim xlWorkBook As Workbook

            Dim xlWorkSheetPart As Worksheet

            xlApp = New Application

            Dim mtcExcelPath As String = Config.configObj.MTCExcelPath

            xlWorkBook = xlApp.Workbooks.Open(mtcExcelPath)

            xlWorkSheetPart = xlWorkBook.Worksheets("Part")

            CenterTextInWorksheet(xlWorkSheetPart)

            SetPartWorkSheet_MTC(xlWorkSheetPart, mtcMtrModelObj.mtcPartList_BEC, mtcMtrModelObj.mtcPartList_DGS, MTC_Author_Name)

            xlWorkSheetPart.Columns.AutoFit()

            'xlApp.DisplayAlerts = False

            'xlWorkBook.Worksheets("Assembly").Delete()
            'xlWorkBook.Worksheets("Sheetmetal").Delete()
            'xlWorkBook.Worksheets("Baseline").Delete()
            'xlWorkBook.Worksheets("Electrical").Delete()

            xlWorkSheetPart.Activate()

            'xlApp.DisplayAlerts = True

            Dim PartName As String = IO.Path.GetFileNameWithoutExtension(mtcMtrModelObj.partPath)
            Dim fileName As String = $"{PartName}_MTC_Report_{System.DateTime.Now.ToString("MMM_d_yyyy_HH_mm")}"
            Dim newname As String = IO.Path.Combine(mtcMtrModelObj.exportDirectoryLocation, $"{fileName}.xlsx")
            xlWorkBook.SaveAs(newname)
            xlWorkBook.Close()
            xlApp.Quit()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub



    '17th Sep 2024
    ''' <summary>
    '''
    ''' </summary>
    ''' <param name="mtcMtrModelObj"></param>
    ''' 
    Public Sub SaveAsMTCExcelSheetMetalDoc(ByVal mtcMtrModelObj As MTC_MTR_Model, MTC_Author_Name As String)

        Try
            Dim excelSheetMetalCol As Integer = 4

            Dim sheetMetalExcelRowCnt As Integer = 24

            Dim xlApp As Application
            Dim xlWorkBook As Workbook

            Dim xlWorkSheetSheetMetal As Worksheet

            xlApp = New Application

            Dim mtcExcelPath As String = Config.configObj.MTCExcelPath

            xlWorkBook = xlApp.Workbooks.Open(mtcExcelPath)

            xlWorkSheetSheetMetal = xlWorkBook.Worksheets("Sheetmetal")

            CenterTextInWorksheet(xlWorkSheetSheetMetal)

            SetSheetMetalWorkSheet_MTC(xlWorkSheetSheetMetal, mtcMtrModelObj.mtcSheetMetalList_BEC, mtcMtrModelObj.mtcSheetMetalList_DGS, MTC_Author_Name)

            xlWorkSheetSheetMetal.Columns.AutoFit()

            'xlApp.DisplayAlerts = False

            'xlWorkBook.Worksheets("Assembly").Delete()
            'xlWorkBook.Worksheets("Part").Delete()
            'xlWorkBook.Worksheets("Baseline").Delete()
            'xlWorkBook.Worksheets("Electrical").Delete()

            xlWorkSheetSheetMetal.Activate()

            'xlApp.DisplayAlerts = True

            Dim SheetMetalname As String = IO.Path.GetFileNameWithoutExtension(mtcMtrModelObj.sheetMetalPath)
            Dim fileName As String = $"{SheetMetalname}_MTC_Report_{System.DateTime.Now.ToString("MMM_d_yyyy_HH_mm")}"
            Dim newname As String = IO.Path.Combine(mtcMtrModelObj.exportDirectoryLocation, $"{fileName}.xlsx")
            xlWorkBook.SaveAs(newname)
            xlWorkBook.Close()
            xlApp.Quit()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    '2nd Sep 2024
    Private Sub CenterTextInWorksheet(ByVal worksheet As Worksheet)
        worksheet.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter
        worksheet.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter
    End Sub

    Private Sub SetSheetMetalWorkSheet_MTC(ByRef xlWorkSheet As Worksheet, ByVal finalPartListBEC As List(Of MTC_SheetMetal), ByVal finalPartListDGS As List(Of MTC_SheetMetal), MTC_Author_Name As String)

        '4th Sep 2024
        'Author Info and time span
        'xlWorkSheet.Cells(1, 1).Interior.Color = System.Drawing.Color.Gray
        'xlWorkSheet.Cells(1, 1).font.Color = System.Drawing.Color.White
        xlWorkSheet.Cells(1, 1).font.Bold = True
        'xlWorkSheet.Cells(1, 1) = $"{MTC_Author_Name}_{System.DateTime.Now.ToString("MMM_d_yyyy_HH_mm")}"
        xlWorkSheet.Cells(1, 3) = $"{MTC_Author_Name} {System.DateTime.Now.ToString("(MMM_d_yyyy_HH_mm)")}"

        '4th Sep 2024
        'Dim cnt As Integer = 1

        Dim cnt As Integer = 2

        Dim excelcol As Integer = 4

        For Each mTCReviewObj As MTC_SheetMetal In finalPartListBEC

            '4th Sep 2024
            ''0. Assembly Name
            'xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
            'xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
            'xlWorkSheet.Cells(1, excelcol).font.Bold = True
            'xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.assemblyName

            '0. Assembly Name
            xlWorkSheet.Cells(cnt, excelcol).Interior.Color = System.Drawing.Color.FromArgb(165, 165, 165)  '9th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(cnt, excelcol).font.Bold = True
            xlWorkSheet.Cells(cnt, excelcol) = mTCReviewObj.assemblyName

            '9th Sep 2024
            Dim borders1 As Borders = xlWorkSheet.Cells(cnt, excelcol).Borders
            borders1.LineStyle = XlLineStyle.xlContinuous
            borders1.Weight = 2D

            '4th Sep 2024
            'SetExcelColumnColorAndBorder(xlWorkSheet, 24, excelcol)
            SetExcelColumnColorAndBorderNewMTC(xlWorkSheet, 24, excelcol)

            '4th Sep 2024
            'If mTCReviewObj.isValidPart = False Then
            '    Dim excelRange As Range = xlWorkSheet.Cells(1, excelcol)
            '    excelRange.Interior.Color = Color.Red
            '    xlWorkSheet.Cells(1, excelcol).AddComment($"Invalid sheetmetal part.")
            'End If


            '26th September 2024

            'If mTCReviewObj.isValidPart = False Then
            '    Dim excelRange As Range = xlWorkSheet.Cells(cnt, excelcol)
            '    excelRange.Interior.Color = Color.Red
            '    xlWorkSheet.Cells(cnt, excelcol).AddComment($"Invalid sheetmetal part.")
            'End If

            If mTCReviewObj.isValidPart = False Then
                Dim excelRange As Range = xlWorkSheet.Cells(cnt, excelcol)
                'excelRange.Interior.Color = Color.Red
                xlWorkSheet.Cells(cnt, excelcol).AddComment("Invalid path.")
            End If


            'If Not mTCReviewObj.isValidPart = False Then
            '    Dim excelRange As Range = xlWorkSheet.Cells(cnt, excelcol)
            '    'excelRange.Interior.Color = Color.Red
            '    xlWorkSheet.Cells(cnt, excelcol).AddComment("Invalid part.")
            'End If

            '1. Eco Number

            xlWorkSheet.Cells(cnt + 1, excelcol) = mTCReviewObj.ecoNumber
            If mTCReviewObj.revisionLevel = "0" Then
                xlWorkSheet.Cells(cnt + 1, excelcol).AddComment($"SOW = {mTCReviewObj.ecoNumber}")
            Else
                xlWorkSheet.Cells(cnt + 1, excelcol).AddComment($"ECO = {mTCReviewObj.ecoNumber}")
            End If

            '2. Part Number
            xlWorkSheet.Cells(cnt + 2, excelcol) = mTCReviewObj.partNumber
            xlWorkSheet.Cells(cnt + 2, excelcol).AddComment("Part Number:" + mTCReviewObj.assemblyName)
            If (mTCReviewObj.assemblyName = "No") Then
                xlWorkSheet.Cells(cnt + 2, excelcol).Interior.Color = colorRed
            End If

            '3. Revision Level
            xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.revisionLevel

            '4. Author
            xlWorkSheet.Cells(cnt + 4, excelcol) = mTCReviewObj.author

            '5. Project Name

            xlWorkSheet.Cells(cnt + 5, excelcol) = mTCReviewObj.projectNameExist
            xlWorkSheet.Cells(cnt + 5, excelcol).AddComment("ProjectName:" + mTCReviewObj.projectName.Replace("/", ""))
            If (mTCReviewObj.projectName = "No") Then
                xlWorkSheet.Cells(cnt + 5, excelcol).Interior.Color = colorRed
            End If

            ''6. Revision Number Correct
            'xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.revisionNumberCorrect
            'xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            'If (mTCReviewObj.revisionNumberCorrect = "No") Then
            '    xlWorkSheet.Cells(cnt + 6, excelcol).Interior.Color = colorRed
            'End If

            '13th Sep 2024
            '6. Revision Number Correct
            xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.revisionNumberCorrect
            xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            If (mTCReviewObj.revisionNumberCorrect = "No") Then
                xlWorkSheet.Cells(cnt + 6, excelcol).Interior.Color = colorRed
            End If

            '7. Document Number correct
            xlWorkSheet.Cells(cnt + 7, excelcol) = mTCReviewObj.documentNumberCorrect
            xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Document Number:" + mTCReviewObj.documentNumber)
            If (mTCReviewObj.documentNumber = "No") Then
                xlWorkSheet.Cells(cnt + 7, excelcol).Interior.Color = colorRed
            End If

            '8. Is the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name)
            xlWorkSheet.Cells(cnt + 8, excelcol) = mTCReviewObj.authorExists
            xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Author:" + mTCReviewObj.author)
            If (mTCReviewObj.authorExists = "No") Then
                xlWorkSheet.Cells(cnt + 8, excelcol).Interior.Color = colorRed
            End If

            '9. Do all technically unused properties have a "dash" populated?
            xlWorkSheet.Cells(cnt + 9, excelcol) = mTCReviewObj.isDashPopulated
            If (mTCReviewObj.isDashPopulated = "No") Then
                xlWorkSheet.Cells(cnt + 9, excelcol).Interior.Color = colorRed
            End If

            '10. Does the Model Title MATCH the M2M Item Master (INV) Description Field?
            xlWorkSheet.Cells(cnt + 10, excelcol) = mTCReviewObj.isTitleMatch_ItemMaster
            xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.title)
            If (mTCReviewObj.isTitleMatch_ItemMaster = "No") Then
                xlWorkSheet.Cells(cnt + 10, excelcol).Interior.Color = colorRed
            End If


            ''11. Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
            'xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.isUOMMatch_M2M
            ''26th September 2024
            'If Not mTCReviewObj.UomProperty = "" Then
            '    xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
            'End If
            'If (mTCReviewObj.isUOMMatch_M2M = "No") Then
            '    xlWorkSheet.Cells(cnt + 11, excelcol).Interior.Color = colorRed
            'End If

            '11. Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
            If mTCReviewObj.isUOMMatch_M2M = "Yes" Then
                '26th September 2024
                If Not mTCReviewObj.UomProperty = "" Then
                    xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.isUOMMatch_M2M
                    xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
                End If
            Else
                '26th September 2024
                If Not mTCReviewObj.UomProperty = "" Then
                    xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.isUOMMatch_M2M
                    xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
                    xlWorkSheet.Cells(cnt + 11, excelcol).Interior.Color = colorRed
                End If
            End If


            '12. Is the mat'l spec field populated? (indicated by RED in metadata audit file)
            xlWorkSheet.Cells(cnt + 12, excelcol) = mTCReviewObj.isMaterialSpecExists
            xlWorkSheet.Cells(cnt + 12, excelcol).AddComment("Material Specification:" + mTCReviewObj.materialSpec)
            If (mTCReviewObj.isMaterialSpecExists = "No") Then
                xlWorkSheet.Cells(cnt + 12, excelcol).Interior.Color = colorRed
            End If

            '13. Is the material used field populated? *
            xlWorkSheet.Cells(cnt + 13, excelcol) = mTCReviewObj.isMaterialUsedExists
            xlWorkSheet.Cells(cnt + 13, excelcol).AddComment("Material Used:" + mTCReviewObj.materialUsed)
            If (mTCReviewObj.isMaterialUsedExists = "No") Then
                xlWorkSheet.Cells(cnt + 13, excelcol).Interior.Color = colorRed
            End If

            '14. Is the bend radius of the part equal to or above the ASTM minimum? *
            xlWorkSheet.Cells(cnt + 14, excelcol) = mTCReviewObj.gageExcelFile
            If (mTCReviewObj.gageExcelFile = "No") Then
                xlWorkSheet.Cells(cnt + 14, excelcol).Interior.Color = colorRed
            End If

            '15. . Is the flat pattern turned ON? (Applicable only to parts needing dxf or cut length) *
            xlWorkSheet.Cells(cnt + 15, excelcol) = mTCReviewObj.isFlatPatternActive
            If (mTCReviewObj.isFlatPatternActive = "No") Then
                xlWorkSheet.Cells(cnt + 15, excelcol).Interior.Color = colorRed
            End If

            '16. Did I use the hole tool for ALL holes requiring fasteners? (include clearance holes for        hardware, tapped holes, And Slots) *
            xlWorkSheet.Cells(cnt + 16, excelcol) = mTCReviewObj.holeToolsUsed
            If (mTCReviewObj.holeToolsUsed = "No") Then
                xlWorkSheet.Cells(cnt + 16, excelcol).Interior.Color = colorRed
            End If

            '17. Is the part "Adjustable"? (part should NOT be adjustable) *
            xlWorkSheet.Cells(cnt + 17, excelcol) = mTCReviewObj.isAdjustatble
            If (mTCReviewObj.isAdjustatble = "Yes") Then
                xlWorkSheet.Cells(cnt + 17, excelcol).Interior.Color = colorRed
            End If

            '18. Does the M2M Source reflect part creation (Make child of stk/pur parent, Project        procurement strategy, material types Like composite, etc.)? *
            'xlWorkSheet.Cells(cnt + 18, excelcol) = mTCReviewObj.m2mSource

            If mTCReviewObj.m2mSource = "" Then
                xlWorkSheet.Cells(cnt + 18, excelcol) = "No"
                xlWorkSheet.Cells(cnt + 18, excelcol).Interior.Color = colorRed
            Else
                xlWorkSheet.Cells(cnt + 18, excelcol) = "Yes"

                xlWorkSheet.Cells(cnt + 18, excelcol).AddComment(mTCReviewObj.m2mSource)

            End If

            '19. What is the modified date?
            xlWorkSheet.Cells(cnt + 19, excelcol) = mTCReviewObj.modifiedDate

            excelcol = excelcol + 1

            '5th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).Interior.Color = System.Drawing.Color.FromArgb(165, 165, 165)  '9th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(cnt, excelcol).font.Bold = True
            xlWorkSheet.Cells(cnt, excelcol) = "Issue Addressed?"

            '9th Sep 2024
            Dim borders As Borders = xlWorkSheet.Cells(cnt, excelcol).Borders
            borders.LineStyle = XlLineStyle.xlContinuous
            borders.Weight = 2D

            AddDropDownListToExcelColumn(xlWorkSheet, 24, excelcol)

            excelcol = excelcol + 1
        Next

        '----------------------------------------------------------------
        'temp 20Feb2024


        For Each mTCReviewObj As MTC_SheetMetal In finalPartListDGS

            '4th Sep 2024
            ''0. Assembly Name
            'xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
            'xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
            'xlWorkSheet.Cells(1, excelcol).font.Bold = True
            'xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.assemblyName

            '0. Assembly Name
            xlWorkSheet.Cells(cnt, excelcol).Interior.Color = System.Drawing.Color.FromArgb(165, 165, 165)  '9th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(cnt, excelcol).font.Bold = True
            xlWorkSheet.Cells(cnt, excelcol) = mTCReviewObj.assemblyName

            '9th Sep 2024
            Dim borders1 As Borders = xlWorkSheet.Cells(cnt, excelcol).Borders
            borders1.LineStyle = XlLineStyle.xlContinuous
            borders1.Weight = 2D

            '4th Sep 2024
            'SetExcelColumnColorAndBorder(xlWorkSheet, 24, excelcol)
            SetExcelColumnColorAndBorderNewMTC(xlWorkSheet, 24, excelcol)

            '4th Sep 2024
            'If mTCReviewObj.isValidPart = False Then
            '    Dim excelRange As Range = xlWorkSheet.Cells(1, excelcol)
            '    excelRange.Interior.Color = Color.Red
            '    xlWorkSheet.Cells(1, excelcol).AddComment($"Invalid sheetmetal part.")
            'End If


            '26th September 2024

            'If mTCReviewObj.isValidPart = False Then
            '    Dim excelRange As Range = xlWorkSheet.Cells(cnt, excelcol)
            '    excelRange.Interior.Color = Color.Red
            '    xlWorkSheet.Cells(cnt, excelcol).AddComment($"Invalid sheetmetal part.")
            'End If

            If mTCReviewObj.isValidPart = False Then
                Dim excelRange As Range = xlWorkSheet.Cells(cnt, excelcol)
                'excelRange.Interior.Color = Color.Red
                xlWorkSheet.Cells(cnt, excelcol).AddComment("Invalid path.")
            End If

            '1. Eco Number

            xlWorkSheet.Cells(cnt + 1, excelcol) = mTCReviewObj.ecoNumber
            If mTCReviewObj.revisionLevel = "0" Then
                xlWorkSheet.Cells(cnt + 1, excelcol).AddComment($"SOW = {mTCReviewObj.ecoNumber}")
            Else
                xlWorkSheet.Cells(cnt + 1, excelcol).AddComment($"ECO = {mTCReviewObj.ecoNumber}")
            End If

            '2. Part Number
            xlWorkSheet.Cells(cnt + 2, excelcol) = mTCReviewObj.partNumber
            xlWorkSheet.Cells(cnt + 2, excelcol).AddComment("Part Number:" + mTCReviewObj.assemblyName)
            If (mTCReviewObj.assemblyName = "No") Then
                xlWorkSheet.Cells(cnt + 2, excelcol).Interior.Color = colorRed
            End If

            '3. Revision Level
            xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.revisionLevel

            '4. Author
            xlWorkSheet.Cells(cnt + 4, excelcol) = mTCReviewObj.author

            '5. Project Name

            xlWorkSheet.Cells(cnt + 5, excelcol) = mTCReviewObj.projectNameExist
            xlWorkSheet.Cells(cnt + 5, excelcol).AddComment("ProjectName:" + mTCReviewObj.projectName.Replace("/", ""))
            If (mTCReviewObj.projectName = "No") Then
                xlWorkSheet.Cells(cnt + 5, excelcol).Interior.Color = colorRed
            End If

            ''6. Revision Number Correct
            'xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.revisionNumberCorrect
            'xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            'If (mTCReviewObj.revisionNumberCorrect = "No") Then
            '    xlWorkSheet.Cells(cnt + 6, excelcol).Interior.Color = colorRed
            'End If

            '13th Sep 2024
            '6. Revision Number Correct
            xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.revisionNumberCorrect
            xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            If (mTCReviewObj.revisionNumberCorrect = "No") Then
                xlWorkSheet.Cells(cnt + 6, excelcol).Interior.Color = colorRed
            End If

            '7. Document Number correct
            xlWorkSheet.Cells(cnt + 7, excelcol) = mTCReviewObj.documentNumberCorrect
            xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Document Number:" + mTCReviewObj.documentNumber)
            If (mTCReviewObj.documentNumber = "No") Then
                xlWorkSheet.Cells(cnt + 7, excelcol).Interior.Color = colorRed
            End If

            '8. Is the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name)
            xlWorkSheet.Cells(cnt + 8, excelcol) = mTCReviewObj.authorExists
            xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Author:" + mTCReviewObj.author)
            If (mTCReviewObj.authorExists = "No") Then
                xlWorkSheet.Cells(cnt + 8, excelcol).Interior.Color = colorRed
            End If

            '9. Do all technically unused properties have a "dash" populated?
            xlWorkSheet.Cells(cnt + 9, excelcol) = mTCReviewObj.isDashPopulated
            If (mTCReviewObj.isDashPopulated = "No") Then
                xlWorkSheet.Cells(cnt + 9, excelcol).Interior.Color = colorRed
            End If

            '10. Does the Model Title MATCH the M2M Item Master (INV) Description Field?
            xlWorkSheet.Cells(cnt + 10, excelcol) = mTCReviewObj.isTitleMatch_ItemMaster
            xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.title)
            If (mTCReviewObj.isTitleMatch_ItemMaster = "No") Then
                xlWorkSheet.Cells(cnt + 10, excelcol).Interior.Color = colorRed
            End If


            ''11. Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
            'xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.isUOMMatch_M2M
            'xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
            'If (mTCReviewObj.isUOMMatch_M2M = "No") Then
            '    xlWorkSheet.Cells(cnt + 11, excelcol).Interior.Color = colorRed
            'End If


            ''11. Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
            'xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.isUOMMatch_M2M
            ''26th September 2024
            'If Not mTCReviewObj.UomProperty = "" Then
            '    xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
            'End If
            'If (mTCReviewObj.isUOMMatch_M2M = "No") Then
            '    xlWorkSheet.Cells(cnt + 11, excelcol).Interior.Color = colorRed
            'End If


            '11. Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
            If mTCReviewObj.isUOMMatch_M2M = "Yes" Then
                '26th September 2024
                If Not mTCReviewObj.UomProperty = "" Then
                    xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.isUOMMatch_M2M
                    xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
                End If
            Else
                '26th September 2024
                If Not mTCReviewObj.UomProperty = "" Then
                    xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.isUOMMatch_M2M
                    xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
                    xlWorkSheet.Cells(cnt + 11, excelcol).Interior.Color = colorRed
                End If
            End If


            '12. Is the mat'l spec field populated? (indicated by RED in metadata audit file)
            xlWorkSheet.Cells(cnt + 12, excelcol) = mTCReviewObj.isMaterialSpecExists
            xlWorkSheet.Cells(cnt + 12, excelcol).AddComment("Material Specification:" + mTCReviewObj.materialSpec)
            If (mTCReviewObj.isMaterialSpecExists = "No") Then
                xlWorkSheet.Cells(cnt + 12, excelcol).Interior.Color = colorRed
            End If

            '13. Is the material used field populated? *
            xlWorkSheet.Cells(cnt + 13, excelcol) = mTCReviewObj.isMaterialUsedExists
            xlWorkSheet.Cells(cnt + 13, excelcol).AddComment("Material Used:" + mTCReviewObj.materialUsed)
            If (mTCReviewObj.isMaterialUsedExists = "No") Then
                xlWorkSheet.Cells(cnt + 13, excelcol).Interior.Color = colorRed
            End If

            '14. Is the bend radius of the part equal to or above the ASTM minimum? *
            xlWorkSheet.Cells(cnt + 14, excelcol) = mTCReviewObj.gageExcelFile
            If (mTCReviewObj.gageExcelFile = "No") Then
                xlWorkSheet.Cells(cnt + 14, excelcol).Interior.Color = colorRed
            End If

            '15. . Is the flat pattern turned ON? (Applicable only to parts needing dxf or cut length) *
            xlWorkSheet.Cells(cnt + 15, excelcol) = mTCReviewObj.isFlatPatternActive
            If (mTCReviewObj.isFlatPatternActive = "No") Then
                xlWorkSheet.Cells(cnt + 15, excelcol).Interior.Color = colorRed
            End If

            '16. Did I use the hole tool for ALL holes requiring fasteners? (include clearance holes for        hardware, tapped holes, And Slots) *
            xlWorkSheet.Cells(cnt + 16, excelcol) = mTCReviewObj.holeToolsUsed
            If (mTCReviewObj.holeToolsUsed = "No") Then
                xlWorkSheet.Cells(cnt + 16, excelcol).Interior.Color = colorRed
            End If

            '17. Is the part "Adjustable"? (part should NOT be adjustable) *
            xlWorkSheet.Cells(cnt + 17, excelcol) = mTCReviewObj.isAdjustatble
            If (mTCReviewObj.isAdjustatble = "Yes") Then
                xlWorkSheet.Cells(cnt + 17, excelcol).Interior.Color = colorRed
            End If

            '18. Does the M2M Source reflect part creation (Make child of stk/pur parent, Project        procurement strategy, material types Like composite, etc.)? *
            'xlWorkSheet.Cells(cnt + 18, excelcol) = mTCReviewObj.m2mSource

            If mTCReviewObj.m2mSource = "" Then
                xlWorkSheet.Cells(cnt + 18, excelcol) = "No"
                xlWorkSheet.Cells(cnt + 18, excelcol).Interior.Color = colorRed
            Else
                xlWorkSheet.Cells(cnt + 18, excelcol) = "Yes"

                xlWorkSheet.Cells(cnt + 18, excelcol).AddComment(mTCReviewObj.m2mSource)

            End If

            '19. What is the modified date?
            xlWorkSheet.Cells(cnt + 19, excelcol) = mTCReviewObj.modifiedDate

            excelcol = excelcol + 1

            '5th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).Interior.Color = System.Drawing.Color.FromArgb(165, 165, 165)  '9th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(cnt, excelcol).font.Bold = True
            xlWorkSheet.Cells(cnt, excelcol) = "Issue Addressed?"

            '9th Sep 2024
            Dim borders As Borders = xlWorkSheet.Cells(cnt, excelcol).Borders
            borders.LineStyle = XlLineStyle.xlContinuous
            borders.Weight = 2D

            AddDropDownListToExcelColumn(xlWorkSheet, 24, excelcol)

            excelcol = excelcol + 1
        Next

    End Sub

    Private Sub SetBaseLineWorkSheet_MTC(ByRef xlWorkSheet As Worksheet, ByVal finalPartListBEC As List(Of MTC_BaseLine), ByVal finalPartListDGS As List(Of MTC_BaseLine), MTC_Author_Name As String)
        '4th Sep 2024
        'Author Info and time span
        'xlWorkSheet.Cells(1, 1).Interior.Color = System.Drawing.Color.Gray
        'xlWorkSheet.Cells(1, 1).font.Color = System.Drawing.Color.White
        xlWorkSheet.Cells(1, 1).font.Bold = True
        'xlWorkSheet.Cells(1, 1) = $"{MTC_Author_Name}_{System.DateTime.Now.ToString("MMM_d_yyyy_HH_mm")}"
        xlWorkSheet.Cells(1, 3) = $"{MTC_Author_Name} {System.DateTime.Now.ToString("(MMM_d_yyyy_HH_mm)")}"

        '4th Sep 2024
        'Dim cnt As Integer = 1
        Dim cnt As Integer = 2

        Dim excelcol As Integer = 4

        For Each mTCReviewObj As MTC_BaseLine In finalPartListBEC

            '4th Sep 2024
            ''0. Assembly Name
            'xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
            'xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
            'xlWorkSheet.Cells(1, excelcol).font.Bold = True
            'xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.assemblyName

            '0. Assembly Name
            xlWorkSheet.Cells(cnt, excelcol).Interior.Color = System.Drawing.Color.FromArgb(165, 165, 165)  '9th Sep 2024   
            xlWorkSheet.Cells(cnt, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(cnt, excelcol).font.Bold = True
            xlWorkSheet.Cells(cnt, excelcol) = mTCReviewObj.assemblyName

            '9th Sep 2024
            Dim borders1 As Borders = xlWorkSheet.Cells(cnt, excelcol).Borders
            borders1.LineStyle = XlLineStyle.xlContinuous
            borders1.Weight = 2D

            '4th Sep 2024
            'SetExcelColumnColorAndBorder(xlWorkSheet, 42, excelcol)
            SetExcelColumnColorAndBorderNewMTC(xlWorkSheet, 42, excelcol)

            '4th Sep 2024
            'If Not mTCReviewObj.isValidBaseLineDirectoryPath Then
            '    Dim excelRange As Range = xlWorkSheet.Cells(1, excelcol)
            '    excelRange.Interior.Color = Color.Red
            '    xlWorkSheet.Cells(1, excelcol).AddComment("Comments:" + "Invalid baseline directory Path.")
            'End If



            '26th September 2024

            'If Not mTCReviewObj.isValidBaseLineDirectoryPath Then
            '    Dim excelRange As Range = xlWorkSheet.Cells(cnt, excelcol)
            '    excelRange.Interior.Color = Color.Red
            '    xlWorkSheet.Cells(cnt, excelcol).AddComment("Comments:" + "Invalid baseline directory Path.")
            'End If


            If Not mTCReviewObj.isValidBaseLineDirectoryPath Then
                Dim excelRange As Range = xlWorkSheet.Cells(cnt, excelcol)
                'excelRange.Interior.Color = Color.Red
                xlWorkSheet.Cells(cnt, excelcol).AddComment("Invalid path.")
            End If


#Region "Old"
            ''1. Is the part number match with M2M? *
            'xlWorkSheet.Cells(cnt + 1, excelcol).AddComment("PartName:" + mTCReviewObj.assemblyName)

            ''2. What is the revision level? *
            'xlWorkSheet.Cells(cnt + 2, excelcol) = mTCReviewObj.revisionLevel

            ''3. Who is the Author of the file? (For BEC authors, type 3 capitalized initials of author)
            'xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.author

            ''4. What type of component?
            'xlWorkSheet.Cells(cnt + 4, excelcol) = mTCReviewObj.category

            ''5. Virtual thread applied for Fasteners?
            'xlWorkSheet.Cells(cnt + 5, excelcol) = mTCReviewObj.isVirtualThreadExists

            ''6. Are ALL features fully defined in the model Pathfinder? * (Sketch fully constraint?)
            'xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.isSketchFullyDefined

            ''7. Any suppressed feature found?
            'xlWorkSheet.Cells(cnt + 7, excelcol) = mTCReviewObj.isSuppressFeatureFound

            ''8. Is the "Mat'l Spec" field populated? (PURCHASED for library components) *
            'xlWorkSheet.Cells(cnt + 8, excelcol) = mTCReviewObj.isMaterialSpecExists
            'xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Material Specification:" + mTCReviewObj.materialSpec)

            ''9. Is the "Material Used" field populated? (PURCHASED for library components) *
            'xlWorkSheet.Cells(cnt + 9, excelcol) = mTCReviewObj.isMaterialUsedExists
            'xlWorkSheet.Cells(cnt + 9, excelcol).AddComment("Material Used:" + mTCReviewObj.materialUsed)

            ''10. Is the "Title" field populated with the correct part description? (This should MATCH the M2M Item Master (INV) Description field) *
            'xlWorkSheet.Cells(cnt + 10, excelcol) = mTCReviewObj.isMaterialDesc_Title

            'xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.title)

            ''11. Is the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) *
            'xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.isAuthorExists
            'xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("Author:" + mTCReviewObj.author)

            ''12. Is the "Keywords" field populated with the FULL M2M Item Master Description as shown in the Comments field? *
            'xlWorkSheet.Cells(cnt + 12, excelcol) = String.Empty

            ''13. Is the "Comments" field populated with the Vendor name and Vendor part number? (It should appear as VENDOR NAME = VENDOR PART NUMBER) *
            'xlWorkSheet.Cells(cnt + 13, excelcol) = mTCReviewObj.isCommentExist
            'xlWorkSheet.Cells(cnt + 13, excelcol).AddComment("Comments:" + mTCReviewObj.comment)

            ''14. Is the "Document Number" field populated with the correct part number? (This should  MATCH the M2M Item Master Part Number field) *
            'xlWorkSheet.Cells(cnt + 14, excelcol) = mTCReviewObj.isCorrectDocumentNumber
            'xlWorkSheet.Cells(cnt + 14, excelcol).AddComment("Document Number:" + mTCReviewObj.documentNumber)

            ''15. Is the "Revision" field populated with the correct revision number? *
            'xlWorkSheet.Cells(cnt + 15, excelcol) = mTCReviewObj.isCorrectRevisionNumber
            'xlWorkSheet.Cells(cnt + 15, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)

            ''16. "Is the ""Project"" field populated with the correct Project Name? (The Project should be BROOKVILLE EQUIPMENT CORP) *"
            'xlWorkSheet.Cells(cnt + 16, excelcol) = mTCReviewObj.isCorrectProjectName

            ''17. Is the "Hardware Part" box checked for hardware components? (This should be for ALL  nuts, bolts, washers, screws, etc.) *
            'xlWorkSheet.Cells(cnt + 17, excelcol) = mTCReviewObj.isHardwarePartBoxChecked
            'xlWorkSheet.Cells(cnt + 17, excelcol).AddComment("Status & HardwarePart:" + mTCReviewObj.documentType + "&" + mTCReviewObj.hardwarePart.ToString())

            ''18. Do all other unused property fields have a "dash" (-) populated? *
            'xlWorkSheet.Cells(cnt + 18, excelcol) = mTCReviewObj.isDashPopulated

            ''19. Does the unit of measures (UOM) match M2M "For Inventory UOM"? (i.e. EA, sqft, in, etc.) *
            'xlWorkSheet.Cells(cnt + 19, excelcol) = mTCReviewObj.isUOMMatch_M2M
            'xlWorkSheet.Cells(cnt + 19, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)

            ''20. Is the M2M Source marked stock/purchased? *
            'xlWorkSheet.Cells(cnt + 20, excelcol) = mTCReviewObj.isM2MSourceStocked

            ''21. Did I use the hole tool for ALL holes requiring fasteners and they match the Vendor supplied Data?(This include clearances holes For hardware, tapped holes, And Slots) *
            'xlWorkSheet.Cells(cnt + 21, excelcol) = mTCReviewObj.isHoleToolUsed

            ''22. Are all inter-part copies, part copies and included geometry broken? *
            'xlWorkSheet.Cells(cnt + 22, excelcol) = String.Empty

            ''23. inter-part copies detected?
            'xlWorkSheet.Cells(cnt + 23, excelcol) = mTCReviewObj.isInterPartCopiesDetected

            ''24. part copies detected?
            'xlWorkSheet.Cells(cnt + 24, excelcol) = mTCReviewObj.isPartCopiesDetected

            ''25. broken  file Path detected?
            'xlWorkSheet.Cells(cnt + 25, excelcol) = mTCReviewObj.isBrokenFilePathDetected

            ''26. Is the model a SE Adjustable Part? (Models should NOT be adjustable) *
            'xlWorkSheet.Cells(cnt + 26, excelcol) = mTCReviewObj.isAdjustable

            ''27. Does this model have any SE Simplified Features? (NO model should contain SE Simplified Features) *
            'xlWorkSheet.Cells(cnt + 27, excelcol) = mTCReviewObj.hasSESimplifiedFeature

            ''28. Do all hardware AND library components in this asm model have an SE Status of Baselined?(If Available Or Released status hardware/library component exists, select "NO" ) *
            'xlWorkSheet.Cells(cnt + 28, excelcol) = mTCReviewObj.hasSEStatusBaseLined

            ''29. What is the last modified date
            'xlWorkSheet.Cells(cnt + 29, excelcol) = mTCReviewObj.modifiedDate
#End Region

            '26th September 2024

            ''1. Is the part number match with M2M? *
            'xlWorkSheet.Cells(cnt + 1, excelcol).AddComment("Part Number::" + mTCReviewObj.assemblyName)
            'If (mTCReviewObj.assemblyName = "No") Then
            '    xlWorkSheet.Cells(cnt + 1, excelcol).Interior.Color = colorRed
            'End If

            '1. Is the part number match with M2M? *   
            xlWorkSheet.Cells(cnt + 1, excelcol) = mTCReviewObj.isPartFound
            xlWorkSheet.Cells(cnt + 1, excelcol).AddComment("Part Number:" + mTCReviewObj.assemblyName)
            If (mTCReviewObj.assemblyName = "No") Then
                xlWorkSheet.Cells(cnt + 1, excelcol).Interior.Color = colorRed
            End If

            '2. What is the revision level? *
            xlWorkSheet.Cells(cnt + 2, excelcol) = mTCReviewObj.revisionLevel

            '3. Who is the Author of the file? (For BEC authors, type 3 capitalized initials of author)
            xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.author
            If (mTCReviewObj.author = "No") Then
                xlWorkSheet.Cells(cnt + 3, excelcol).Interior.Color = colorRed
            End If

            '4. What type of component?
            xlWorkSheet.Cells(cnt + 4, excelcol) = mTCReviewObj.category

            '5. Virtual thread applied for Fasteners?
            xlWorkSheet.Cells(cnt + 5, excelcol) = mTCReviewObj.isVirtualThreadExists
            If (mTCReviewObj.isVirtualThreadExists = "No") Then
                xlWorkSheet.Cells(cnt + 5, excelcol).Interior.Color = colorRed
            End If

            '6. Are ALL features fully defined in the model Pathfinder? * (Sketch fully constraint?)
            xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.isSketchFullyDefined
            If (mTCReviewObj.isSketchFullyDefined = "No") Then
                xlWorkSheet.Cells(cnt + 6, excelcol).Interior.Color = colorRed
            End If

            '7. Any suppressed feature found?
            xlWorkSheet.Cells(cnt + 7, excelcol) = mTCReviewObj.isSuppressFeatureFound
            If (mTCReviewObj.isSuppressFeatureFound = "Yes") Then
                xlWorkSheet.Cells(cnt + 7, excelcol).Interior.Color = colorRed
            End If

            '8. Is the "Mat'l Spec" field populated? (PURCHASED for library components) *
            xlWorkSheet.Cells(cnt + 8, excelcol) = mTCReviewObj.isMaterialSpecExists
            xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Material Specification:" + mTCReviewObj.materialSpec)
            If (mTCReviewObj.isMaterialSpecExists = "No") Then
                xlWorkSheet.Cells(cnt + 8, excelcol).Interior.Color = colorRed
            End If

            '9. Is the "Material Used" field populated? (PURCHASED for library components) *
            xlWorkSheet.Cells(cnt + 9, excelcol) = mTCReviewObj.isMaterialUsedExists
            xlWorkSheet.Cells(cnt + 9, excelcol).AddComment("Material Used:" + mTCReviewObj.materialUsed)
            If (mTCReviewObj.isMaterialUsedExists = "No") Then
                xlWorkSheet.Cells(cnt + 9, excelcol).Interior.Color = colorRed
            End If

            '10. Is the "Title" field populated with the correct part description? (This should MATCH the M2M Item Master (INV) Description field) *
            xlWorkSheet.Cells(cnt + 10, excelcol) = mTCReviewObj.isMaterialDesc_Title

            xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.title)
            If (mTCReviewObj.title = "No") Then
                xlWorkSheet.Cells(cnt + 10, excelcol).Interior.Color = colorRed
            End If

            '11. Is the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) *
            xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.isAuthorExists
            xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("Author:" + mTCReviewObj.author)
            If (mTCReviewObj.author = "No") Then
                xlWorkSheet.Cells(cnt + 11, excelcol).Interior.Color = colorRed
            End If

            '12. Is the "Document Number" field populated with the correct part number? (This should  MATCH the M2M Item Master Part Number field) *
            xlWorkSheet.Cells(cnt + 12, excelcol) = mTCReviewObj.isCorrectDocumentNumber
            xlWorkSheet.Cells(cnt + 12, excelcol).AddComment("Document Number:" + mTCReviewObj.documentNumber)
            If (mTCReviewObj.documentNumber = "No") Then
                xlWorkSheet.Cells(cnt + 12, excelcol).Interior.Color = colorRed
            End If

            ''13. Is the "Revision" field populated with the correct revision number? *
            'xlWorkSheet.Cells(cnt + 13, excelcol) = mTCReviewObj.isCorrectRevisionNumber
            'xlWorkSheet.Cells(cnt + 13, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            'If (mTCReviewObj.revisionLevel = "No") Then
            '    xlWorkSheet.Cells(cnt + 13, excelcol).Interior.Color = colorRed
            'End If

            '13th Sep 2024
            '13. Is the "Revision" field populated with the correct revision number? *
            xlWorkSheet.Cells(cnt + 13, excelcol) = mTCReviewObj.isCorrectRevisionNumber
            xlWorkSheet.Cells(cnt + 13, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            If (mTCReviewObj.isCorrectRevisionNumber = "No") Then
                xlWorkSheet.Cells(cnt + 6, excelcol).Interior.Color = colorRed
            End If


            '14. "Is the ""Project"" field populated with the correct Project Name? (The Project should be BROOKVILLE EQUIPMENT CORP) *"
            xlWorkSheet.Cells(cnt + 14, excelcol) = mTCReviewObj.isCorrectProjectName
            If (mTCReviewObj.isCorrectProjectName = "No") Then
                xlWorkSheet.Cells(cnt + 14, excelcol).Interior.Color = colorRed
            End If

            '15. Is the "Hardware Part" box checked for hardware components? (This should be for ALL  nuts, bolts, washers, screws, etc.) *
            xlWorkSheet.Cells(cnt + 15, excelcol) = mTCReviewObj.isHardwarePartBoxChecked
            xlWorkSheet.Cells(cnt + 15, excelcol).AddComment("Status & HardwarePart:" + mTCReviewObj.documentType + "&" + mTCReviewObj.hardwarePart.ToString())
            If (mTCReviewObj.isHardwarePartBoxChecked = "No") Then
                xlWorkSheet.Cells(cnt + 15, excelcol).Interior.Color = colorRed
            End If

            '16. Do all other unused property fields have a "dash" (-) populated? *
            xlWorkSheet.Cells(cnt + 16, excelcol) = mTCReviewObj.isDashPopulated
            If (mTCReviewObj.isDashPopulated = "No") Then
                xlWorkSheet.Cells(cnt + 16, excelcol).Interior.Color = colorRed
            End If

            ''17. Does the unit of measures (UOM) match M2M "For Inventory UOM"? (i.e. EA, sqft, in, etc.) *
            'xlWorkSheet.Cells(cnt + 17, excelcol) = mTCReviewObj.isUOMMatch_M2M
            'xlWorkSheet.Cells(cnt + 17, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
            'If (mTCReviewObj.UomProperty = "No") Then
            '    xlWorkSheet.Cells(cnt + 17, excelcol).Interior.Color = colorRed
            'End If


            '17. Does the unit of measures (UOM) match M2M "For Inventory UOM"? (i.e. EA, sqft, in, etc.) *
            If mTCReviewObj.isUOMMatch_M2M = "Yes" Then
                '26th September 2024
                If Not mTCReviewObj.UomProperty = "" Then
                    xlWorkSheet.Cells(cnt + 17, excelcol) = mTCReviewObj.isUOMMatch_M2M
                    xlWorkSheet.Cells(cnt + 17, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
                End If
            Else
                '26th September 2024
                If Not mTCReviewObj.UomProperty = "" Then
                    xlWorkSheet.Cells(cnt + 17, excelcol) = mTCReviewObj.isUOMMatch_M2M
                    xlWorkSheet.Cells(cnt + 17, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
                    xlWorkSheet.Cells(cnt + 17, excelcol).Interior.Color = colorRed
                End If
            End If


            '18. Is the M2M Source marked stock/purchased? *
            xlWorkSheet.Cells(cnt + 18, excelcol) = mTCReviewObj.isM2MSourceStocked
            If (mTCReviewObj.isM2MSourceStocked = "No") Then
                xlWorkSheet.Cells(cnt + 18, excelcol).Interior.Color = colorRed
            End If

            '19. Did I use the hole tool for ALL holes requiring fasteners and they match the Vendor supplied Data?(This include clearances holes For hardware, tapped holes, And Slots) *
            xlWorkSheet.Cells(cnt + 19, excelcol) = mTCReviewObj.isHoleToolUsed
            If (mTCReviewObj.isHoleToolUsed = "No") Then
                xlWorkSheet.Cells(cnt + 19, excelcol).Interior.Color = colorRed
            End If

            '20. inter-part copies detected?
            xlWorkSheet.Cells(cnt + 20, excelcol) = mTCReviewObj.isInterPartCopiesDetected
            If (mTCReviewObj.isInterPartCopiesDetected = "Yes") Then
                xlWorkSheet.Cells(cnt + 20, excelcol).Interior.Color = colorRed
            End If

            '21. part copies detected?
            xlWorkSheet.Cells(cnt + 21, excelcol) = mTCReviewObj.isPartCopiesDetected
            If (mTCReviewObj.isPartCopiesDetected = "Yes") Then
                xlWorkSheet.Cells(cnt + 21, excelcol).Interior.Color = colorRed
            End If

            '22. broken  file Path detected?
            xlWorkSheet.Cells(cnt + 22, excelcol) = mTCReviewObj.isBrokenFilePathDetected
            If (mTCReviewObj.isBrokenFilePathDetected = "Yes") Then
                xlWorkSheet.Cells(cnt + 22, excelcol).Interior.Color = colorRed
            End If

            '23. Is the model a SE Adjustable Part? (Models should NOT be adjustable) *
            xlWorkSheet.Cells(cnt + 23, excelcol) = mTCReviewObj.isAdjustable
            If (mTCReviewObj.isAdjustable = "Yes") Then
                xlWorkSheet.Cells(cnt + 23, excelcol).Interior.Color = colorRed
            End If

            '24. Does this model have any SE Simplified Features? (NO model should contain SE Simplified Features) *
            xlWorkSheet.Cells(cnt + 24, excelcol) = mTCReviewObj.hasSESimplifiedFeature
            If (mTCReviewObj.hasSESimplifiedFeature = "Yes") Then
                xlWorkSheet.Cells(cnt + 24, excelcol).Interior.Color = colorRed
            End If

            '25. Do all hardware AND library components in this asm model have an SE Status of Baselined?(If Available Or Released status hardware/library component exists, select "NO" ) *
            xlWorkSheet.Cells(cnt + 25, excelcol) = mTCReviewObj.hasSEStatusBaseLined
            If (mTCReviewObj.hasSEStatusBaseLined = "No") Then
                xlWorkSheet.Cells(cnt + 25, excelcol).Interior.Color = colorRed
            End If

            '26. What is the last modified date
            xlWorkSheet.Cells(cnt + 26, excelcol) = mTCReviewObj.modifiedDate

            excelcol = excelcol + 1

            '5th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).Interior.Color = System.Drawing.Color.FromArgb(165, 165, 165)  '9th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(cnt, excelcol).font.Bold = True
            xlWorkSheet.Cells(cnt, excelcol) = "Issue Addressed?"

            '9th Sep 2024
            Dim borders As Borders = xlWorkSheet.Cells(cnt, excelcol).Borders
            borders.LineStyle = XlLineStyle.xlContinuous
            borders.Weight = 2D

            AddDropDownListToExcelColumn(xlWorkSheet, 42, excelcol)

            excelcol = excelcol + 1
        Next

        '---------------------------------------------
        'temp 20Feb2024

        For Each mTCReviewObj As MTC_BaseLine In finalPartListDGS

            '4th Sep 2024
            ''0. Assembly Name
            'xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
            'xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
            'xlWorkSheet.Cells(1, excelcol).font.Bold = True
            'xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.assemblyName

            '0. Assembly Name
            xlWorkSheet.Cells(cnt, excelcol).Interior.Color = System.Drawing.Color.FromArgb(165, 165, 165)  '9th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(cnt, excelcol).font.Bold = True
            xlWorkSheet.Cells(cnt, excelcol) = mTCReviewObj.assemblyName

            '9th Sep 2024
            Dim borders1 As Borders = xlWorkSheet.Cells(cnt, excelcol).Borders
            borders1.LineStyle = XlLineStyle.xlContinuous
            borders1.Weight = 2D

            '4th Sep 2024
            'SetExcelColumnColorAndBorder(xlWorkSheet, 42, excelcol)
            SetExcelColumnColorAndBorderNewMTC(xlWorkSheet, 42, excelcol)

            '4th Sep 2024
            'If Not mTCReviewObj.isValidBaseLineDirectoryPath Then
            '    Dim excelRange As Range = xlWorkSheet.Cells(1, excelcol)
            '    excelRange.Interior.Color = Color.Red
            '    xlWorkSheet.Cells(1, excelcol).AddComment("Comments:" + "Invalid baseline directory Path.")
            'End If



            '26th September 2024

            'If Not mTCReviewObj.isValidBaseLineDirectoryPath Then
            '    Dim excelRange As Range = xlWorkSheet.Cells(cnt, excelcol)
            '    excelRange.Interior.Color = Color.Red
            '    xlWorkSheet.Cells(cnt, excelcol).AddComment("Comments:" + "Invalid baseline directory Path.")
            'End If


            If Not mTCReviewObj.isValidBaseLineDirectoryPath Then
                Dim excelRange As Range = xlWorkSheet.Cells(cnt, excelcol)
                'excelRange.Interior.Color = Color.Red
                xlWorkSheet.Cells(cnt, excelcol).AddComment("Invalid path.")
            End If


#Region "Old"
            ''1. Is the part number match with M2M? *
            'xlWorkSheet.Cells(cnt + 1, excelcol).AddComment("PartName:" + mTCReviewObj.assemblyName)

            ''2. What is the revision level? *
            'xlWorkSheet.Cells(cnt + 2, excelcol) = mTCReviewObj.revisionLevel

            ''3. Who is the Author of the file? (For BEC authors, type 3 capitalized initials of author)
            'xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.author

            ''4. What type of component?
            'xlWorkSheet.Cells(cnt + 4, excelcol) = mTCReviewObj.category

            ''5. Virtual thread applied for Fasteners?
            'xlWorkSheet.Cells(cnt + 5, excelcol) = mTCReviewObj.isVirtualThreadExists

            ''6. Are ALL features fully defined in the model Pathfinder? * (Sketch fully constraint?)
            'xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.isSketchFullyDefined

            ''7. Any suppressed feature found?
            'xlWorkSheet.Cells(cnt + 7, excelcol) = mTCReviewObj.isSuppressFeatureFound

            ''8. Is the "Mat'l Spec" field populated? (PURCHASED for library components) *
            'xlWorkSheet.Cells(cnt + 8, excelcol) = mTCReviewObj.isMaterialSpecExists
            'xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Material Specification:" + mTCReviewObj.materialSpec)

            ''9. Is the "Material Used" field populated? (PURCHASED for library components) *
            'xlWorkSheet.Cells(cnt + 9, excelcol) = mTCReviewObj.isMaterialUsedExists
            'xlWorkSheet.Cells(cnt + 9, excelcol).AddComment("Material Used:" + mTCReviewObj.materialUsed)

            ''10. Is the "Title" field populated with the correct part description? (This should MATCH the M2M Item Master (INV) Description field) *
            'xlWorkSheet.Cells(cnt + 10, excelcol) = mTCReviewObj.isMaterialDesc_Title

            'xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.title)

            ''11. Is the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) *
            'xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.isAuthorExists
            'xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("Author:" + mTCReviewObj.author)

            ''12. Is the "Keywords" field populated with the FULL M2M Item Master Description as shown in the Comments field? *
            'xlWorkSheet.Cells(cnt + 12, excelcol) = String.Empty

            ''13. Is the "Comments" field populated with the Vendor name and Vendor part number? (It should appear as VENDOR NAME = VENDOR PART NUMBER) *
            'xlWorkSheet.Cells(cnt + 13, excelcol) = mTCReviewObj.isCommentExist
            'xlWorkSheet.Cells(cnt + 13, excelcol).AddComment("Comments:" + mTCReviewObj.comment)

            ''14. Is the "Document Number" field populated with the correct part number? (This should  MATCH the M2M Item Master Part Number field) *
            'xlWorkSheet.Cells(cnt + 14, excelcol) = mTCReviewObj.isCorrectDocumentNumber
            'xlWorkSheet.Cells(cnt + 14, excelcol).AddComment("Document Number:" + mTCReviewObj.documentNumber)

            ''15. Is the "Revision" field populated with the correct revision number? *
            'xlWorkSheet.Cells(cnt + 15, excelcol) = mTCReviewObj.isCorrectRevisionNumber
            'xlWorkSheet.Cells(cnt + 15, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)

            ''16. "Is the ""Project"" field populated with the correct Project Name? (The Project should be BROOKVILLE EQUIPMENT CORP) *"
            'xlWorkSheet.Cells(cnt + 16, excelcol) = mTCReviewObj.isCorrectProjectName

            ''17. Is the "Hardware Part" box checked for hardware components? (This should be for ALL  nuts, bolts, washers, screws, etc.) *
            'xlWorkSheet.Cells(cnt + 17, excelcol) = mTCReviewObj.isHardwarePartBoxChecked
            'xlWorkSheet.Cells(cnt + 17, excelcol).AddComment("Status & HardwarePart:" + mTCReviewObj.documentType + "&" + mTCReviewObj.hardwarePart.ToString())

            ''18. Do all other unused property fields have a "dash" (-) populated? *
            'xlWorkSheet.Cells(cnt + 18, excelcol) = mTCReviewObj.isDashPopulated

            ''19. Does the unit of measures (UOM) match M2M "For Inventory UOM"? (i.e. EA, sqft, in, etc.) *
            'xlWorkSheet.Cells(cnt + 19, excelcol) = mTCReviewObj.isUOMMatch_M2M
            'xlWorkSheet.Cells(cnt + 19, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)

            ''20. Is the M2M Source marked stock/purchased? *
            'xlWorkSheet.Cells(cnt + 20, excelcol) = mTCReviewObj.isM2MSourceStocked

            ''21. Did I use the hole tool for ALL holes requiring fasteners and they match the Vendor supplied Data?(This include clearances holes For hardware, tapped holes, And Slots) *
            'xlWorkSheet.Cells(cnt + 21, excelcol) = mTCReviewObj.isHoleToolUsed

            ''22. Are all inter-part copies, part copies and included geometry broken? *
            'xlWorkSheet.Cells(cnt + 22, excelcol) = String.Empty

            ''23. inter-part copies detected?
            'xlWorkSheet.Cells(cnt + 23, excelcol) = mTCReviewObj.isInterPartCopiesDetected

            ''24. part copies detected?
            'xlWorkSheet.Cells(cnt + 24, excelcol) = mTCReviewObj.isPartCopiesDetected

            ''25. broken  file Path detected?
            'xlWorkSheet.Cells(cnt + 25, excelcol) = mTCReviewObj.isBrokenFilePathDetected

            ''26. Is the model a SE Adjustable Part? (Models should NOT be adjustable) *
            'xlWorkSheet.Cells(cnt + 26, excelcol) = mTCReviewObj.isAdjustable

            ''27. Does this model have any SE Simplified Features? (NO model should contain SE Simplified Features) *
            'xlWorkSheet.Cells(cnt + 27, excelcol) = mTCReviewObj.hasSESimplifiedFeature

            ''28. Do all hardware AND library components in this asm model have an SE Status of Baselined?(If Available Or Released status hardware/library component exists, select "NO" ) *
            'xlWorkSheet.Cells(cnt + 28, excelcol) = mTCReviewObj.hasSEStatusBaseLined

            ''29. What is the last modified date
            'xlWorkSheet.Cells(cnt + 29, excelcol) = mTCReviewObj.modifiedDate
#End Region

            '26th September 2024

            ''1. Is the part number match with M2M? *
            'xlWorkSheet.Cells(cnt + 1, excelcol).AddComment("Part Number::" + mTCReviewObj.assemblyName)
            'If (mTCReviewObj.assemblyName = "No") Then
            '    xlWorkSheet.Cells(cnt + 1, excelcol).Interior.Color = colorRed
            'End If

            '1. Is the part number match with M2M? *   
            xlWorkSheet.Cells(cnt + 1, excelcol) = mTCReviewObj.isPartFound
            xlWorkSheet.Cells(cnt + 1, excelcol).AddComment("Part Number:" + mTCReviewObj.assemblyName)
            If (mTCReviewObj.assemblyName = "No") Then
                xlWorkSheet.Cells(cnt + 1, excelcol).Interior.Color = colorRed
            End If

            '2. What is the revision level? *
            xlWorkSheet.Cells(cnt + 2, excelcol) = mTCReviewObj.revisionLevel

            '3. Who is the Author of the file? (For BEC authors, type 3 capitalized initials of author)
            xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.author
            If (mTCReviewObj.author = "No") Then
                xlWorkSheet.Cells(cnt + 3, excelcol).Interior.Color = colorRed
            End If

            '4. What type of component?
            xlWorkSheet.Cells(cnt + 4, excelcol) = mTCReviewObj.category

            '5. Virtual thread applied for Fasteners?
            xlWorkSheet.Cells(cnt + 5, excelcol) = mTCReviewObj.isVirtualThreadExists
            If (mTCReviewObj.isVirtualThreadExists = "No") Then
                xlWorkSheet.Cells(cnt + 5, excelcol).Interior.Color = colorRed
            End If

            '6. Are ALL features fully defined in the model Pathfinder? * (Sketch fully constraint?)
            xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.isSketchFullyDefined
            If (mTCReviewObj.isSketchFullyDefined = "No") Then
                xlWorkSheet.Cells(cnt + 6, excelcol).Interior.Color = colorRed
            End If

            '7. Any suppressed feature found?
            xlWorkSheet.Cells(cnt + 7, excelcol) = mTCReviewObj.isSuppressFeatureFound
            If (mTCReviewObj.isSuppressFeatureFound = "Yes") Then
                xlWorkSheet.Cells(cnt + 7, excelcol).Interior.Color = colorRed
            End If

            '8. Is the "Mat'l Spec" field populated? (PURCHASED for library components) *
            xlWorkSheet.Cells(cnt + 8, excelcol) = mTCReviewObj.isMaterialSpecExists
            xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Material Specification:" + mTCReviewObj.materialSpec)
            If (mTCReviewObj.isMaterialSpecExists = "No") Then
                xlWorkSheet.Cells(cnt + 8, excelcol).Interior.Color = colorRed
            End If

            '9. Is the "Material Used" field populated? (PURCHASED for library components) *
            xlWorkSheet.Cells(cnt + 9, excelcol) = mTCReviewObj.isMaterialUsedExists
            xlWorkSheet.Cells(cnt + 9, excelcol).AddComment("Material Used:" + mTCReviewObj.materialUsed)
            If (mTCReviewObj.isMaterialUsedExists = "No") Then
                xlWorkSheet.Cells(cnt + 9, excelcol).Interior.Color = colorRed
            End If

            '10. Is the "Title" field populated with the correct part description? (This should MATCH the M2M Item Master (INV) Description field) *
            xlWorkSheet.Cells(cnt + 10, excelcol) = mTCReviewObj.isMaterialDesc_Title

            xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.title)
            If (mTCReviewObj.title = "No") Then
                xlWorkSheet.Cells(cnt + 10, excelcol).Interior.Color = colorRed
            End If

            '11. Is the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) *
            xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.isAuthorExists
            xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("Author:" + mTCReviewObj.author)
            If (mTCReviewObj.author = "No") Then
                xlWorkSheet.Cells(cnt + 11, excelcol).Interior.Color = colorRed
            End If

            '12. Is the "Document Number" field populated with the correct part number? (This should  MATCH the M2M Item Master Part Number field) *
            xlWorkSheet.Cells(cnt + 12, excelcol) = mTCReviewObj.isCorrectDocumentNumber
            xlWorkSheet.Cells(cnt + 12, excelcol).AddComment("Document Number:" + mTCReviewObj.documentNumber)
            If (mTCReviewObj.documentNumber = "No") Then
                xlWorkSheet.Cells(cnt + 12, excelcol).Interior.Color = colorRed
            End If

            ''13. Is the "Revision" field populated with the correct revision number? *
            'xlWorkSheet.Cells(cnt + 13, excelcol) = mTCReviewObj.isCorrectRevisionNumber
            'xlWorkSheet.Cells(cnt + 13, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            'If (mTCReviewObj.revisionLevel = "No") Then
            '    xlWorkSheet.Cells(cnt + 13, excelcol).Interior.Color = colorRed
            'End If

            '13th Sep 2024
            '13. Is the "Revision" field populated with the correct revision number? *
            xlWorkSheet.Cells(cnt + 13, excelcol) = mTCReviewObj.isCorrectRevisionNumber
            xlWorkSheet.Cells(cnt + 13, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            If (mTCReviewObj.isCorrectRevisionNumber = "No") Then
                xlWorkSheet.Cells(cnt + 6, excelcol).Interior.Color = colorRed
            End If


            '14. "Is the ""Project"" field populated with the correct Project Name? (The Project should be BROOKVILLE EQUIPMENT CORP) *"
            xlWorkSheet.Cells(cnt + 14, excelcol) = mTCReviewObj.isCorrectProjectName
            If (mTCReviewObj.isCorrectProjectName = "No") Then
                xlWorkSheet.Cells(cnt + 14, excelcol).Interior.Color = colorRed
            End If

            '15. Is the "Hardware Part" box checked for hardware components? (This should be for ALL  nuts, bolts, washers, screws, etc.) *
            xlWorkSheet.Cells(cnt + 15, excelcol) = mTCReviewObj.isHardwarePartBoxChecked
            xlWorkSheet.Cells(cnt + 15, excelcol).AddComment("Status & HardwarePart:" + mTCReviewObj.documentType + "&" + mTCReviewObj.hardwarePart.ToString())
            If (mTCReviewObj.isHardwarePartBoxChecked = "No") Then
                xlWorkSheet.Cells(cnt + 15, excelcol).Interior.Color = colorRed
            End If

            '16. Do all other unused property fields have a "dash" (-) populated? *
            xlWorkSheet.Cells(cnt + 16, excelcol) = mTCReviewObj.isDashPopulated
            If (mTCReviewObj.isDashPopulated = "No") Then
                xlWorkSheet.Cells(cnt + 16, excelcol).Interior.Color = colorRed
            End If

            ''17. Does the unit of measures (UOM) match M2M "For Inventory UOM"? (i.e. EA, sqft, in, etc.) *
            'xlWorkSheet.Cells(cnt + 17, excelcol) = mTCReviewObj.isUOMMatch_M2M
            'xlWorkSheet.Cells(cnt + 17, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
            'If (mTCReviewObj.UomProperty = "No") Then
            '    xlWorkSheet.Cells(cnt + 17, excelcol).Interior.Color = colorRed
            'End If


            '17. Does the unit of measures (UOM) match M2M "For Inventory UOM"? (i.e. EA, sqft, in, etc.) *
            If mTCReviewObj.isUOMMatch_M2M = "Yes" Then
                '26th September 2024
                If Not mTCReviewObj.UomProperty = "" Then
                    xlWorkSheet.Cells(cnt + 17, excelcol) = mTCReviewObj.isUOMMatch_M2M
                    xlWorkSheet.Cells(cnt + 17, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
                End If
            Else
                '26th September 2024
                If Not mTCReviewObj.UomProperty = "" Then
                    xlWorkSheet.Cells(cnt + 17, excelcol) = mTCReviewObj.isUOMMatch_M2M
                    xlWorkSheet.Cells(cnt + 17, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
                    xlWorkSheet.Cells(cnt + 17, excelcol).Interior.Color = colorRed
                End If
            End If


            '18. Is the M2M Source marked stock/purchased? *
            xlWorkSheet.Cells(cnt + 18, excelcol) = mTCReviewObj.isM2MSourceStocked
            If (mTCReviewObj.isM2MSourceStocked = "No") Then
                xlWorkSheet.Cells(cnt + 18, excelcol).Interior.Color = colorRed
            End If

            '19. Did I use the hole tool for ALL holes requiring fasteners and they match the Vendor supplied Data?(This include clearances holes For hardware, tapped holes, And Slots) *
            xlWorkSheet.Cells(cnt + 19, excelcol) = mTCReviewObj.isHoleToolUsed
            If (mTCReviewObj.isHoleToolUsed = "No") Then
                xlWorkSheet.Cells(cnt + 19, excelcol).Interior.Color = colorRed
            End If

            '20. inter-part copies detected?
            xlWorkSheet.Cells(cnt + 20, excelcol) = mTCReviewObj.isInterPartCopiesDetected
            If (mTCReviewObj.isInterPartCopiesDetected = "Yes") Then
                xlWorkSheet.Cells(cnt + 20, excelcol).Interior.Color = colorRed
            End If

            '21. part copies detected?
            xlWorkSheet.Cells(cnt + 21, excelcol) = mTCReviewObj.isPartCopiesDetected
            If (mTCReviewObj.isPartCopiesDetected = "Yes") Then
                xlWorkSheet.Cells(cnt + 21, excelcol).Interior.Color = colorRed
            End If

            '22. broken  file Path detected?
            xlWorkSheet.Cells(cnt + 22, excelcol) = mTCReviewObj.isBrokenFilePathDetected
            If (mTCReviewObj.isBrokenFilePathDetected = "Yes") Then
                xlWorkSheet.Cells(cnt + 22, excelcol).Interior.Color = colorRed
            End If

            '23. Is the model a SE Adjustable Part? (Models should NOT be adjustable) *
            xlWorkSheet.Cells(cnt + 23, excelcol) = mTCReviewObj.isAdjustable
            If (mTCReviewObj.isAdjustable = "Yes") Then
                xlWorkSheet.Cells(cnt + 23, excelcol).Interior.Color = colorRed
            End If

            '24. Does this model have any SE Simplified Features? (NO model should contain SE Simplified Features) *
            xlWorkSheet.Cells(cnt + 24, excelcol) = mTCReviewObj.hasSESimplifiedFeature
            If (mTCReviewObj.hasSESimplifiedFeature = "Yes") Then
                xlWorkSheet.Cells(cnt + 24, excelcol).Interior.Color = colorRed
            End If

            '25. Do all hardware AND library components in this asm model have an SE Status of Baselined?(If Available Or Released status hardware/library component exists, select "NO" ) *
            xlWorkSheet.Cells(cnt + 25, excelcol) = mTCReviewObj.hasSEStatusBaseLined
            If (mTCReviewObj.hasSEStatusBaseLined = "No") Then
                xlWorkSheet.Cells(cnt + 25, excelcol).Interior.Color = colorRed
            End If

            '26. What is the last modified date
            xlWorkSheet.Cells(cnt + 26, excelcol) = mTCReviewObj.modifiedDate

            excelcol = excelcol + 1

            '5th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).Interior.Color = System.Drawing.Color.FromArgb(165, 165, 165)  '9th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(cnt, excelcol).font.Bold = True
            xlWorkSheet.Cells(cnt, excelcol) = "Issue Addressed?"

            '9th Sep 2024
            Dim borders As Borders = xlWorkSheet.Cells(cnt, excelcol).Borders
            borders.LineStyle = XlLineStyle.xlContinuous
            borders.Weight = 2D

            AddDropDownListToExcelColumn(xlWorkSheet, 42, excelcol)

            excelcol = excelcol + 1
        Next
    End Sub

    Private Sub SetElectricalPartWorkSheet_MTC(ByRef xlWorkSheet As Worksheet, ByVal finalPartListBEC As List(Of MTC_Electrical), ByVal finalPartListDGS As List(Of MTC_Electrical), MTC_Author_Name As String)

        '4th Sep 2024
        'Author Info and time span
        'xlWorkSheet.Cells(1, 1).Interior.Color = System.Drawing.Color.Gray
        'xlWorkSheet.Cells(1, 1).font.Color = System.Drawing.Color.White
        xlWorkSheet.Cells(1, 1).font.Bold = True
        'xlWorkSheet.Cells(1, 1) = $"{MTC_Author_Name}_{System.DateTime.Now.ToString("MMM_d_yyyy_HH_mm")}"
        xlWorkSheet.Cells(1, 3) = $"{MTC_Author_Name} {System.DateTime.Now.ToString("(MMM_d_yyyy_HH_mm)")}"

        '4th Sep 2024
        'Dim cnt As Integer = 1
        Dim cnt As Integer = 2

        Dim excelcol As Integer = 4

        For Each mTCReviewObj As MTC_Electrical In finalPartListBEC

            '4th Sep 2024
            ''0. Assembly Name
            'xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
            'xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
            'xlWorkSheet.Cells(1, excelcol).font.Bold = True
            'xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.assemblyName

            '0. Assembly Name
            xlWorkSheet.Cells(cnt, excelcol).Interior.Color = System.Drawing.Color.FromArgb(165, 165, 165)  '9th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(cnt, excelcol).font.Bold = True
            xlWorkSheet.Cells(cnt, excelcol) = mTCReviewObj.assemblyName

            '9th Sep 2024
            Dim borders1 As Borders = xlWorkSheet.Cells(cnt, excelcol).Borders
            borders1.LineStyle = XlLineStyle.xlContinuous
            borders1.Weight = 2D


            '4th Sep 2024
            'SetExcelColumnColorAndBorder(xlWorkSheet, 11, excelcol)
            SetExcelColumnColorAndBorderNewMTC(xlWorkSheet, 11, excelcol)


            '1. Eco Number
            xlWorkSheet.Cells(cnt + 1, excelcol) = mTCReviewObj.ecoNumber
            '26th september 2024
            If mTCReviewObj.ecoNumber = "" Then
                xlWorkSheet.Cells(cnt + 1, excelcol) = "0"
            End If
            If mTCReviewObj.revisionLevel = "0" Then
                xlWorkSheet.Cells(cnt + 1, excelcol).AddComment($"SOW = {mTCReviewObj.ecoNumber}")
            Else
                xlWorkSheet.Cells(cnt + 1, excelcol).AddComment($"ECO = {mTCReviewObj.ecoNumber}")
            End If

            '2. Part Number
            xlWorkSheet.Cells(cnt + 2, excelcol) = mTCReviewObj.partNumber
            If mTCReviewObj.isPartFound = True Then
                xlWorkSheet.Cells(cnt + 2, excelcol).AddComment("Part Number:" + mTCReviewObj.assemblyName)
            Else
                xlWorkSheet.Cells(cnt + 2, excelcol).AddComment("Part Number:" + mTCReviewObj.assemblyName)
            End If


            '3. Revision Level
            xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.revisionLevel

            '4. Author
            xlWorkSheet.Cells(cnt + 4, excelcol) = mTCReviewObj.author

            '5. Project Name
            xlWorkSheet.Cells(cnt + 5, excelcol) = mTCReviewObj.projectNameExist
            xlWorkSheet.Cells(cnt + 5, excelcol).AddComment("ProjectName:" + mTCReviewObj.projectName)

            ''6. Revision Number Correct
            'xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.revisionNumberCorrect
            'If mTCReviewObj.revisionNumberCorrect = "Yes" Then
            '    xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            'Else
            '    xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            'End If

            '13th Sep 2024
            '6. Revision Number Correct
            xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.revisionNumberCorrect
            xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            If (mTCReviewObj.revisionNumberCorrect = "No") Then
                xlWorkSheet.Cells(cnt + 6, excelcol).Interior.Color = colorRed
            End If

            '7. Document Number correct
            xlWorkSheet.Cells(cnt + 7, excelcol) = mTCReviewObj.documentNumberCorrect
            If mTCReviewObj.documentNumberCorrect = "Yes" Then

                xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Document Number:" + mTCReviewObj.documentNumber)
            Else
                xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Document Number:" + mTCReviewObj.documentNumber)
            End If

            '8. Is the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name)
            xlWorkSheet.Cells(cnt + 8, excelcol) = mTCReviewObj.authorExists
            If mTCReviewObj.authorExists = "Yes" Then
                xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Author:" + mTCReviewObj.author)
            Else
                xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Author:" + mTCReviewObj.author)
            End If

            '9. Do all technically unused properties have a "dash" populated?
            xlWorkSheet.Cells(cnt + 9, excelcol) = mTCReviewObj.isDashPopulated

            '10. Does the Model Title MATCH the M2M Item Master (INV) Description Field?
            xlWorkSheet.Cells(cnt + 10, excelcol) = mTCReviewObj.isTitleMatch_ItemMaster
            If mTCReviewObj.isTitleMatch_ItemMaster = "Yes" Then

                xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.title)
            Else
                xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.title)
            End If


            '11. What is the last modified date?
            xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.modifiedDate


            excelcol = excelcol + 1

            '5th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).Interior.Color = System.Drawing.Color.FromArgb(165, 165, 165)  '9th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(cnt, excelcol).font.Bold = True
            xlWorkSheet.Cells(cnt, excelcol) = "Issue Addressed?"

            '9th Sep 2024
            Dim borders As Borders = xlWorkSheet.Cells(cnt, excelcol).Borders
            borders.LineStyle = XlLineStyle.xlContinuous
            borders.Weight = 2D

            AddDropDownListToExcelColumn(xlWorkSheet, 11, excelcol)

            excelcol = excelcol + 1
        Next

        '-------------------------------------------------------------------------
        'temp 20Feb2024

        For Each mTCReviewObj As MTC_Electrical In finalPartListDGS

            '4th Sep 2024
            ''0. Assembly Name
            'xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
            'xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
            'xlWorkSheet.Cells(1, excelcol).font.Bold = True
            'xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.assemblyName

            '0. Assembly Name
            xlWorkSheet.Cells(cnt, excelcol).Interior.Color = System.Drawing.Color.FromArgb(165, 165, 165)  '9th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(cnt, excelcol).font.Bold = True
            xlWorkSheet.Cells(cnt, excelcol) = mTCReviewObj.assemblyName

            '9th Sep 2024
            Dim borders1 As Borders = xlWorkSheet.Cells(cnt, excelcol).Borders
            borders1.LineStyle = XlLineStyle.xlContinuous
            borders1.Weight = 2D

            '4th Sep 2024
            'SetExcelColumnColorAndBorder(xlWorkSheet, 11, excelcol)
            SetExcelColumnColorAndBorderNewMTC(xlWorkSheet, 11, excelcol)


            '1. Eco Number
            '26th september 2024
            xlWorkSheet.Cells(cnt + 1, excelcol) = mTCReviewObj.ecoNumber
            If mTCReviewObj.ecoNumber = "" Then
                xlWorkSheet.Cells(cnt + 1, excelcol) = "0"
            End If
            If mTCReviewObj.revisionLevel = "0" Then
                xlWorkSheet.Cells(cnt + 1, excelcol).AddComment($"SOW = {mTCReviewObj.ecoNumber}")
            Else
                xlWorkSheet.Cells(cnt + 1, excelcol).AddComment($"ECO = {mTCReviewObj.ecoNumber}")
            End If

            '2. Part Number
            xlWorkSheet.Cells(cnt + 2, excelcol) = mTCReviewObj.partNumber
            If mTCReviewObj.isPartFound = True Then
                xlWorkSheet.Cells(cnt + 2, excelcol).AddComment("Part Number:" + mTCReviewObj.assemblyName)
            Else
                xlWorkSheet.Cells(cnt + 2, excelcol).AddComment("Part Number:" + mTCReviewObj.assemblyName)
            End If


            '3. Revision Level
            xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.revisionLevel

            '4. Author
            xlWorkSheet.Cells(cnt + 4, excelcol) = mTCReviewObj.author

            '5. Project Name
            xlWorkSheet.Cells(cnt + 5, excelcol) = mTCReviewObj.projectNameExist
            xlWorkSheet.Cells(cnt + 5, excelcol).AddComment("ProjectName:" + mTCReviewObj.projectName)

            ''6. Revision Number Correct
            'xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.revisionNumberCorrect
            'If mTCReviewObj.revisionNumberCorrect = "Yes" Then
            '    xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            'Else
            '    xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            'End If

            '13th Sep 2024
            '6. Revision Number Correct
            xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.revisionNumberCorrect
            xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            If (mTCReviewObj.revisionNumberCorrect = "No") Then
                xlWorkSheet.Cells(cnt + 6, excelcol).Interior.Color = colorRed
            End If

            '7. Document Number correct
            xlWorkSheet.Cells(cnt + 7, excelcol) = mTCReviewObj.documentNumberCorrect
            If mTCReviewObj.documentNumberCorrect = "Yes" Then

                xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Document Number:" + mTCReviewObj.documentNumber)
            Else
                xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Document Number:" + mTCReviewObj.documentNumber)
            End If

            '8. Is the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name)
            xlWorkSheet.Cells(cnt + 8, excelcol) = mTCReviewObj.authorExists
            If mTCReviewObj.authorExists = "Yes" Then
                xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Author:" + mTCReviewObj.author)
            Else
                xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Author:" + mTCReviewObj.author)
            End If

            '9. Do all technically unused properties have a "dash" populated?
            xlWorkSheet.Cells(cnt + 9, excelcol) = mTCReviewObj.isDashPopulated

            '10. Does the Model Title MATCH the M2M Item Master (INV) Description Field?
            xlWorkSheet.Cells(cnt + 10, excelcol) = mTCReviewObj.isTitleMatch_ItemMaster
            If mTCReviewObj.isTitleMatch_ItemMaster = "Yes" Then

                xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.title)
            Else
                xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.title)
            End If


            '11. What is the last modified date?
            xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.modifiedDate


            excelcol = excelcol + 1

            '5th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).Interior.Color = System.Drawing.Color.FromArgb(165, 165, 165)  '9th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(cnt, excelcol).font.Bold = True
            xlWorkSheet.Cells(cnt, excelcol) = "Issue Addressed?"

            '9th Sep 2024
            Dim borders As Borders = xlWorkSheet.Cells(cnt, excelcol).Borders
            borders.LineStyle = XlLineStyle.xlContinuous
            borders.Weight = 2D

            AddDropDownListToExcelColumn(xlWorkSheet, 11, excelcol)

            excelcol = excelcol + 1
        Next
    End Sub

    Private Sub SetPartWorkSheet_MTC(ByRef xlWorkSheet As Worksheet, ByVal finalAssemblyListBEC As List(Of MTC_Part), ByVal finalAssemblyListDGS As List(Of MTC_Part), MTC_Author_Name As String)

        '4th Sep 2024
        'Author Info and time span
        'xlWorkSheet.Cells(1, 1).Interior.Color = System.Drawing.Color.Gray
        'xlWorkSheet.Cells(1, 1).font.Color = System.Drawing.Color.White
        xlWorkSheet.Cells(1, 1).font.Bold = True
        'xlWorkSheet.Cells(1, 1) = $"{MTC_Author_Name}_{System.DateTime.Now.ToString("MMM_d_yyyy_HH_mm")}"
        xlWorkSheet.Cells(1, 3) = $"{MTC_Author_Name} {System.DateTime.Now.ToString("(MMM_d_yyyy_HH_mm)")}"

        '4th Sep 2024
        'Dim cnt As Integer = 1

        Dim cnt As Integer = 2

        Dim excelcol As Integer = 4

        For Each mTCReviewObj As MTC_Part In finalAssemblyListBEC


            If mTCReviewObj.assemblyName.Contains("210-03378") Then
                Debug.Print("aaaa")
            End If

            '4th Sep 2024
            'If mTCReviewObj.isValidPart = False Then
            '    Dim excelRange As Range = xlWorkSheet.Cells(1, excelcol)
            '    excelRange.Interior.Color = Color.Red
            '    xlWorkSheet.Cells(1, excelcol).AddComment($"Invalid part.")
            'End If



            '26th September 2024

            'If mTCReviewObj.isValidPart = False Then
            '    Dim excelRange As Range = xlWorkSheet.Cells(cnt, excelcol)
            '    excelRange.Interior.Color = Color.Red
            '    xlWorkSheet.Cells(cnt, excelcol).AddComment($"Invalid part.")
            'End If

            If Not mTCReviewObj.isValidPart = False Then
                Dim excelRange As Range = xlWorkSheet.Cells(cnt, excelcol)
                'excelRange.Interior.Color = Color.Red
                xlWorkSheet.Cells(cnt, excelcol).AddComment("Invalid path.")
            End If


            '4th Sep 2024
            ''0. Assembly Name
            'xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
            'xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
            'xlWorkSheet.Cells(1, excelcol).font.Bold = True
            'xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.assemblyName

            '0. Assembly Name
            xlWorkSheet.Cells(cnt, excelcol).Interior.Color = System.Drawing.Color.FromArgb(165, 165, 165)  '9th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(cnt, excelcol).font.Bold = True
            xlWorkSheet.Cells(cnt, excelcol) = mTCReviewObj.assemblyName

            '9th Sep 2024
            Dim borders1 As Borders = xlWorkSheet.Cells(cnt, excelcol).Borders
            borders1.LineStyle = XlLineStyle.xlContinuous
            borders1.Weight = 2D

            '4th Sep 2024
            'SetExcelColumnColorAndBorder(xlWorkSheet, 22, excelcol)
            SetExcelColumnColorAndBorderNewMTC(xlWorkSheet, 22, excelcol)

            '1. Eco Number
            xlWorkSheet.Cells(cnt + 1, excelcol) = mTCReviewObj.ecoNumber
            If mTCReviewObj.revisionLevel = "0" Then
                xlWorkSheet.Cells(cnt + 1, excelcol).AddComment($"SOW = {mTCReviewObj.ecoNumber}")
            Else
                xlWorkSheet.Cells(cnt + 1, excelcol).AddComment($"ECO = {mTCReviewObj.ecoNumber}")
            End If

            '2. Part Number
            xlWorkSheet.Cells(cnt + 2, excelcol) = mTCReviewObj.partNumber
            If mTCReviewObj.isPartFound = True Then
                xlWorkSheet.Cells(cnt + 2, excelcol).AddComment("Part Number:" + mTCReviewObj.assemblyName)
            Else
                xlWorkSheet.Cells(cnt + 2, excelcol).AddComment("Part Number:" + mTCReviewObj.assemblyName)

                xlWorkSheet.Cells(cnt + 2, excelcol).Interior.Color = colorRed

            End If

            '3. Revision Level
            xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.revisionLevel

            '4. Author
            xlWorkSheet.Cells(cnt + 4, excelcol) = mTCReviewObj.author
            If (mTCReviewObj.author = "No") Then
                xlWorkSheet.Cells(cnt + 4, excelcol).Interior.Color = colorRed
            End If

            '5. Is the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name)
            xlWorkSheet.Cells(cnt + 5, excelcol) = mTCReviewObj.authorExists

            If mTCReviewObj.authorExists = "Yes" Then
                xlWorkSheet.Cells(cnt + 5, excelcol).AddComment("Author:" + mTCReviewObj.author)
            Else
                xlWorkSheet.Cells(cnt + 5, excelcol).AddComment("Author:" + mTCReviewObj.author)
                xlWorkSheet.Cells(cnt + 5, excelcol).Interior.Color = colorRed
            End If

            '6. Project Name
            xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.projectNameExist
            xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("ProjectName:" + mTCReviewObj.projectName)
            If (mTCReviewObj.projectNameExist = "No") Then
                xlWorkSheet.Cells(cnt + 6, excelcol).Interior.Color = colorRed
            End If


            ''7. Revision Number Correct
            'xlWorkSheet.Cells(cnt + 7, excelcol) = mTCReviewObj.revisionNumberCorrect
            'If mTCReviewObj.revisionNumberCorrect = "Yes" Then
            '    xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            'Else
            '    xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            '    xlWorkSheet.Cells(cnt + 7, excelcol).Interior.Color = colorRed
            'End If

            '13th Sep 2024
            '7. Revision Number Correct
            xlWorkSheet.Cells(cnt + 7, excelcol) = mTCReviewObj.revisionNumberCorrect
            xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            If (mTCReviewObj.revisionNumberCorrect = "No") Then
                xlWorkSheet.Cells(cnt + 7, excelcol).Interior.Color = colorRed
            End If

            '8. Document Number correct
            xlWorkSheet.Cells(cnt + 8, excelcol) = mTCReviewObj.documentNumberCorrect
            If mTCReviewObj.documentNumberCorrect = "Yes" Then

                xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Document Number:" + mTCReviewObj.documentNumber)
            Else
                xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Document Number:" + mTCReviewObj.documentNumber)
                xlWorkSheet.Cells(cnt + 8, excelcol).Interior.Color = colorRed
            End If

            '9. Do all technically unused properties have a "dash" populated?
            xlWorkSheet.Cells(cnt + 9, excelcol) = mTCReviewObj.isDashPopulated
            If (mTCReviewObj.isDashPopulated = "No") Then
                xlWorkSheet.Cells(cnt + 9, excelcol).Interior.Color = colorRed
            End If


            '10. Does the Model Title MATCH the M2M Item Master (INV) Description Field?
            xlWorkSheet.Cells(cnt + 10, excelcol) = mTCReviewObj.isTitleMatch_ItemMaster
            If mTCReviewObj.isTitleMatch_ItemMaster = "Yes" Then

                xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.title)
            Else
                xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.title)
                xlWorkSheet.Cells(cnt + 10, excelcol).Interior.Color = colorRed
            End If

            '11. Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
            If mTCReviewObj.isUOMMatch_M2M = "Yes" Then 'If m2MDataObj.M2Mmeasure = mTCReviewObj.UomProperty Then
                '26th September 2024
                If Not mTCReviewObj.UomProperty = "" Then
                    xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.isUOMMatch_M2M
                    xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
                End If
            Else
                '26th September 2024
                If Not mTCReviewObj.UomProperty = "" Then
                    xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.isUOMMatch_M2M
                    xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
                    xlWorkSheet.Cells(cnt + 11, excelcol).Interior.Color = colorRed
                End If
            End If

            '12. Is the mat'l spec field populated? (indicated by RED in metadata audit file)
            xlWorkSheet.Cells(cnt + 12, excelcol) = mTCReviewObj.isMaterialSpecExists
            xlWorkSheet.Cells(cnt + 12, excelcol).AddComment("Material Specification:" + mTCReviewObj.materialSpec)
            If (mTCReviewObj.isMaterialSpecExists = "No") Then
                xlWorkSheet.Cells(cnt + 12, excelcol).Interior.Color = colorRed
            End If


            '13. Is the material used field populated? *
            xlWorkSheet.Cells(cnt + 13, excelcol) = mTCReviewObj.isMaterialUsedExists
            xlWorkSheet.Cells(cnt + 13, excelcol).AddComment("Material Used:" + mTCReviewObj.materialUsed)
            If (mTCReviewObj.isMaterialUsedExists = "No") Then
                xlWorkSheet.Cells(cnt + 13, excelcol).Interior.Color = colorRed
            End If

            '14. Did I use the hole tool for ALL holes requiring fasteners? (include clearance holes for hardware, tapped holes, And Slots) *
            '37
            xlWorkSheet.Cells(cnt + 14, excelcol) = mTCReviewObj.isHoleToolUsed
            If (mTCReviewObj.isHoleToolUsed = "No") Then
                xlWorkSheet.Cells(cnt + 14, excelcol).Interior.Color = colorRed
            End If


            '15. Is sketch fully constrain?
            xlWorkSheet.Cells(cnt + 15, excelcol) = $"{mTCReviewObj.isSketchFullyConstraint}"
            If (mTCReviewObj.isSketchFullyConstraint = "No") Then
                xlWorkSheet.Cells(cnt + 15, excelcol).Interior.Color = colorRed
            End If


            '16. Have any suppressed (unused) features been removed from the model Pathfinder? *
            xlWorkSheet.Cells(cnt + 16, excelcol) = mTCReviewObj.haveSuppressedFeatureRemoved
            If (mTCReviewObj.haveSuppressedFeatureRemoved = "Yes") Then
                xlWorkSheet.Cells(cnt + 16, excelcol).Interior.Color = colorRed
            End If

            '17. Is the part "Adjustable"? (part should NOT be adjustable) *
            '19
            xlWorkSheet.Cells(cnt + 17, excelcol) = mTCReviewObj.isAdjustable
            If (mTCReviewObj.isAdjustable = "Yes") Then
                xlWorkSheet.Cells(cnt + 17, excelcol).Interior.Color = colorRed
            End If


            '18. Does the M2M Source reflect part creation (Make child of stk/pur parent, Project  procurement strategy, material types Like composite, etc.)? *
            '23
            If mTCReviewObj.m2mSource = "" Then

                xlWorkSheet.Cells(cnt + 18, excelcol) = "No"
                xlWorkSheet.Cells(cnt + 18, excelcol).Interior.Color = colorRed
            Else

                xlWorkSheet.Cells(cnt + 18, excelcol) = "Yes"

                xlWorkSheet.Cells(cnt + 18, excelcol).AddComment(mTCReviewObj.m2mSource)

            End If

            '19
            xlWorkSheet.Cells(cnt + 19, excelcol) = mTCReviewObj.modifiedDate

            excelcol = excelcol + 1

            '5th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).Interior.Color = System.Drawing.Color.FromArgb(165, 165, 165)  '9th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(cnt, excelcol).font.Bold = True
            xlWorkSheet.Cells(cnt, excelcol) = "Issue Addressed?"

            '9th Sep 2024
            Dim borders As Borders = xlWorkSheet.Cells(cnt, excelcol).Borders
            borders.LineStyle = XlLineStyle.xlContinuous
            borders.Weight = 2D

            AddDropDownListToExcelColumn(xlWorkSheet, 22, excelcol)

            excelcol = excelcol + 1
        Next

        '--------------------------------------------------------
        'temp 20Feb2024

        For Each mTCReviewObj As MTC_Part In finalAssemblyListDGS


            If mTCReviewObj.assemblyName.Contains("210-03378") Then
                Debug.Print("aaaa")
            End If

            '4th Sep 2024
            'If mTCReviewObj.isValidPart = False Then
            '    Dim excelRange As Range = xlWorkSheet.Cells(1, excelcol)
            '    excelRange.Interior.Color = Color.Red
            '    xlWorkSheet.Cells(1, excelcol).AddComment($"Invalid part.")
            'End If


            '26th September 2024

            'If mTCReviewObj.isValidPart = False Then
            '    Dim excelRange As Range = xlWorkSheet.Cells(cnt, excelcol)
            '    excelRange.Interior.Color = Color.Red
            '    xlWorkSheet.Cells(cnt, excelcol).AddComment($"Invalid part.")
            'End If

            If Not mTCReviewObj.isValidPart = False Then
                Dim excelRange As Range = xlWorkSheet.Cells(cnt, excelcol)
                'excelRange.Interior.Color = Color.Red
                xlWorkSheet.Cells(cnt, excelcol).AddComment("Invalid path.")
            End If


            '4th Sep 2024
            ''0. Assembly Name
            'xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
            'xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
            'xlWorkSheet.Cells(1, excelcol).font.Bold = True
            'xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.assemblyName

            '0. Assembly Name
            xlWorkSheet.Cells(cnt, excelcol).Interior.Color = System.Drawing.Color.FromArgb(165, 165, 165)  '9th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(cnt, excelcol).font.Bold = True
            xlWorkSheet.Cells(cnt, excelcol) = mTCReviewObj.assemblyName

            '9th Sep 2024
            Dim borders1 As Borders = xlWorkSheet.Cells(cnt, excelcol).Borders
            borders1.LineStyle = XlLineStyle.xlContinuous
            borders1.Weight = 2D

            '4th Sep 2024
            'SetExcelColumnColorAndBorder(xlWorkSheet, 22, excelcol)
            SetExcelColumnColorAndBorderNewMTC(xlWorkSheet, 22, excelcol)

            '1. Eco Number
            xlWorkSheet.Cells(cnt + 1, excelcol) = mTCReviewObj.ecoNumber
            If mTCReviewObj.revisionLevel = "0" Then
                xlWorkSheet.Cells(cnt + 1, excelcol).AddComment($"SOW = {mTCReviewObj.ecoNumber}")
            Else
                xlWorkSheet.Cells(cnt + 1, excelcol).AddComment($"ECO = {mTCReviewObj.ecoNumber}")
            End If

            '2. Part Number
            xlWorkSheet.Cells(cnt + 2, excelcol) = mTCReviewObj.partNumber
            If mTCReviewObj.isPartFound = True Then
                xlWorkSheet.Cells(cnt + 2, excelcol).AddComment("Part Number:" + mTCReviewObj.assemblyName)
            Else
                xlWorkSheet.Cells(cnt + 2, excelcol).AddComment("Part Number:" + mTCReviewObj.assemblyName)

                xlWorkSheet.Cells(cnt + 2, excelcol).Interior.Color = colorRed

            End If

            '3. Revision Level
            xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.revisionLevel

            '4. Author
            xlWorkSheet.Cells(cnt + 4, excelcol) = mTCReviewObj.author
            If (mTCReviewObj.author = "No") Then
                xlWorkSheet.Cells(cnt + 4, excelcol).Interior.Color = colorRed
            End If

            '5. Is the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name)
            xlWorkSheet.Cells(cnt + 5, excelcol) = mTCReviewObj.authorExists

            If mTCReviewObj.authorExists = "Yes" Then
                xlWorkSheet.Cells(cnt + 5, excelcol).AddComment("Author:" + mTCReviewObj.author)
            Else
                xlWorkSheet.Cells(cnt + 5, excelcol).AddComment("Author:" + mTCReviewObj.author)
                xlWorkSheet.Cells(cnt + 5, excelcol).Interior.Color = colorRed
            End If

            '6. Project Name
            xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.projectNameExist
            xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("ProjectName:" + mTCReviewObj.projectName)
            If (mTCReviewObj.projectNameExist = "No") Then
                xlWorkSheet.Cells(cnt + 6, excelcol).Interior.Color = colorRed
            End If


            ''7. Revision Number Correct
            'xlWorkSheet.Cells(cnt + 7, excelcol) = mTCReviewObj.revisionNumberCorrect
            'If mTCReviewObj.revisionNumberCorrect = "Yes" Then
            '    xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            'Else
            '    xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            '    xlWorkSheet.Cells(cnt + 7, excelcol).Interior.Color = colorRed
            'End If

            '13th Sep 2024
            '7. Revision Number Correct
            xlWorkSheet.Cells(cnt + 7, excelcol) = mTCReviewObj.revisionNumberCorrect
            xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            If (mTCReviewObj.revisionNumberCorrect = "No") Then
                xlWorkSheet.Cells(cnt + 7, excelcol).Interior.Color = colorRed
            End If

            '8. Document Number correct
            xlWorkSheet.Cells(cnt + 8, excelcol) = mTCReviewObj.documentNumberCorrect
            If mTCReviewObj.documentNumberCorrect = "Yes" Then

                xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Document Number:" + mTCReviewObj.documentNumber)
            Else
                xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Document Number:" + mTCReviewObj.documentNumber)
                xlWorkSheet.Cells(cnt + 8, excelcol).Interior.Color = colorRed
            End If

            '9. Do all technically unused properties have a "dash" populated?
            xlWorkSheet.Cells(cnt + 9, excelcol) = mTCReviewObj.isDashPopulated
            If (mTCReviewObj.isDashPopulated = "No") Then
                xlWorkSheet.Cells(cnt + 9, excelcol).Interior.Color = colorRed
            End If


            '10. Does the Model Title MATCH the M2M Item Master (INV) Description Field?
            xlWorkSheet.Cells(cnt + 10, excelcol) = mTCReviewObj.isTitleMatch_ItemMaster
            If mTCReviewObj.isTitleMatch_ItemMaster = "Yes" Then

                xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.title)
            Else
                xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.title)
                xlWorkSheet.Cells(cnt + 10, excelcol).Interior.Color = colorRed
            End If

            ''11. Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
            'xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.isUOMMatch_M2M
            'If mTCReviewObj.isUOMMatch_M2M = "Yes" Then 'If m2MDataObj.M2Mmeasure = mTCReviewObj.UomProperty Then
            '    xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
            'Else
            '    xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
            '    xlWorkSheet.Cells(cnt + 11, excelcol).Interior.Color = colorRed
            'End If

            '11. Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
            If mTCReviewObj.isUOMMatch_M2M = "Yes" Then 'If m2MDataObj.M2Mmeasure = mTCReviewObj.UomProperty Then
                '26th September 2024
                If Not mTCReviewObj.UomProperty = "" Then
                    xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.isUOMMatch_M2M
                    xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
                End If
            Else
                '26th September 2024
                If Not mTCReviewObj.UomProperty = "" Then
                    xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.isUOMMatch_M2M
                    xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
                    xlWorkSheet.Cells(cnt + 11, excelcol).Interior.Color = colorRed
                End If
            End If

            '12. Is the mat'l spec field populated? (indicated by RED in metadata audit file)
            xlWorkSheet.Cells(cnt + 12, excelcol) = mTCReviewObj.isMaterialSpecExists
            xlWorkSheet.Cells(cnt + 12, excelcol).AddComment("Material Specification:" + mTCReviewObj.materialSpec)
            If (mTCReviewObj.isMaterialSpecExists = "No") Then
                xlWorkSheet.Cells(cnt + 12, excelcol).Interior.Color = colorRed
            End If


            '13. Is the material used field populated? *
            xlWorkSheet.Cells(cnt + 13, excelcol) = mTCReviewObj.isMaterialUsedExists
            xlWorkSheet.Cells(cnt + 13, excelcol).AddComment("Material Used:" + mTCReviewObj.materialUsed)
            If (mTCReviewObj.isMaterialUsedExists = "No") Then
                xlWorkSheet.Cells(cnt + 13, excelcol).Interior.Color = colorRed
            End If

            '14. Did I use the hole tool for ALL holes requiring fasteners? (include clearance holes for hardware, tapped holes, And Slots) *
            '37
            xlWorkSheet.Cells(cnt + 14, excelcol) = mTCReviewObj.isHoleToolUsed
            If (mTCReviewObj.isHoleToolUsed = "No") Then
                xlWorkSheet.Cells(cnt + 14, excelcol).Interior.Color = colorRed
            End If


            '15. Is sketch fully constrain?
            xlWorkSheet.Cells(cnt + 15, excelcol) = $"{mTCReviewObj.isSketchFullyConstraint}"
            If (mTCReviewObj.isSketchFullyConstraint = "No") Then
                xlWorkSheet.Cells(cnt + 15, excelcol).Interior.Color = colorRed
            End If


            '16. Have any suppressed (unused) features been removed from the model Pathfinder? *
            xlWorkSheet.Cells(cnt + 16, excelcol) = mTCReviewObj.haveSuppressedFeatureRemoved
            If (mTCReviewObj.haveSuppressedFeatureRemoved = "Yes") Then
                xlWorkSheet.Cells(cnt + 16, excelcol).Interior.Color = colorRed
            End If

            '17. Is the part "Adjustable"? (part should NOT be adjustable) *
            '19
            xlWorkSheet.Cells(cnt + 17, excelcol) = mTCReviewObj.isAdjustable
            If (mTCReviewObj.isAdjustable = "Yes") Then
                xlWorkSheet.Cells(cnt + 17, excelcol).Interior.Color = colorRed
            End If


            '18. Does the M2M Source reflect part creation (Make child of stk/pur parent, Project  procurement strategy, material types Like composite, etc.)? *
            '23
            If mTCReviewObj.m2mSource = "" Then

                xlWorkSheet.Cells(cnt + 18, excelcol) = "No"
                xlWorkSheet.Cells(cnt + 18, excelcol).Interior.Color = colorRed
            Else

                xlWorkSheet.Cells(cnt + 18, excelcol) = "Yes"

                xlWorkSheet.Cells(cnt + 18, excelcol).AddComment(mTCReviewObj.m2mSource)

            End If

            '19
            xlWorkSheet.Cells(cnt + 19, excelcol) = mTCReviewObj.modifiedDate

            excelcol = excelcol + 1

            '5th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).Interior.Color = System.Drawing.Color.FromArgb(165, 165, 165)  '9th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(cnt, excelcol).font.Bold = True
            xlWorkSheet.Cells(cnt, excelcol) = "Issue Addressed?"

            '9th Sep 2024
            Dim borders As Borders = xlWorkSheet.Cells(cnt, excelcol).Borders
            borders.LineStyle = XlLineStyle.xlContinuous
            borders.Weight = 2D

            AddDropDownListToExcelColumn(xlWorkSheet, 22, excelcol)

            excelcol = excelcol + 1
        Next

    End Sub

    Private Sub SetAssemblyWorkSheet_MTC(ByRef xlWorkSheet As Worksheet, ByVal finalAssemblyListBEC As List(Of MTC_Assembly), ByVal finalAssemblyListDGS As List(Of MTC_Assembly), MTC_Author_Name As String)

        '4th Sep 2024
        'Author Info and time span
        'xlWorkSheet.Cells(1, 1).Interior.Color = System.Drawing.Color.Gray
        'xlWorkSheet.Cells(1, 1).font.Color = System.Drawing.Color.White
        xlWorkSheet.Cells(1, 1).font.Bold = True
        'xlWorkSheet.Cells(1, 1) = $"{MTC_Author_Name}_{System.DateTime.Now.ToString("MMM_d_yyyy_HH_mm")}"
        xlWorkSheet.Cells(1, 3) = $"{MTC_Author_Name} {System.DateTime.Now.ToString("(MMM_d_yyyy_HH_mm)")}"
        '4th Sep 2024
        'Dim cnt As Integer = 1
        Dim cnt As Integer = 2

        'Dim finalAssemblyList As List(Of MTC_Assembly) = New List(Of MTC_Assembly)()
        'finalAssemblyList.AddRange(mtcMtrModelObj.mtcAssemblyList_BEC)
        'finalAssemblyList.AddRange(mtcMtrModelObj.mtrAssemblyList_DGS)

        Dim excelcol As Integer = 4

        For Each mTCReviewObj As MTC_Assembly In finalAssemblyListBEC

            '4th Sep 2024
            ''0. Assembly Name
            'xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
            'xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
            'xlWorkSheet.Cells(1, excelcol).font.Bold = True
            'xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.assemblyName

            '0. Assembly Name
            xlWorkSheet.Cells(cnt, excelcol).Interior.Color = System.Drawing.Color.FromArgb(165, 165, 165)  '9th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(cnt, excelcol).font.Bold = True
            xlWorkSheet.Cells(cnt, excelcol) = mTCReviewObj.assemblyName

            '9th Sep 2024
            Dim borders1 As Borders = xlWorkSheet.Cells(cnt, excelcol).Borders
            borders1.LineStyle = XlLineStyle.xlContinuous
            borders1.Weight = 2D

            '4th Sep 2024
            'SetExcelColumnColorAndBorder(xlWorkSheet, 23, excelcol)
            SetExcelColumnColorAndBorderNewMTC(xlWorkSheet, 23, excelcol)

            '1. Eco Number

            xlWorkSheet.Cells(cnt + 1, excelcol) = mTCReviewObj.ecoNumber
            If mTCReviewObj.revisionLevel = "0" Then
                xlWorkSheet.Cells(cnt + 1, excelcol).AddComment($"SOW = {mTCReviewObj.ecoNumber}")
            Else
                xlWorkSheet.Cells(cnt + 1, excelcol).AddComment($"ECO = {mTCReviewObj.ecoNumber}")
            End If

            '2. Part Number

            xlWorkSheet.Cells(cnt + 2, excelcol) = mTCReviewObj.partNumber

            If mTCReviewObj.isPartFound = True Then

                xlWorkSheet.Cells(cnt + 2, excelcol).AddComment("Part Number:" + mTCReviewObj.assemblyName)

            Else
                xlWorkSheet.Cells(cnt + 2, excelcol).AddComment("Part Number:" + mTCReviewObj.assemblyName)
                xlWorkSheet.Cells(cnt + 2, excelcol).Interior.Color = colorRed
            End If

            '3. Revision Level
            xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.revisionLevel

            '4. Author
            xlWorkSheet.Cells(cnt + 4, excelcol) = mTCReviewObj.author

            '5. Project Name

            xlWorkSheet.Cells(cnt + 5, excelcol) = mTCReviewObj.projectNameExist
            Try
                If mTCReviewObj.projectNameExist = "Yes" Then

                    xlWorkSheet.Cells(cnt + 5, excelcol).AddComment("ProjectName:" + mTCReviewObj.projectName)
                Else

                    xlWorkSheet.Cells(cnt + 5, excelcol).AddComment("ProjectName:" + mTCReviewObj.projectName)
                    xlWorkSheet.Cells(cnt + 5, excelcol).Interior.Color = colorRed
                End If
            Catch ex As Exception

            End Try

            ''6. Revision Number Correct
            'xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.revisionLevel

            'If mTCReviewObj.revisionLevel = "No" Then
            '    xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            '    xlWorkSheet.Cells(cnt + 6, excelcol).Interior.Color = colorRed
            'Else
            '    xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)

            'End If

            '13th Sep 2024
            '6. Revision Number Correct
            xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.revisionNumberCorrect
            xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            If (mTCReviewObj.revisionNumberCorrect = "No") Then
                xlWorkSheet.Cells(cnt + 6, excelcol).Interior.Color = colorRed
            End If

            '7. Document Number correct
            xlWorkSheet.Cells(cnt + 7, excelcol) = mTCReviewObj.documentNumberCorrect

            If mTCReviewObj.documentNumberCorrect = "Yes" Then
                xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Document Number:" + mTCReviewObj.documentNumber)
            Else
                xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Document Number:" + mTCReviewObj.documentNumber)
                xlWorkSheet.Cells(cnt + 7, excelcol).Interior.Color = colorRed
            End If

            '8. Is the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name)
            xlWorkSheet.Cells(cnt + 8, excelcol) = mTCReviewObj.authorExists

            If mTCReviewObj.authorExists = "Yes" Then
                xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Author:" + mTCReviewObj.author)
            Else
                xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Author:" + mTCReviewObj.author)
                xlWorkSheet.Cells(cnt + 8, excelcol).Interior.Color = colorRed
            End If

            '9. Do all technically unused properties have a "dash" populated?
            If (mTCReviewObj.isDashPopulated = "Yes") Then
                xlWorkSheet.Cells(cnt + 9, excelcol) = mTCReviewObj.isDashPopulated
            Else
                xlWorkSheet.Cells(cnt + 9, excelcol) = mTCReviewObj.isDashPopulated
                xlWorkSheet.Cells(cnt + 9, excelcol).Interior.Color = colorRed
            End If

            '10. Does the Model Title MATCH the M2M Item Master (INV) Description Field?
            xlWorkSheet.Cells(cnt + 10, excelcol) = mTCReviewObj.isTitleMatch_ItemMaster
            If mTCReviewObj.isTitleMatch_ItemMaster = "Yes" Then

                xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.title)
            Else
                xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.title)
                xlWorkSheet.Cells(cnt + 10, excelcol).Interior.Color = colorRed
            End If

            '11. Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)         
            If mTCReviewObj.isUOMMatch_M2M = "Yes" Then
                '26th September 2024
                If Not mTCReviewObj.UomProperty = "" Then
                    xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.isUOMMatch_M2M
                    xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
                End If
            Else
                '26th September 2024
                If Not mTCReviewObj.UomProperty = "" Then
                    xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.isUOMMatch_M2M
                    xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
                    xlWorkSheet.Cells(cnt + 11, excelcol).Interior.Color = colorRed
                End If
            End If

            '12. Perform a Parts List Report. How many total BOM items are in this assembly/weldment? (Not qty of all parts, only line items that will show up on draft PL) *
            xlWorkSheet.Cells(cnt + 12, excelcol) = $"{mTCReviewObj.partListCount}"

            '13. Is interfernces found in assembly?
            xlWorkSheet.Cells(cnt + 13, excelcol) = mTCReviewObj.isInterferenceFound
            If (mTCReviewObj.isInterferenceFound = "Yes") Then
                xlWorkSheet.Cells(cnt + 13, excelcol).Interior.Color = colorRed
            End If

            '14. inter-part copies detected
            xlWorkSheet.Cells(cnt + 14, excelcol) = mTCReviewObj.isInterPartCopiesDetected
            If (mTCReviewObj.isInterPartCopiesDetected = "Yes") Then
                xlWorkSheet.Cells(cnt + 14, excelcol).Interior.Color = colorRed
            End If

            '15. part copies detected
            xlWorkSheet.Cells(cnt + 15, excelcol) = mTCReviewObj.isPartCopiesDetected
            If (mTCReviewObj.isPartCopiesDetected = "Yes") Then
                xlWorkSheet.Cells(cnt + 15, excelcol).Interior.Color = colorRed
            End If

            '16. broken  file Path detected
            xlWorkSheet.Cells(cnt + 16, excelcol) = mTCReviewObj.isBrokenFilePathDetected
            If (mTCReviewObj.isBrokenFilePathDetected = "Yes") Then
                xlWorkSheet.Cells(cnt + 16, excelcol).Interior.Color = colorRed
            End If
            '17. modified date
            xlWorkSheet.Cells(cnt + 17, excelcol) = mTCReviewObj.modifiedDate

            '18. isadjustable
            xlWorkSheet.Cells(cnt + 18, excelcol) = mTCReviewObj.isAdjustable

            excelcol = excelcol + 1

            '5th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).Interior.Color = System.Drawing.Color.FromArgb(165, 165, 165)  '9th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(cnt, excelcol).font.Bold = True
            xlWorkSheet.Cells(cnt, excelcol) = "Issue Addressed?"

            '9th Sep 2024
            Dim borders As Borders = xlWorkSheet.Cells(cnt, excelcol).Borders
            borders.LineStyle = XlLineStyle.xlContinuous
            borders.Weight = 2D

            AddDropDownListToExcelColumn(xlWorkSheet, 23, excelcol)

            excelcol = excelcol + 1
        Next

        '--------------------------------------------------------
        'temp 20Feb2024

        For Each mTCReviewObj As MTC_Assembly In finalAssemblyListDGS

            '4th Sep 2024
            ''0. Assembly Name
            'xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
            'xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
            'xlWorkSheet.Cells(1, excelcol).font.Bold = True
            'xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.assemblyName

            '0. Assembly Name
            xlWorkSheet.Cells(cnt, excelcol).Interior.Color = System.Drawing.Color.FromArgb(165, 165, 165)  '9th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(cnt, excelcol).font.Bold = True
            xlWorkSheet.Cells(cnt, excelcol) = mTCReviewObj.assemblyName

            '9th Sep 2024
            Dim borders1 As Borders = xlWorkSheet.Cells(cnt, excelcol).Borders
            borders1.LineStyle = XlLineStyle.xlContinuous
            borders1.Weight = 2D

            '4th Sep 2024
            'SetExcelColumnColorAndBorder(xlWorkSheet, 23, excelcol)
            SetExcelColumnColorAndBorderNewMTC(xlWorkSheet, 23, excelcol)

            '1. Eco Number

            xlWorkSheet.Cells(cnt + 1, excelcol) = mTCReviewObj.ecoNumber
            If mTCReviewObj.revisionLevel = "0" Then
                xlWorkSheet.Cells(cnt + 1, excelcol).AddComment($"SOW = {mTCReviewObj.ecoNumber}")
            Else
                xlWorkSheet.Cells(cnt + 1, excelcol).AddComment($"ECO = {mTCReviewObj.ecoNumber}")
            End If

            '2. Part Number

            xlWorkSheet.Cells(cnt + 2, excelcol) = mTCReviewObj.partNumber

            If mTCReviewObj.isPartFound = True Then

                xlWorkSheet.Cells(cnt + 2, excelcol).AddComment("Part Number:" + mTCReviewObj.assemblyName)

            Else
                xlWorkSheet.Cells(cnt + 2, excelcol).AddComment("Part Number:" + mTCReviewObj.assemblyName)
                xlWorkSheet.Cells(cnt + 2, excelcol).Interior.Color = colorRed
            End If

            '3. Revision Level
            xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.revisionLevel

            '4. Author
            xlWorkSheet.Cells(cnt + 4, excelcol) = mTCReviewObj.author

            '5. Project Name

            xlWorkSheet.Cells(cnt + 5, excelcol) = mTCReviewObj.projectNameExist
            Try
                If mTCReviewObj.projectNameExist = "Yes" Then

                    xlWorkSheet.Cells(cnt + 5, excelcol).AddComment("ProjectName:" + mTCReviewObj.projectName)
                Else

                    xlWorkSheet.Cells(cnt + 5, excelcol).AddComment("ProjectName:" + mTCReviewObj.projectName)
                    xlWorkSheet.Cells(cnt + 5, excelcol).Interior.Color = colorRed
                End If
            Catch ex As Exception

            End Try

            ''6. Revision Number Correct
            'xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.revisionLevel

            'If mTCReviewObj.revisionLevel = "No" Then
            '    xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            '    xlWorkSheet.Cells(cnt + 6, excelcol).Interior.Color = colorRed
            'Else
            '    xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)

            'End If


            '13th Sep 2024
            '6. Revision Number Correct
            xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.revisionNumberCorrect
            xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revisionLevel)
            If (mTCReviewObj.revisionNumberCorrect = "No") Then
                xlWorkSheet.Cells(cnt + 6, excelcol).Interior.Color = colorRed
            End If

            '7. Document Number correct
            xlWorkSheet.Cells(cnt + 7, excelcol) = mTCReviewObj.documentNumberCorrect

            If mTCReviewObj.documentNumberCorrect = "Yes" Then
                xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Document Number:" + mTCReviewObj.documentNumber)
            Else
                xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Document Number:" + mTCReviewObj.documentNumber)
                xlWorkSheet.Cells(cnt + 7, excelcol).Interior.Color = colorRed
            End If

            '8. Is the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name)
            xlWorkSheet.Cells(cnt + 8, excelcol) = mTCReviewObj.authorExists

            If mTCReviewObj.authorExists = "Yes" Then
                xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Author:" + mTCReviewObj.author)
            Else
                xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Author:" + mTCReviewObj.author)
                xlWorkSheet.Cells(cnt + 8, excelcol).Interior.Color = colorRed
            End If

            '9. Do all technically unused properties have a "dash" populated?
            If (mTCReviewObj.isDashPopulated = "Yes") Then
                xlWorkSheet.Cells(cnt + 9, excelcol) = mTCReviewObj.isDashPopulated
            Else
                xlWorkSheet.Cells(cnt + 9, excelcol) = mTCReviewObj.isDashPopulated
                xlWorkSheet.Cells(cnt + 9, excelcol).Interior.Color = colorRed
            End If

            '10. Does the Model Title MATCH the M2M Item Master (INV) Description Field?
            xlWorkSheet.Cells(cnt + 10, excelcol) = mTCReviewObj.isTitleMatch_ItemMaster
            If mTCReviewObj.isTitleMatch_ItemMaster = "Yes" Then

                xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.title)
            Else
                xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.title)
                xlWorkSheet.Cells(cnt + 10, excelcol).Interior.Color = colorRed
            End If

            ''11. Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
            'xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.isUOMMatch_M2M
            'If mTCReviewObj.isUOMMatch_M2M = "Yes" Then
            '    xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
            'Else
            '    xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
            '    xlWorkSheet.Cells(cnt + 11, excelcol).Interior.Color = colorRed
            'End If

            '11. Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)         
            If mTCReviewObj.isUOMMatch_M2M = "Yes" Then
                '26th September 2024
                If Not mTCReviewObj.UomProperty = "" Then
                    xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.isUOMMatch_M2M
                    xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
                End If
            Else
                '26th September 2024
                If Not mTCReviewObj.UomProperty = "" Then
                    xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.isUOMMatch_M2M
                    xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
                    xlWorkSheet.Cells(cnt + 11, excelcol).Interior.Color = colorRed
                End If
            End If

            '12. Perform a Parts List Report. How many total BOM items are in this assembly/weldment? (Not qty of all parts, only line items that will show up on draft PL) *
            xlWorkSheet.Cells(cnt + 12, excelcol) = $"{mTCReviewObj.partListCount}"

            '13. Is interfernces found in assembly?
            xlWorkSheet.Cells(cnt + 13, excelcol) = mTCReviewObj.isInterferenceFound
            If (mTCReviewObj.isInterferenceFound = "Yes") Then
                xlWorkSheet.Cells(cnt + 13, excelcol).Interior.Color = colorRed
            End If

            '14. inter-part copies detected
            xlWorkSheet.Cells(cnt + 14, excelcol) = mTCReviewObj.isInterPartCopiesDetected
            If (mTCReviewObj.isInterPartCopiesDetected = "Yes") Then
                xlWorkSheet.Cells(cnt + 14, excelcol).Interior.Color = colorRed
            End If

            '15. part copies detected
            xlWorkSheet.Cells(cnt + 15, excelcol) = mTCReviewObj.isPartCopiesDetected
            If (mTCReviewObj.isPartCopiesDetected = "Yes") Then
                xlWorkSheet.Cells(cnt + 15, excelcol).Interior.Color = colorRed
            End If

            '16. broken  file Path detected
            xlWorkSheet.Cells(cnt + 16, excelcol) = mTCReviewObj.isBrokenFilePathDetected
            If (mTCReviewObj.isBrokenFilePathDetected = "Yes") Then
                xlWorkSheet.Cells(cnt + 16, excelcol).Interior.Color = colorRed
            End If
            '17. modified date
            xlWorkSheet.Cells(cnt + 17, excelcol) = mTCReviewObj.modifiedDate

            '18. isadjustable
            xlWorkSheet.Cells(cnt + 18, excelcol) = mTCReviewObj.isAdjustable

            excelcol = excelcol + 1

            '5th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).Interior.Color = System.Drawing.Color.FromArgb(165, 165, 165)  '9th Sep 2024
            xlWorkSheet.Cells(cnt, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(cnt, excelcol).font.Bold = True
            xlWorkSheet.Cells(cnt, excelcol) = "Issue Addressed?"

            '9th Sep 2024
            Dim borders As Borders = xlWorkSheet.Cells(cnt, excelcol).Borders
            borders.LineStyle = XlLineStyle.xlContinuous
            borders.Weight = 2D

            AddDropDownListToExcelColumn(xlWorkSheet, 23, excelcol)

            excelcol = excelcol + 1

        Next

        ''4th Sep 2024
        'Dim lastUsedColumn As Integer = excelcol - 1

        '' Merge cells in the first row for all the used columns
        'If lastUsedColumn >= 4 Then
        '    xlWorkSheet.Range(xlWorkSheet.Cells(1, 4), xlWorkSheet.Cells(1, lastUsedColumn)).Merge()
        'End If

        '' Center align the merged cell
        'xlWorkSheet.Cells(1, 4).HorizontalAlignment = XlHAlign.xlHAlignCenter
        'xlWorkSheet.Cells(1, 4).VerticalAlignment = XlVAlign.xlVAlignCenter
    End Sub

    Private Sub SetHorizontalAlignment(ByRef xlWorkSheet As Worksheet)
        Try
            xlWorkSheet.Columns.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SetExcelColumnColorAndBorder(ByRef xlWorkSheet As Worksheet, ByVal maxRowCnt As Integer, ByVal excelcol As Integer)
        Try

            If Not CLng(excelcol) > 3 Then
                Exit Sub
            End If
            Dim excelRange1 As Range = xlWorkSheet.Cells(1, excelcol)
            excelRange1.Cells.NumberFormat = "@"

#Region "Set excel column colour and set border"

            Dim rowCount As Integer = 2
            For index = 1 To maxRowCnt

                Try
                    Dim isOdd As Boolean = False
                    If CLng(excelcol) Mod 2 > 0 Then
                        isOdd = True
                    End If
                    Dim excelRange As Range = xlWorkSheet.Cells(rowCount, excelcol)
                    If isOdd Then
                        excelRange.Interior.Color = color1
                    Else
                        excelRange.Interior.Color = color2
                    End If

                    Dim borders As Borders = excelRange.Borders
                    borders.LineStyle = XlLineStyle.xlContinuous
                    borders.Weight = 2D

                    ' Dim excelRange1 As Range = xlCells3.Range($"{startCol}{rowCount.ToString()}", $"G{rowCount.ToString()}")
                    excelRange.Cells.NumberFormat = "@"

                Catch ex As Exception
                End Try
                rowCount += 1
            Next

#End Region

        Catch ex As Exception

        End Try
    End Sub

    '5th Sep 2024
    Private Sub SetExcelColumnColorAndBorderNewMTC(ByRef xlWorkSheet As Worksheet, ByVal maxRowCnt As Integer, ByVal excelcol As Integer)
        Try

            If Not CLng(excelcol) > 3 Then
                Exit Sub
            End If
            '4th Sep 2024
            'Dim excelRange1 As Range = xlWorkSheet.Cells(1, excelcol)
            Dim excelRange1 As Range = xlWorkSheet.Cells(2, excelcol)
            excelRange1.Cells.NumberFormat = "@"

#Region "Set excel column colour and set border"

            '4th Sep 2024
            'Dim rowCount As Integer = 2

            Dim rowCount As Integer = 3

            For index = 1 To maxRowCnt

                Try
                    Dim isOdd As Boolean = False
                    If CLng(excelcol) Mod 2 > 0 Then
                        isOdd = True
                    End If
                    Dim excelRange As Range = xlWorkSheet.Cells(rowCount, excelcol)
                    If isOdd Then
                        excelRange.Interior.Color = color1
                    Else
                        excelRange.Interior.Color = color2
                    End If

                    Dim borders As Borders = excelRange.Borders
                    borders.LineStyle = XlLineStyle.xlContinuous
                    borders.Weight = 2D

                    ' Dim excelRange1 As Range = xlCells3.Range($"{startCol}{rowCount.ToString()}", $"G{rowCount.ToString()}")
                    excelRange.Cells.NumberFormat = "@"

                Catch ex As Exception
                End Try
                rowCount += 1
            Next

#End Region

        Catch ex As Exception

        End Try
    End Sub

    '5th Sep 2024
    Private Sub AddDropDownListToExcelColumn(ByRef xlWorkSheet As Worksheet, ByVal maxRowCnt As Integer, ByVal excelcol As Integer)
        Try
            Dim rowCount As Integer = 3

            For index = 1 To maxRowCnt
                Try
                    Dim excelRange As Range = xlWorkSheet.Cells(rowCount, excelcol)

                    '9th Sep 2024
                    excelRange.Interior.Color = color2
                    Dim borders As Borders = excelRange.Borders
                    borders.LineStyle = XlLineStyle.xlContinuous
                    borders.Weight = 2D

                    ' Add a drop-down list (data validation) with "YES, NO, Remark"
                    With excelRange.Validation
                        .Delete() ' Remove any existing validation
                        .Add(Type:=XlDVType.xlValidateList,
                         AlertStyle:=XlDVAlertStyle.xlValidAlertStop,
                         Operator:=XlFormatConditionOperator.xlBetween,
                         Formula1:="YES,NO,Remark")
                        .IgnoreBlank = True
                        .InCellDropdown = True
                    End With

                Catch ex As Exception
                    MsgBox("Error to add drop-down list: " + ex.Message + vbNewLine + ex.StackTrace)
                End Try

                rowCount += 1
            Next
        Catch ex As Exception
            MsgBox("Error to add drop-down list: " + ex.Message + vbNewLine + ex.StackTrace)
        End Try
    End Sub

    Private Sub SetSheetMetalWorkSheet_MTR(ByRef xlWorkSheet As Worksheet, ByVal finalPartList As List(Of MTR_SheetMetal))

        Dim cnt As Integer = 1

        Dim excelcol As Integer = 4

        For Each mTCReviewObj As MTR_SheetMetal In finalPartList

            '0. Assembly Name
            xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
            xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(1, excelcol).font.Bold = True
            xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.assemblyName

            SetExcelColumnColorAndBorder(xlWorkSheet, 26, excelcol)

            If mTCReviewObj.isValidPart = False Then
                Dim excelRange As Range = xlWorkSheet.Cells(1, excelcol)
                excelRange.Interior.Color = Color.Red
                xlWorkSheet.Cells(1, excelcol).AddComment($"Invalid sheetmetal part.")
            End If

            '1.

            xlWorkSheet.Cells(cnt + 1, excelcol) = mTCReviewObj.isFeatureFullyConstrained

            '2.
            xlWorkSheet.Cells(cnt + 2, excelcol) = mTCReviewObj.verifySuppressFeature

            '3.
            xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.isAdjustable

            '4.
            xlWorkSheet.Cells(cnt + 4, excelcol) = mTCReviewObj.isInterPartCopiesDetected

            '5.
            xlWorkSheet.Cells(cnt + 5, excelcol) = mTCReviewObj.isPartCopiesDetected

            '6.
            xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.isValidAllCategories

            '7.
            xlWorkSheet.Cells(cnt + 7, excelcol) = mTCReviewObj.verifyWeightMass

            '8.
            xlWorkSheet.Cells(cnt + 8, excelcol) = mTCReviewObj.verifyUpdateOnFileSave

            '9.
            xlWorkSheet.Cells(cnt + 9, excelcol) = mTCReviewObj.isGeometryBroken

            '10.
            xlWorkSheet.Cells(cnt + 10, excelcol) = mTCReviewObj.author

            '11.
            xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.modifiedDate

            excelcol = excelcol + 1
        Next

    End Sub

    Private Sub SetPartWorkSheet_MTR(ByRef xlWorkSheet As Worksheet, ByVal finalPartList As List(Of MTR_Part))

        Dim cnt As Integer = 1

        Dim excelcol As Integer = 4

        For Each mTCReviewObj As MTR_Part In finalPartList

            '0. Assembly Name
            xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
            xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(1, excelcol).font.Bold = True
            xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.assemblyName

            SetExcelColumnColorAndBorder(xlWorkSheet, 18, excelcol)

            '1.

            xlWorkSheet.Cells(cnt + 1, excelcol) = mTCReviewObj.isFeatureFullyConstrained

            '2.
            xlWorkSheet.Cells(cnt + 2, excelcol) = mTCReviewObj.verifySuppressFeature

            '3.
            xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.isAdjustable

            '4.
            xlWorkSheet.Cells(cnt + 4, excelcol) = mTCReviewObj.isValidAllCategories

            '5.
            xlWorkSheet.Cells(cnt + 5, excelcol) = mTCReviewObj.verifyFastenerHardwarePart

            '6.
            xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.verifyWeightMass

            '7.
            xlWorkSheet.Cells(cnt + 7, excelcol) = mTCReviewObj.verifyUpdateOnFileSave

            '8.
            xlWorkSheet.Cells(cnt + 8, excelcol) = mTCReviewObj.author

            '9.
            xlWorkSheet.Cells(cnt + 9, excelcol) = mTCReviewObj.modifiedDate

            excelcol = excelcol + 1
        Next

    End Sub

    Private Sub SetAssemblyWorkSheet_MTR(ByRef xlWorkSheet As Worksheet, ByVal finalAssemblyList As List(Of MTR_Assembly))

        Dim cnt As Integer = 1
        Dim excelcol As Integer = 4

        For Each mTCReviewObj As MTR_Assembly In finalAssemblyList

            '0. Assembly Name
            xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
            xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
            xlWorkSheet.Cells(1, excelcol).font.Bold = True
            xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.assemblyName

            SetExcelColumnColorAndBorder(xlWorkSheet, 27, excelcol)

            '1.

            xlWorkSheet.Cells(cnt + 1, excelcol) = mTCReviewObj.isAdjustable

            '2.

            xlWorkSheet.Cells(cnt + 2, excelcol) = mTCReviewObj.isInterPartCopiesDetected

            '3.
            xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.isPartCopiesDetected

            '4.
            xlWorkSheet.Cells(cnt + 4, excelcol) = mTCReviewObj.isValidAllCategories

            '5.

            xlWorkSheet.Cells(cnt + 5, excelcol) = mTCReviewObj.isAssemblyFeatureExist

            '6.
            xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.isMatingPartInterferenceChecked

            '7.
            xlWorkSheet.Cells(cnt + 7, excelcol) = mTCReviewObj.verifyInterference

            '8.
            xlWorkSheet.Cells(cnt + 8, excelcol) = mTCReviewObj.verifyUpdateOnFileSave

            '9.
            xlWorkSheet.Cells(cnt + 9, excelcol) = mTCReviewObj.isGeometryBroken

            '10.
            xlWorkSheet.Cells(cnt + 10, excelcol) = mTCReviewObj.author

            '11.
            xlWorkSheet.Cells(cnt + 11, excelcol) = mTCReviewObj.modifiedDate

            excelcol = excelcol + 1
        Next

    End Sub

    Public Sub SaveAsMTRExcel(ByVal mtcMtrModelObj As MTC_MTR_Model, ByVal mtcExcelType As String)

        Dim excelAsmCol As Integer = 4
        Dim excelPartCol As Integer = 4
        Dim excelSheetMetalCol As Integer = 4
        Dim excelBaseLineCol As Integer = 4
        Dim excelElectricalCol As Integer = 4

        Dim asmExcelRowCnt As Integer = 25
        Dim partExcelRowCnt As Integer = 17
        Dim sheetMetalExcelRowCnt As Integer = 25


        Dim xlApp As Application
        Dim xlWorkBook As Workbook

        Dim xlWorkSheetAssembly As Worksheet
        Dim xlWorkSheetPart As Worksheet
        Dim xlWorkSheetSheetMetal As Worksheet

        xlApp = New Application
        'Dim aPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location
        'Dim dirPath As String = IO.Path.GetDirectoryName(aPath)
        'Dim mtcExcelPath As String = IO.Path.Combine(dirPath, $"MTR_BEC.xlsx")
        Dim mtrExcelPath As String = Config.configObj.MTRExcelPath

        xlWorkBook = xlApp.Workbooks.Open(mtrExcelPath)

        xlWorkSheetAssembly = xlWorkBook.Worksheets("Assembly")
        xlWorkSheetPart = xlWorkBook.Worksheets("Part")
        xlWorkSheetSheetMetal = xlWorkBook.Worksheets("Sheetmetal")

        SetHorizontalAlignment(xlWorkSheetAssembly)
        SetHorizontalAlignment(xlWorkSheetPart)
        SetHorizontalAlignment(xlWorkSheetSheetMetal)

        If mtcExcelType = "BEC" Then
            SetAssemblyWorkSheet_MTR(xlWorkSheetAssembly, mtcMtrModelObj.mtrAssemblyList_BEC)
            SetPartWorkSheet_MTR(xlWorkSheetPart, mtcMtrModelObj.mtrPartList_BEC)
            SetSheetMetalWorkSheet_MTR(xlWorkSheetSheetMetal, mtcMtrModelObj.mtrSheetMetalList_BEC)
        Else
            SetAssemblyWorkSheet_MTR(xlWorkSheetAssembly, mtcMtrModelObj.mtrAssemblyList_DGS)
            SetPartWorkSheet_MTR(xlWorkSheetPart, mtcMtrModelObj.mtrPartList_DGS)
            SetSheetMetalWorkSheet_MTR(xlWorkSheetSheetMetal, mtcMtrModelObj.mtrSheetMetalList_DGS)
        End If

        xlWorkSheetAssembly.Columns.AutoFit()
        xlWorkSheetPart.Columns.AutoFit()
        xlWorkSheetSheetMetal.Columns.AutoFit()

        Dim Asmname As String = IO.Path.GetFileNameWithoutExtension(mtcMtrModelObj.assemblyPath)
        Dim fileName As String = $"{Asmname}_MTR_Report_{mtcExcelType}_ {System.DateTime.Now.ToString("MMM_d_yyyy_HH_mm")}"
        Dim newname As String = IO.Path.Combine(mtcMtrModelObj.exportDirectoryLocation, $"{fileName}.xlsx")
        xlWorkBook.SaveAs(newname)
        xlWorkBook.Close()
        xlApp.Quit()

    End Sub

    Private Sub SetAssemblyData_RoutingSequence(ByRef xlWorkSheet As Worksheet, ByVal finalPartList As List(Of RoutingSequence_Assembly))

        Dim cnt As Integer = 1

        Dim excelRow As Integer = 2

        For Each mTCReviewObj As RoutingSequence_Assembly In finalPartList


            Dim excelRange1 As Range = xlWorkSheet.Cells.Range($"A{excelRow.ToString()}", $"U{excelRow.ToString()}")
            'excelRange1.Cells.NumberFormat = "@"

            Dim excelRange2 As Range = xlWorkSheet.Cells(excelRow, 1)
            excelRange2.Cells.NumberFormat = "@"

            xlWorkSheet.Cells(excelRow, 1) = mTCReviewObj.partNumber

            xlWorkSheet.Cells(excelRow, 2) = mTCReviewObj.massItem

            xlWorkSheet.Cells(excelRow, 3) = mTCReviewObj.m2mFSource

            xlWorkSheet.Cells(excelRow, 4) = mTCReviewObj.pmi

            xlWorkSheet.Cells(excelRow, 5) = mTCReviewObj.projectName

            xlWorkSheet.Cells(excelRow, 6) = mTCReviewObj.title

            xlWorkSheet.Cells(excelRow, 7) = mTCReviewObj.filePath

            xlWorkSheet.Cells(excelRow, 8) = mTCReviewObj.lastAuthor

            xlWorkSheet.Cells(excelRow, 9) = mTCReviewObj.floc

            xlWorkSheet.Cells(excelRow, 10) = mTCReviewObj.fbin

            xlWorkSheet.Cells(excelRow, 11) = mTCReviewObj.qAQC

            xlWorkSheet.Cells(excelRow, 12) = mTCReviewObj.quantity

            excelRow = excelRow + 1
        Next

    End Sub

    Private Sub SetStructureData_RoutingSequence(ByRef xlWorkSheet As Worksheet, ByVal finalPartList As List(Of RoutingSequence_Structure))

        Dim cnt As Integer = 1

        Dim excelRow As Integer = 2

        For Each mTCReviewObj As RoutingSequence_Structure In finalPartList


            Dim excelRange1 As Range = xlWorkSheet.Cells.Range($"A{excelRow.ToString()}", $"U{excelRow.ToString()}")
            'excelRange1.Cells.NumberFormat = "@"

            Dim excelRange2 As Range = xlWorkSheet.Cells(excelRow, 1)
            excelRange2.Cells.NumberFormat = "@"

            xlWorkSheet.Cells(excelRow, 1) = mTCReviewObj.partNumber

            xlWorkSheet.Cells(excelRow, 2) = mTCReviewObj.material

            xlWorkSheet.Cells(excelRow, 3) = mTCReviewObj.materialSpec

            xlWorkSheet.Cells(excelRow, 4) = mTCReviewObj.materialUsed

            xlWorkSheet.Cells(excelRow, 5) = mTCReviewObj.massItem

            xlWorkSheet.Cells(excelRow, 6) = mTCReviewObj.holeFeature.ToUpper()

            xlWorkSheet.Cells(excelRow, 7) = mTCReviewObj.holeFit

            xlWorkSheet.Cells(excelRow, 8) = mTCReviewObj.holeQty

            xlWorkSheet.Cells(excelRow, 9) = mTCReviewObj.m2mfSource

            xlWorkSheet.Cells(excelRow, 10) = mTCReviewObj.PMI

            xlWorkSheet.Cells(excelRow, 11) = mTCReviewObj.projectName

            xlWorkSheet.Cells(excelRow, 12) = mTCReviewObj.materialDescription

            xlWorkSheet.Cells(excelRow, 13) = mTCReviewObj.filePath

            xlWorkSheet.Cells(excelRow, 14) = mTCReviewObj.m2mflocation

            xlWorkSheet.Cells(excelRow, 15) = mTCReviewObj.m2mFbin

            xlWorkSheet.Cells(excelRow, 17) = mTCReviewObj.quantity

            xlWorkSheet.Cells(excelRow, 18) = mTCReviewObj.category

            excelRow = excelRow + 1
        Next

    End Sub

    Private Sub SetSheetMetalWorkSheetData_RoutingSequence(ByRef xlWorkSheet As Worksheet, ByVal finalPartList As List(Of RoutingSequence_SheetMetal))

        Dim cnt As Integer = 1

        'Dim excelcol As Integer = 4
        Dim excelRow As Integer = 2

        For Each mTCReviewObj As RoutingSequence_SheetMetal In finalPartList



            Dim excelRange1 As Range = xlWorkSheet.Cells.Range($"A{excelRow.ToString()}", $"U{excelRow.ToString()}")
            'excelRange1.Cells.NumberFormat = "@"

            Dim excelRange2 As Range = xlWorkSheet.Cells(excelRow, 1)
            excelRange2.Cells.NumberFormat = "@"

            xlWorkSheet.Cells(excelRow, 1) = mTCReviewObj.partNumber




            xlWorkSheet.Cells(excelRow, 2) = mTCReviewObj.material

            Dim materialThicknessWithoutUnit As String() = mTCReviewObj.materialThickness.Split(" ")
            If materialThicknessWithoutUnit.Count > 0 Then
                mTCReviewObj.materialThickness = materialThicknessWithoutUnit(0)
            End If
            xlWorkSheet.Cells(excelRow, 3) = mTCReviewObj.materialThickness


            'Dim materialSpecWithoutUnit As String() = mTCReviewObj.materialSpec.Split(" ")
            'If materialSpecWithoutUnit.Count > 0 Then
            '    mTCReviewObj.materialSpec = materialSpecWithoutUnit(0)
            'End If
            xlWorkSheet.Cells(excelRow, 4) = mTCReviewObj.materialSpec

            xlWorkSheet.Cells(excelRow, 5) = mTCReviewObj.materialUsed


            'Dim densitySpecWithoutUnit As String() = mTCReviewObj.density.Split(" ")
            'If densitySpecWithoutUnit.Count > 0 Then
            '    mTCReviewObj.density = densitySpecWithoutUnit(0)
            'End If
            'xlWorkSheet.Cells(excelRow, 6) = mTCReviewObj.density

            Dim MassSpecWithoutUnit As String() = mTCReviewObj.massItem.Split(" ")
            If MassSpecWithoutUnit.Count > 0 Then
                mTCReviewObj.density = MassSpecWithoutUnit(0)
            End If
            xlWorkSheet.Cells(excelRow, 6) = mTCReviewObj.density



            Dim bendRadiusWithoutUnit As String() = mTCReviewObj.bendRadius.Split(" ")
            If bendRadiusWithoutUnit.Count > 0 Then
                mTCReviewObj.bendRadius = bendRadiusWithoutUnit(0)
            End If
            xlWorkSheet.Cells(excelRow, 7) = mTCReviewObj.bendRadius



            Dim flatPatternXWithoutUnit As String() = mTCReviewObj.flat_Pattern_Model_CutSizeX.Split(" ")
            If flatPatternXWithoutUnit.Count > 0 Then
                mTCReviewObj.flat_Pattern_Model_CutSizeX = flatPatternXWithoutUnit(0)
            End If
            xlWorkSheet.Cells(excelRow, 8) = mTCReviewObj.flat_Pattern_Model_CutSizeX



            Dim flatPatternYWithoutUnit As String() = mTCReviewObj.flat_Pattern_Model_CutSizeY.Split(" ")
            If flatPatternYWithoutUnit.Count > 0 Then
                mTCReviewObj.flat_Pattern_Model_CutSizeY = flatPatternYWithoutUnit(0)
            End If
            xlWorkSheet.Cells(excelRow, 9) = mTCReviewObj.flat_Pattern_Model_CutSizeY



            xlWorkSheet.Cells(excelRow, 10) = mTCReviewObj.holeFeature.ToUpper()



            xlWorkSheet.Cells(excelRow, 11) = mTCReviewObj.holeFit

            xlWorkSheet.Cells(excelRow, 12) = mTCReviewObj.louverExists.ToUpper()

            xlWorkSheet.Cells(excelRow, 13) = mTCReviewObj.hem_Bead_GussetExists.ToUpper()

            xlWorkSheet.Cells(excelRow, 14) = mTCReviewObj.bendQty

            xlWorkSheet.Cells(excelRow, 15) = mTCReviewObj.holeQty

            xlWorkSheet.Cells(excelRow, 16) = mTCReviewObj.m2mfSource

            If mTCReviewObj.materialSpec.ToUpper.Contains("PERFORATED") Or mTCReviewObj.materialSpec.ToUpper.Contains("EXPAND") Then
                xlWorkSheet.Cells(excelRow, 17) = "TRUE"
            Else
                xlWorkSheet.Cells(excelRow, 17) = "FALSE"
            End If

            xlWorkSheet.Cells(excelRow, 19) = mTCReviewObj.projectName

            xlWorkSheet.Cells(excelRow, 20) = mTCReviewObj.materialDescription

            xlWorkSheet.Cells(excelRow, 21) = mTCReviewObj.filePath

            xlWorkSheet.Cells(excelRow, 22) = mTCReviewObj.quantity

            excelRow = excelRow + 1
        Next

    End Sub

    Private Sub SetAssembly_RoutingSequence(ByRef xlWorkSheet As Worksheet, ByVal finalPartList As List(Of RoutingSequence_Assembly))

        Dim cnt As Integer = 1
        Dim excelRow As Integer = 2

        For Each mTCReviewObj As RoutingSequence_Assembly In finalPartList

            Dim excelRange1 As Range = xlWorkSheet.Cells.Range($"A{excelRow.ToString()}", $"A{excelRow.ToString()}")
            'excelRange1.Cells.NumberFormat = "@"

            Dim excelRange2 As Range = xlWorkSheet.Cells(excelRow, 1)
            excelRange2.Cells.NumberFormat = "@"

            xlWorkSheet.Cells(excelRow, 1) = mTCReviewObj.partNumber

            If excelRow > 2 Then
                'xlWorkSheet.Cells(excelRow, 2) = xlWorkSheet.Cells(2, 2)
                With xlWorkSheet
                    .Range("B2:AK2").Copy()
                    .Range($"B{excelRow}:AK{excelRow}").PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas)
                End With
            End If


            excelRow = excelRow + 1
        Next

    End Sub

    Private Sub SetStructure_RoutingSequence(ByRef xlWorkSheet As Worksheet, ByVal finalPartList As List(Of RoutingSequence_Structure))

        Dim cnt As Integer = 1
        Dim excelRow As Integer = 2

        For Each mTCReviewObj As RoutingSequence_Structure In finalPartList

            Dim excelRange1 As Range = xlWorkSheet.Cells.Range($"A{excelRow.ToString()}", $"A{excelRow.ToString()}")
            excelRange1.Cells.NumberFormat = "@"

            Dim excelRange2 As Range = xlWorkSheet.Cells(excelRow, 1)
            excelRange2.Cells.NumberFormat = "@"
            xlWorkSheet.Cells(excelRow, 1) = mTCReviewObj.partNumber


            If excelRow > 2 Then
                'xlWorkSheet.Cells(excelRow, 2) = xlWorkSheet.Cells(2, 2)
                With xlWorkSheet
                    .Range("B2:AK2").Copy()
                    .Range($"B{excelRow}:AK{excelRow}").PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas)
                End With
            End If


            excelRow = excelRow + 1
        Next

    End Sub

    Private Sub SetSheetMetalWorkSheet_RoutingSequence(ByRef xlWorkSheet As Worksheet, ByVal finalPartList As List(Of RoutingSequence_SheetMetal))

        Dim cnt As Integer = 1
        Dim excelRow As Integer = 2

        For Each mTCReviewObj As RoutingSequence_SheetMetal In finalPartList

            Dim excelRange1 As Range = xlWorkSheet.Cells.Range($"A{excelRow.ToString()}", $"A{excelRow.ToString()}")
            excelRange1.Cells.NumberFormat = "@"


            Dim excelRange2 As Range = xlWorkSheet.Cells(excelRow, 1)
            excelRange2.Cells.NumberFormat = "@"
            xlWorkSheet.Cells(excelRow, 1) = mTCReviewObj.partNumber

            If excelRow > 2 Then
                'xlWorkSheet.Cells(excelRow, 2) = xlWorkSheet.Cells(2, 2)
                With xlWorkSheet
                    .Range("B2:AK2").Copy()
                    .Range($"B{excelRow}:AK{excelRow}").PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas)
                End With
            End If


            excelRow = excelRow + 1
        Next

    End Sub

    Public Sub SaveAsRoutingSequenceExcel(ByVal mtcMtrModelObj As MTC_MTR_Model)
        '9th Sep 2024 'Added try..catch
        Try
            Dim excelAsmCol As Integer = 4
            Dim excelPartCol As Integer = 4
            Dim excelSheetMetalCol As Integer = 4
            Dim excelBaseLineCol As Integer = 4
            Dim excelElectricalCol As Integer = 4

            Dim asmExcelRowCnt As Integer = 25
            Dim partExcelRowCnt As Integer = 17
            Dim sheetMetalExcelRowCnt As Integer = 25


            Dim xlApp As Application
            Dim xlWorkBook As Workbook

            'Dim xlWorkSheetAssembly As Worksheet
            'Dim xlWorkSheetPart As Worksheet
            Dim xlWorkSheetSheetMetalData As Worksheet
            Dim xlWorkSheetSheetMetal As Worksheet

            Dim xlWorkSheetStructureData As Worksheet
            Dim xlWorkSheetStructure As Worksheet

            Dim xlWorkSheetAssemblyData As Worksheet
            Dim xlWorkSheetAssembly As Worksheet

            Dim xlWorkSheetMiscPartData As Worksheet
            Dim xlWorkSheetMiscPart As Worksheet

            xlApp = New Application
            'Dim aPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location
            'Dim dirPath As String = IO.Path.GetDirectoryName(aPath)
            'Dim routingSeqExcelPath As String = IO.Path.Combine(dirPath, $"Routing_Sequence_Report.xlsx")
            Dim routingSeqExcelPath As String = Config.configObj.RoutingSequenceExcelPath

            xlWorkBook = xlApp.Workbooks.Open(routingSeqExcelPath)

            'xlWorkSheetAssembly = xlWorkBook.Worksheets("Assembly")
            'xlWorkSheetPart = xlWorkBook.Worksheets("Part")
            xlWorkSheetSheetMetalData = xlWorkBook.Worksheets("Sheetmetal Data")
            xlWorkSheetSheetMetal = xlWorkBook.Worksheets("SheetMetal")

            xlWorkSheetStructureData = xlWorkBook.Worksheets("Structure Data")
            xlWorkSheetStructure = xlWorkBook.Worksheets("Structure")

            xlWorkSheetAssemblyData = xlWorkBook.Worksheets("Assembly Data")
            xlWorkSheetAssembly = xlWorkBook.Worksheets("Assembly")

            xlWorkSheetMiscPartData = xlWorkBook.Worksheets("Misc. Parts Data")
            xlWorkSheetMiscPart = xlWorkBook.Worksheets("Misc. Parts")

            'SetHorizontalAlignment(xlWorkSheetAssembly)
            'SetHorizontalAlignment(xlWorkSheetPart)
            SetHorizontalAlignment(xlWorkSheetSheetMetalData)
            SetHorizontalAlignment(xlWorkSheetStructureData)
            SetHorizontalAlignment(xlWorkSheetAssemblyData)
            SetHorizontalAlignment(xlWorkSheetMiscPartData)

            'SetAssemblyWorkSheet_MTR(xlWorkSheetAssembly, mtcMtrModelObj.mtrAssemblyList_BEC)
            'SetPartWorkSheet_MTR(xlWorkSheetPart, mtcMtrModelObj.mtrPartList_BEC)

            'Sheet Metal
            Dim l As List(Of RoutingSequence_SheetMetal) = New List(Of RoutingSequence_SheetMetal)()
            l.AddRange(mtcMtrModelObj.routingSequenceSheetMetalList_BEC)
            l.AddRange(mtcMtrModelObj.routingSequenceSheetMetalList_DGS)

            SetSheetMetalWorkSheetData_RoutingSequence(xlWorkSheetSheetMetalData, l)

            SetSheetMetalWorkSheet_RoutingSequence(xlWorkSheetSheetMetal, l)

            xlWorkSheetSheetMetalData.Columns.AutoFit()

            'Structure
            Dim l1 As List(Of RoutingSequence_Structure) = New List(Of RoutingSequence_Structure)()
            l1.AddRange(mtcMtrModelObj.routingSequenceStructureList_BEC)
            l1.AddRange(mtcMtrModelObj.routingSequenceStructureList_DGS)

            SetStructureData_RoutingSequence(xlWorkSheetStructureData, l1)

            SetStructure_RoutingSequence(xlWorkSheetStructure, l1)

            xlWorkSheetStructureData.Columns.AutoFit()

            'Assembly
            Dim l2 As List(Of RoutingSequence_Assembly) = New List(Of RoutingSequence_Assembly)()
            l2.AddRange(mtcMtrModelObj.routingSequenceAssemblyList_BEC)
            l2.AddRange(mtcMtrModelObj.routingSequenceAssemblyList_DGS)

            SetAssemblyData_RoutingSequence(xlWorkSheetAssemblyData, l2)

            SetAssembly_RoutingSequence(xlWorkSheetAssembly, l2)

            xlWorkSheetStructureData.Columns.AutoFit()


            'Misc
            Dim l3 As List(Of RoutingSequence_Structure) = New List(Of RoutingSequence_Structure)()
            l3.AddRange(mtcMtrModelObj.routingSequenceMiscList_BEC)
            l3.AddRange(mtcMtrModelObj.routingSequenceMiscList_DGS)

            SetStructureData_RoutingSequence(xlWorkSheetMiscPartData, l3)

            SetStructure_RoutingSequence(xlWorkSheetMiscPart, l3)

            xlWorkSheetMiscPartData.Columns.AutoFit()


            Dim Asmname As String = IO.Path.GetFileNameWithoutExtension(mtcMtrModelObj.assemblyPath)
            Dim fileName As String = $"{Asmname}_RoutingSequence_Report_{System.DateTime.Now.ToString("MMM_d_yyyy_HH_mm")}"

            '2nd Sep 2024
            Dim newname As String = IO.Path.Combine(mtcMtrModelObj.export_Routing_Report_DirectoryLocation, $"{fileName}.xlsx")
            xlWorkBook.SaveAs(newname)
            xlWorkBook.Close()
            xlApp.Quit()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

End Class