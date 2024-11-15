Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports SolidEdgeDraft
Imports SolidEdgePart
Imports WK.Libraries.BetterFolderBrowserNS

Imports NLog
Imports NLog.Config
Imports SolidEdgeAssembly

Public Class MTC_MTR_ReviewForm
    Dim log As Logger = LogManager.GetCurrentClassLogger()

    Dim objApp As SolidEdgeFramework.Application = Nothing
    Dim objAssemblyDocument As SolidEdgeAssembly.AssemblyDocument = Nothing
    Dim dtAssemblyData As System.Data.DataTable
    Dim dtfilter As System.Data.DataTable
    Dim projectnamelst As List(Of String) = New List(Of String)
    Dim Authorlst As List(Of String) = New List(Of String)
    Dim dicData As Dictionary(Of String, DataSet) = New Dictionary(Of String, DataSet)()
    Dim objApplication As SolidEdgeFramework.Application = Nothing


    Dim dtM2M As System.Data.DataTable = New System.Data.DataTable("M2M")

    Dim excelcol As Integer = 3
    'Dim columnName As Char = "C"c
    Dim columnChar As Char = "Z"
    Dim color1 As System.Drawing.Color = Color.FromArgb(226, 239, 218)
    Dim color2 As System.Drawing.Color = Color.FromArgb(252, 228, 214)
    Private Sub btnExportExcel_Click(sender As Object, e As EventArgs) Handles btnExportExcel.Click

        Dim exportDirectoryLocation As String = GetExportDirectoryLocation()

        Dim dtData As System.Data.DataTable = DirectCast(dgvDocumentDetails.DataSource, System.Data.DataTable)

        waitStartSave()

        For Each dr As DataRow In dtData.Rows

            If dr("Select").ToString().ToUpper() = "TRUE" Then

                SaveAsExcel(dr, exportDirectoryLocation)

            End If

        Next
        WaitEndSave()
        MsgBox("Process completed.")
    End Sub

    Private Function GetExportDirectoryLocation() As String
        Dim exportDirectoryLocation As String = String.Empty
        Try
            Dim folderpath As String = ""
            folderpath = browseFolderAdvanced()
            If Not folderpath = String.Empty Then
                exportDirectoryLocation = folderpath
            End If
        Catch ex1 As Exception
            Try
                Dim MyFolderBrowser As Ookii.Dialogs.VistaFolderBrowserDialog = New Ookii.Dialogs.VistaFolderBrowserDialog
                Dim dlgResult As DialogResult = MyFolderBrowser.ShowDialog
                If dlgResult = DialogResult.OK Then
                    exportDirectoryLocation = MyFolderBrowser.SelectedPath
                End If
            Catch ex As Exception
                If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
                    Dim path As String = FolderBrowserDialog1.SelectedPath
                    exportDirectoryLocation = path
                End If
            End Try

        End Try
        Return exportDirectoryLocation
    End Function

    Public Shared Function browseFolderAdvanced() As String

        Dim folderpath As String = ""
        Try
            Dim BetterFolderBrowser As New BetterFolderBrowser()

            BetterFolderBrowser.Title = "Select folders"

            BetterFolderBrowser.RootFolder = "C:\\"

            BetterFolderBrowser.Multiselect = False
            If BetterFolderBrowser.ShowDialog() = DialogResult.OK Then
                folderpath = BetterFolderBrowser.SelectedFolder
            End If
        Catch ex As Exception
            MsgBox("Advanced Browse folder" + ex.Message + vbNewLine + ex.StackTrace)
        End Try


        Return folderpath
    End Function

    Public Sub SaveAsExcel(ByVal dr As DataRow, ByVal exportDirectoryLocation As String)


        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        xlApp = New Application
        Dim aPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location
        Dim dirPath As String = IO.Path.GetDirectoryName(aPath)
        Dim mtcExcelPath As String = IO.Path.Combine(dirPath, "MTC.xlsx")

        xlWorkBook = xlApp.Workbooks.Open(mtcExcelPath) '"C:\Users\milipatel\Downloads\MTC.xlsx")
        xlWorkSheet = xlWorkBook.Worksheets("MTC")

        '1,1 > SrNO
        '1,2 > MTC Quest
        '1,3 > Value
        '1,4 > Validation

        'Item Number
        'File Name(no extension)
        'Quantity
        'Flat_Pattern_Model_CutSizeX
        'Flat_Pattern_Model_CutSizeY
        'Material Used
        'Title
        'Comments
        'File Name(full path)
        'Document Number
        'Author
        'Revision Number
        'Component Type Order
        'Assembly Order

        Dim fileNameWithoutExt As String = String.Empty
        Dim revision As String = String.Empty
        Dim author As String = String.Empty
        Dim documentno As String = String.Empty
        Dim materialused As String = String.Empty
        Dim matlspec As String = String.Empty
        Dim fullpath As String = String.Empty
        Dim lastsaved As String = String.Empty
        Dim density As String = String.Empty
        Dim projectname As String = String.Empty
        Dim Isassembly As String = "No"
        Dim isadjustable As String = String.Empty
        Dim properties As String = "Yes"
        Dim customprop As String = String.Empty

        Try

            fileNameWithoutExt = dr("File Name (no extension)").ToString()
            revision = dr("Revision Number").ToString()
            author = dr("Author").ToString()
            documentno = dr("Document Number").ToString()
            materialused = dr("Material Used").ToString()
            matlspec = dr("MATL SPEC").ToString()
            fullpath = dr("File Name (full path)").ToString()
            lastsaved = dr("Last Saved Version").ToString()
            density = dr("Density")
            projectname = dr("Project")

        Catch ex As Exception
        End Try
        If fullpath.EndsWith(".asm") Then
            Isassembly = "Yes"

        End If



        If fullpath.EndsWith(".par") Or fullpath.EndsWith(".psm") Then
            Dim documents As SolidEdgeFramework.Documents = Nothing

            Dim objSheetMetalDocument As SolidEdgePart.SheetMetalDocument = Nothing
            Dim objPartDocument As SolidEdgePart.PartDocument = Nothing
            Dim adjustable As Boolean
            documents = objApp.Documents
            objSheetMetalDocument = DirectCast(documents.Open(fullpath), SolidEdgeFramework.SolidEdgeDocument)
            adjustable = objSheetMetalDocument.IsAdjustablePart

            If adjustable = "True" Then
                isadjustable = "Yes"
            ElseIf adjustable = "False" Then
                isadjustable = "No"
            End If


            'Dim propSets As SolidEdgeFramework.PropertySets = objSheetMetalDocument.Properties
            'Dim custProps As SolidEdgeFramework.Properties = propSets.Item("Custom")
            'Dim summary As SolidEdgeFramework.Properties = propSets.Item("SummaryInformation")

            'For Each prop1 As SolidEdgeFramework.[Property] In custProps
            '    If prop1.Value = "" Then
            '        properties = "No"
            '        objSheetMetalDocument.Close(SaveChanges:=False)
            '        Exit For
            '    End If

            'Next
            If properties = "Yes" Then


                'For Each prop1 As SolidEdgeFramework.[Property] In summary
                '    If prop1.Value = "" Then
                '        properties = "No"
                '        objSheetMetalDocument.Close(SaveChanges:=False)
                '        Exit For
                '    End If
                'Next

                '

            End If
            objSheetMetalDocument.Close(SaveChanges:=False)

        End If



        'edit the cell with new value
        xlWorkSheet.Cells(3, 3) = fileNameWithoutExt
        xlWorkSheet.Cells(4, 3) = revision
        xlWorkSheet.Cells(8, 3) = revision
        xlWorkSheet.Cells(7, 3) = projectname

        xlWorkSheet.Cells(5, 3) = author
        xlWorkSheet.Cells(10, 3) = author

        xlWorkSheet.Cells(9, 3) = documentno

        xlWorkSheet.Cells(22, 3) = matlspec
        xlWorkSheet.Cells(29, 3) = matlspec

        ' xlWorkSheet.Cells(11, 3) = properties
        xlWorkSheet.Cells(21, 3) = Isassembly
        xlWorkSheet.Cells(23, 3) = materialused
        xlWorkSheet.Cells(15, 3) = lastsaved

        If density Is Nothing Then
            xlWorkSheet.Cells(16, 3) = "No"
        Else
            xlWorkSheet.Cells(16, 3) = "Yes"
        End If

        xlWorkSheet.Cells(41, 3) = isadjustable


        Dim fileName As String = $"{fileNameWithoutExt}_MTC_Report"
        Dim newname As String = IO.Path.Combine(exportDirectoryLocation, $"{fileName}.xlsx") '  "C:\Users\milipatel\Downloads\MTC11111.xlsx"
        xlWorkBook.SaveAs(newname)
        xlWorkBook.Close()
        xlApp.Quit()

    End Sub


    Public Sub TestPhysicProp(ByVal theDocument As SolidEdgePart.SheetMetalDocument)
        Try
            Dim Status As Integer

            Dim Density, Accuracy, Volume, Area, Mass, RelativeAccuracyAchieved As Double

            Dim CenterOfGravity() As Double = New Double() {}

            Dim CenterOfVolume() As Double = New Double() {}

            Dim GlobalMomentsOfInertia() As Double = New Double() {}

            Dim PrincipalMomentsOfInertia() As Double = New Double() {}

            Dim PrincipalAxes() As Double = New Double() {}

            Dim RadiiOfGyration() As Double = New Double() {}


            Density = CDbl(theDocument.Variables.Item("PhysicalProperties_Density").Value)

            Accuracy = CDbl(theDocument.Variables.Item("PhysicalProperties_Accuracy").Value)

            Mass = 0


            If theDocument.PhysicalPropertiesStatus = PhysicalPropertiesStatusConstants.sePhysicalPropertiesStatus_User Then

            End If
            theDocument.Models.Item(1).ComputePhysicalProperties(Density, Accuracy, Volume, Area, Mass, CenterOfGravity,
                                                                 CenterOfVolume, GlobalMomentsOfInertia, PrincipalMomentsOfInertia, PrincipalAxes, RadiiOfGyration, RelativeAccuracyAchieved, Status)

            theDocument.Models.Item(1).Recompute()
            'Dim a As PhysicalPropertiesStatusConstants
            'Debug.Print("")
            'MsgBox("")
            'End If
        Catch ex As Exception
            MsgBox(ex.StackTrace)
        End Try
    End Sub

    Private Function CheckPartConstructionFeature(ByRef partDoc As PartDocument) As String
        Dim resultCheckPart As String = String.Empty
        Try

            Dim copyConstructionCount As Integer = partDoc.Constructions.CopyConstructions.Count
            If copyConstructionCount > 0 Then
                resultCheckPart = "No"
            Else
                resultCheckPart = "Yes"
            End If

        Catch ex As Exception
            resultCheckPart = ex.Message
        End Try
        Return resultCheckPart
    End Function

    Private Function IsGeomtryBroken(ByRef objSheetMetalDocument As SolidEdgePart.SheetMetalDocument, ByRef mTCReviewObj As MTCReview) As String

        Dim geoMetryBroken As String = "Yes"
        Try
            Dim ipl As SolidEdgeFramework.InterpartLinks = objSheetMetalDocument.InterpartLinks
            If ipl.Count > 1 Then
                geoMetryBroken = "No"
            End If
        Catch ex As Exception
        End Try
        Return geoMetryBroken

    End Function

    Private Function IsGeomtryBroken_assembly(ByRef objSheetMetalDocument As SolidEdgeAssembly.AssemblyDocument) As String

        Dim geoMetryBroken As String = "Yes"
        Try
            Dim ipl As SolidEdgeFramework.InterpartLinks = objSheetMetalDocument.InterpartLinks
            If ipl.Count > 1 Then
                geoMetryBroken = "No"
            End If
        Catch ex As Exception
        End Try
        Return geoMetryBroken

    End Function

    Private Function interpartcopycheck(ByVal partDoc As SolidEdgeFramework.SolidEdgeDocument, ByRef mTCReviewObj As MTCReview) As String
        Dim CheckPart As String = String.Empty
        Dim interpartcheck As String = String.Empty

        Dim result As String = String.Empty
        Try

            'part copies detected
            Dim copyConstructionCount As Integer = partDoc.Constructions.CopyConstructions.Count
            If copyConstructionCount > 0 Then
                CheckPart = "Yes"
                mTCReviewObj.partCopiesDetected = "Yes"
            Else
                CheckPart = "No"
                mTCReviewObj.partCopiesDetected = "No"
            End If

            'Inter part copies detected
            Dim interpartlinkscount As Integer = partDoc.InterpartLinks.count
            If interpartlinkscount > 0 Then
                interpartcheck = "Yes"
                mTCReviewObj.interPartCopiesDetected = "Yes"
            Else
                interpartcheck = "No"
                mTCReviewObj.interPartCopiesDetected = "No"
            End If




            If CheckPart = "Yes" And interpartcheck = "Yes" Then
                result = "Yes"
            Else
                result = "N0"
            End If

        Catch ex As Exception

        End Try
        Return result
    End Function

    Private Function CheckPartConstructionFeature2(ByRef partDoc As SheetMetalDocument) As String
        Dim resultCheckPart As String = String.Empty

        Try

            Dim copyConstructionCount As Integer = partDoc.Constructions.CopyConstructions.Count
            If copyConstructionCount > 0 Then
                resultCheckPart = "No"
            Else
                resultCheckPart = "Yes"
            End If

        Catch ex As Exception
            resultCheckPart = ex.Message
        End Try
        Return resultCheckPart
    End Function

    Private Function CheckAssemblyFeatureExistence(ByRef asssemblyDocObj As AssemblyDocument) As String

        Dim result As String = String.Empty
        Try

            Dim asmModelObj As SolidEdgePart.Model = asssemblyDocObj.AssemblyModel
            Dim asmModelFeatures As Features = asmModelObj.Features
            Dim cnt As Integer = asmModelFeatures.Count
            If cnt > 0 Then
                result = "Yes"
            Else
                result = "No"
            End If

        Catch ex As Exception
            result = ex.Message
        End Try
        Return result
    End Function

    Private Function getgagename(ByVal objSMDoc As SolidEdgePart.SheetMetalDocument) As Integer
        '   Dim objApplication As SolidEdgeFramework.Application = Nothing

        Dim myMatTable As SolidEdgeFramework.MatTable = Nothing
        Dim strCurrGageName As String = ""
        Dim strGageFilePath As String = ""
        Dim nMTLUsingExcel As Integer
        Dim strMTLGageTableName As String = ""
        Dim nDocUsingExcel As Integer = 0
        Dim strDocGageTableName As String = ""
        Dim nCountBendRadiusVals As Integer
        Dim nCountBendAngleVals As Integer
        Dim nCountNFVals As Integer


        Try


            myMatTable = objApplication.GetMaterialTable()


            If (objSMDoc Is Nothing) Then
                MsgBox("Failed to get Sheet Metal Document object.")
            End If
            Call myMatTable.GetPSMGaugeInfoForDoc(objSMDoc,
                                      strCurrGageName,
                                      strGageFilePath,
                                      nMTLUsingExcel,
                                      strMTLGageTableName,
                                      nDocUsingExcel,
                                      strDocGageTableName,
                                      nCountBendRadiusVals,
                                      nCountBendAngleVals,
                                      nCountNFVals)

        Catch ex As Exception

        End Try
        Return nDocUsingExcel

    End Function

    Public Sub SaveAsExcel2(ByVal dt As System.Data.DataTable, ByVal exportDirectoryLocation As String, ByVal excelName As String)

        Dim mtc As MTCReview = New MTCReview()
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        xlApp = New Application
        Dim aPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location
        Dim dirPath As String = IO.Path.GetDirectoryName(aPath)
        Dim mtcExcelPath As String = IO.Path.Combine(dirPath, $"MTC.xlsx")

        xlWorkBook = xlApp.Workbooks.Open(mtcExcelPath) '"C:\Users\milipatel\Downloads\MTC.xlsx")
        xlWorkSheet = xlWorkBook.Worksheets("MTC")
        xlWorkSheet.Columns.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

        For Each dr As DataRow In dt.Rows

            Dim itemnumber As String = String.Empty
            Dim fileNameWithoutExt As String = String.Empty
            Dim revision As String = String.Empty
            Dim author As String = String.Empty
            Dim documentno As String = String.Empty
            Dim materialused As String = String.Empty
            Dim matlspec As String = String.Empty
            Dim fullpath As String = String.Empty
            Dim lastsaved As String = String.Empty
            Dim density As String = String.Empty
            Dim projectname As String = String.Empty
            'Dim Isassembly As String = "No"
            Dim isadjustable As String = String.Empty
            Dim properties As String = "Yes"
            Dim customprop As String = String.Empty
            Dim columname As String = String.Empty
            Dim Documenttype As String = String.Empty
            Dim adjustable As Boolean
            Dim hardwarepart As Boolean
            Dim isflatpattern As String = String.Empty
            Dim iscutout As String = String.Empty

            Dim revisionlevel As String = String.Empty
            Dim authorcheck As String = String.Empty
            Dim UomProperty As String = String.Empty
            Dim Title As String = String.Empty
            Dim comments As String = String.Empty
            Dim category As String = String.Empty
            Dim keywords As String = String.Empty
            Dim SEfeatures As String = String.Empty
            Dim partlistcount As Integer = 0
            Dim interferencereport As String = String.Empty

            Dim checkPartFeature As String = String.Empty
            Dim checkAssemblyFeature As String = String.Empty
            Dim gageeexcelfile As String = String.Empty
            Dim issupress As String = String.Empty
            Dim sketchisfullydefined As String = String.Empty
            Dim allinterpartcopycheck As String = String.Empty

            Try


                '210-07293.psm:1
                If Not dr("Item Number") Is Nothing Then
                    itemnumber = dr("Item Number").ToString()
                End If

                If Not dr("File Name (no extension)") Is Nothing Then
                    fileNameWithoutExt = dr("File Name (no extension)").ToString()
                    If fileNameWithoutExt.Contains("_") Then
                        Dim myDelims As String() = New String() {"_"}
                        Dim splitfilename = fileNameWithoutExt.Split(myDelims, StringSplitOptions.None)
                        ' Dim splitfilename As String() = fileNameWithoutExt.Split("_")
                        fileNameWithoutExt = splitfilename(0)
                        revisionlevel = splitfilename(1)
                    End If
                    'If fileNameWithoutExt.Contains("_") Then
                    '    Dim splitfilename As String() = fileNameWithoutExt.Split("_")
                    '    fileNameWithoutExt = splitfilename(0)
                    '    revisionlevel = splitfilename(1)
                    'End If
                End If

                If Not dr("Revision Number") Is Nothing Then
                    revision = dr("Revision Number").ToString()
                End If

                If Not dr("Title") Is Nothing Then
                    Title = dr("Title").ToString()
                End If

                If Not dr("Author") Is Nothing Then
                    author = dr("Author").ToString()
                End If

                If Not dr("Document Number") Is Nothing Then
                    documentno = dr("Document Number").ToString()
                End If

                If Not dr("Comments") Is Nothing Then
                    comments = dr("Comments").ToString()
                End If

                If Not dr("Category") Is Nothing Then
                    category = dr("Category").ToString()
                End If

                If Not dr("Material Used") Is Nothing Then
                    materialused = dr("Material Used").ToString()
                End If

                If Not dr("MATL SPEC") Is Nothing Then
                    matlspec = dr("MATL SPEC").ToString()
                End If

                If Not dr("File Name (full path)") Is Nothing Then
                    fullpath = dr("File Name (full path)").ToString()
                End If

                If Not dr("Last Author") Is Nothing Then
                    lastsaved = dr("Last Author").ToString()
                End If

                If Not dr("Density") Is Nothing Then
                    density = dr("Density").ToString()
                End If

                If Not dr("Project") Is Nothing Then
                    projectname = dr("Project").ToString()
                End If

                If Not dr("Status Text") Is Nothing Then
                    Documenttype = dr("Status Text").ToString()
                End If

                If Not dr("UOM") Is Nothing Then
                    UomProperty = dr("UOM").ToString()
                End If

                If Not dr("Keywords") Is Nothing Then
                    keywords = dr("Keywords").ToString()
                End If

                Dim ispartfound As Boolean
                Dim dv As DataView = New DataView(dtM2M)
                Dim partno As String = ExcelUtil.ExcelMtcReview.fpartno.ToString()
                dv.RowFilter = $"{partno}='{fileNameWithoutExt}'"

                Dim M2Mdescript As String = String.Empty
                Dim M2Mmeasure As String = String.Empty
                Dim M2Msource As String = String.Empty
                Dim M2Mvendorname As String = String.Empty
                Dim M2Mflocation As String = String.Empty
                Dim M2MFbin As String = String.Empty
                'Dim M2Mkeywords As String
                Dim M2Mpartno As String = String.Empty


                For Each drv As DataRowView In dv

                    M2Mpartno = drv(ExcelUtil.ExcelMtcReview.fpartno.ToString)
                    M2Mdescript = drv(ExcelUtil.ExcelMtcReview.fdescript.ToString)
                    M2Mmeasure = drv(ExcelUtil.ExcelMtcReview.fmeasure.ToString)
                    M2Msource = drv(ExcelUtil.ExcelMtcReview.fsource.ToString)
                    M2Mvendorname = drv(ExcelUtil.ExcelMtcReview.VendorName.ToString)
                    M2Mflocation = drv(ExcelUtil.ExcelMtcReview.flocation.ToString)
                    M2MFbin = drv(ExcelUtil.ExcelMtcReview.Fbin.ToString)
                    '  M2Mkeywords = drv(ExcelUtil.excelMtcReview.fKeywords.ToString)

                Next

                Dim commentformate As String = "***" + M2Mvendorname + "=" + M2Mpartno + "***"
                If dv.Count = 0 Then
                    ispartfound = False
                Else
                    ispartfound = True
                End If


                Dim rowCount As Integer = 2

                For index = 1 To 50

                    Try

                        Dim isOdd As Boolean = False
                        If CLng(excelcol) Mod 2 > 0 Then
                            isOdd = True
                        End If

                        Dim excelRange As Range = xlWorkSheet.Cells(rowCount, excelcol)
                        ' Dim excelRange As Range = xlWorkSheet.Range($"{columnChar.ToString()}{rowCount.ToString()}", $"{columnChar.ToString()}{rowCount.ToString()}")
                        If isOdd Then
                            excelRange.Interior.Color = color1
                            '  xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
                        Else
                            excelRange.Interior.Color = color2
                        End If

                        Dim borders As Borders = excelRange.Borders
                        borders.LineStyle = XlLineStyle.xlContinuous
                        borders.Weight = 2D
                    Catch ex As Exception
                    End Try
                    rowCount += 1

                Next

                ' columnChar = Char.Parse(((CInt(columnChar)) + 1).ToString())
                ' columnChar = Chr(Asc(columnChar) + 1)

                If dr("Select").ToString().ToUpper() = "TRUE" Then


                    Dim documents As SolidEdgeFramework.Documents = objApp.Documents

                    If fullpath.EndsWith(".psm") Then

                        If itemnumber.Contains("*") Then
                            allinterpartcopycheck = "Yes"
                            xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
                            xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
                            xlWorkSheet.Cells(1, excelcol).font.Bold = True
                            xlWorkSheet.Cells(1, excelcol) = fileNameWithoutExt
                            xlWorkSheet.Cells(48, excelcol) = allinterpartcopycheck
                            xlWorkSheet.Cells(48, excelcol).AddComment("Document Link Broken")
                            Exit For

                        End If

                        Dim objSheetMetalDocument As SolidEdgePart.SheetMetalDocument = Nothing
                        objSheetMetalDocument = DirectCast(documents.Open(fullpath), SolidEdgeFramework.SolidEdgeDocument)
                        Threading.Thread.Sleep(2000)
                        'TestPhysicProp(objSheetMetalDocument)
                        ' Dim updateonfilesave As Boolean = objSheetMetalDocument.UpdateOnFileSave


                        'objApp.StartCommand(25038)

                        'objApp.StartCommand(45000)
                        'objSheetMetalDocument.Save()
                        'Exit Sub

                        'Check sketch fullydefined or not 
                        sketchisfullydefined = sketchdefined(objSheetMetalDocument)
                        Dim excelfilecount As Integer = getgagename(objSheetMetalDocument)

                        If excelfilecount = 0 Then
                            gageeexcelfile = "No"
                        ElseIf excelfilecount = 1 Then
                            gageeexcelfile = "Yes"
                        End If

                        Dim flatpatternmodel As SolidEdgePart.FlatPatternModels = objSheetMetalDocument.FlatPatternModels
                        If flatpatternmodel.Count = 0 Then
                            isflatpattern = "No"
                        Else
                            isflatpattern = "Yes"
                        End If




                        Dim models As SolidEdgePart.Models = objSheetMetalDocument.Models
                        Dim model As SolidEdgePart.Model = Nothing
                        objSheetMetalDocument.UpdateOnFileSave = False
                        objSheetMetalDocument.Models.Item(1).Recompute()


                        'Check supress 
                        Dim features As SolidEdgePart.Features = objSheetMetalDocument.Models.Item(1).Features

                        For i = 1 To features.Count

                            Dim obj As Object = features.Item(i)
                            Dim supressvariable As Boolean = obj.suppress
                            If supressvariable = True Then
                                issupress = "Yes"
                                Exit For
                            End If

                        Next
                        'models = objSheetMetalDocument.Models
                        iscutout = ""
                        Try
                            model = models.Item(1)
                            Try
                                If model.ExtrudedCutouts.Count > 0 Then
                                    Dim cutouts As SolidEdgePart.ExtrudedCutouts = model.ExtrudedCutouts
                                    For Each cutout As SolidEdgePart.ExtrudedCutout In cutouts
                                        Dim profile As Profile = cutout.Profile
                                        If profile.Circles2d.Count = "0" Then
                                            iscutout = "Yes"
                                        Else
                                            iscutout = "No"
                                        End If
                                        ' Dim circle As SolidEdgeFrameworkSupport.Circles2d = profile.Circles2d.co
                                    Next
                                ElseIf model.NormalCutouts.Count > 0 Then
                                    Dim cutouts As SolidEdgePart.NormalCutouts = model.NormalCutouts
                                    For Each cutout As SolidEdgePart.NormalCutout In cutouts
                                        'Try
                                        '    ' Dim noOfProfile As Object
                                        '    '  Dim prfArray As Object
                                        '    '  cutout.GetProfiles(noOfProfile, prfArray)
                                        'Catch ex As Exception

                                        'End Try

                                        Dim profile As Profile = cutout.Profile
                                        If profile.Circles2d.Count = "0" Then
                                            iscutout = "Yes"
                                        Else
                                            iscutout = "No"
                                        End If
                                        ' Dim circle As SolidEdgeFrameworkSupport.Circles2d = profile.Circles2d.co
                                    Next
                                End If
                            Catch ex As Exception

                            End Try


                        Catch ex As Exception
                        End Try


                        'Dim status As SolidEdgeAssembly.OccurrenceStatusConstants = Nothing
                        'Dim occur As SolidEdgeAssembly.Occurrence = objSheetMetalDocument

                        'Dim cutout As SolidEdgePart.ExtrudedCutout = cutouts.Item(1)
                        Try
                            adjustable = objSheetMetalDocument.IsAdjustablePart
                        Catch ex As Exception

                            Debug.Print("aaaa")
                        End Try

                        'Check SEFeatures
                        Try
                            Dim SEModels As SimplifiedModels = objSheetMetalDocument.SimplifiedModels
                            If SEModels.Count > 0 Then
                                SEfeatures = "Yes"
                            Else
                                SEfeatures = "No"
                            End If


                        Catch ex As Exception


                        End Try

                        'check all-interpartlink,copypart,geomtry broken

                        allinterpartcopycheck = interpartcopycheck(objSheetMetalDocument, mtc)

                        checkPartFeature = CheckPartConstructionFeature2(objSheetMetalDocument)
                        objSheetMetalDocument.Close(SaveChanges:=False)

                    ElseIf fullpath.EndsWith(".par") Then

                        If itemnumber.Contains("*") Then
                            allinterpartcopycheck = "No"
                            xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
                            xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
                            xlWorkSheet.Cells(1, excelcol).font.Bold = True
                            xlWorkSheet.Cells(1, excelcol) = fileNameWithoutExt
                            xlWorkSheet.Cells(48, excelcol) = allinterpartcopycheck
                            xlWorkSheet.Cells(48, excelcol).AddComment("Document Link Broken")
                            Exit For

                        End If
                        Dim objPartDocument As SolidEdgePart.PartDocument = Nothing
                        objPartDocument = DirectCast(documents.Open(fullpath), SolidEdgeFramework.SolidEdgeDocument)
                        adjustable = objPartDocument.IsAdjustablePart
                        hardwarepart = objPartDocument.HardwareFile

                        Try
                            isflatpattern = "No"
                            Dim flatpatternmodel As SolidEdgePart.FlatPatternModels = objPartDocument.FlatPatternModels
                            If flatpatternmodel.Count = 0 Then
                                isflatpattern = "No"
                            Else
                                isflatpattern = "Yes"
                            End If
                        Catch ex As Exception

                        End Try



                        Dim models As SolidEdgePart.Models = Nothing
                        Dim model As SolidEdgePart.Model = Nothing
                        models = objPartDocument.Models

                        'check skect is fully defined or not
                        sketchisfullydefined = sketchdefined(objPartDocument)

                        'Check Supress
                        Dim features As SolidEdgePart.Features = objPartDocument.Models.Item(1).Features

                        For i = 1 To features.Count

                            Dim obj As Object = features.Item(i)
                            Dim supressvariable As Boolean = obj.suppress
                            If supressvariable = True Then
                                issupress = "Yes"
                                Exit For
                            End If

                        Next

                        iscutout = "Yes"
                        Try
                            model = models.Item(1)
                            Dim cutouts As SolidEdgePart.ExtrudedCutouts = model.ExtrudedCutouts
                            For Each cutout As SolidEdgePart.ExtrudedCutout In cutouts

                                Try
                                    Dim profile As Profile = cutout.Profile
                                    If profile.Circles2d.Count = "0" Then
                                        iscutout = "Yes"
                                    Else
                                        iscutout = "No"
                                    End If
                                Catch ex As Exception

                                End Try
                            Next
                        Catch ex As Exception

                        End Try
                        objPartDocument.UpdateOnFileSave = False
                        objPartDocument.Models.Item(1).Recompute()
                        ' Dim occurances As SolidEdgeAssembly.Occurrences = objAssemblyDocument.Occurrences
                        '  Dim a As Integer = objPartDocument.OccurrenceID
                        ' Dim occur As SolidEdgeAssembly.Occurrence = occurances.GetOccurrence(a)

                        'check all-interpartlink,copypart,geomtry broken

                        allinterpartcopycheck = interpartcopycheck(objPartDocument, mtc)

                        checkPartFeature = CheckPartConstructionFeature(objPartDocument)
                        objPartDocument.Close(SaveChanges:=False)
                    ElseIf fullpath.EndsWith(".asm") Then
                        Dim asmdoc As SolidEdgeAssembly.AssemblyDocument = Nothing
                        asmdoc = DirectCast(documents.Open(fullpath), SolidEdgeFramework.SolidEdgeDocument)
                        Dim occurrences As SolidEdgeAssembly.Occurrences = asmdoc.Occurrences
                        Dim partlst As List(Of String) = New List(Of String)

                        For Each occur As SolidEdgeAssembly.Occurrence In occurrences
                            If Not partlst.Contains(occur.Name) Then
                                partlst.Add(occur.Name)
                                partlistcount = partlistcount + 1
                            End If

                        Next
                        checkAssemblyFeature = CheckAssemblyFeatureExistence(asmdoc)

                        asmdoc.Close(False)
                    End If


                    If adjustable = "True" Then
                        isadjustable = "Yes"
                    ElseIf adjustable = "False" Then
                        isadjustable = "No"
                    End If
                End If

                Dim activedoc As SolidEdgeAssembly.AssemblyDocument = objApplication.ActiveDocument
                Dim asmpathdir As String = IO.Path.GetDirectoryName(activedoc.FullName)
                Dim statusfile As Boolean = False
                Dim interference As Boolean = False

                Dim files() As String = IO.Directory.GetFiles(asmpathdir, "*.txt")

                For Each file As String In files
                    Dim interferencefile As String = IO.Path.GetFileName(file)


                    If interferencefile.Contains("Status") Then

                        statusfile = True

                    End If
                    If interferencefile.Contains("InterferenceReport") Then

                        interference = True

                    End If

                    '' "ChildInterferenceReportStatus"
                    '' "TopLevelInterferenceReportStatus"
                    'If interferencefile.StartsWith(activedoc.Name) Then
                    '    '    interferencereport = "Yes"
                    'Else
                    '    interferencereport = "No"
                    'End If
                Next

                'objApplication.StartCommand(33090)


                xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
                xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
                xlWorkSheet.Cells(1, excelcol).font.Bold = True
                ' xlWorkSheet.Cells(1, excelcol).AddComment("MILI " + DateTime.Now.ToString())
                xlWorkSheet.Cells(1, excelcol) = fileNameWithoutExt
                'xlWorkSheet.Cells(3, excelcol) = fileNameWithoutExt
                xlWorkSheet.Cells(4, excelcol) = revision
                xlWorkSheet.Cells(5, excelcol) = author

                If ispartfound = True Then
                    xlWorkSheet.Cells(3, excelcol) = "Yes"
                    xlWorkSheet.Cells(3, excelcol).AddComment("PartName:" + fileNameWithoutExt)
                Else
                    xlWorkSheet.Cells(3, excelcol) = "No"
                    xlWorkSheet.Cells(3, excelcol).AddComment("PartName:" + fileNameWithoutExt)

                End If

                If projectnamelst.Contains(projectname) Then
                    xlWorkSheet.Cells(6, excelcol) = "Yes"
                    xlWorkSheet.Cells(6, excelcol).AddComment("ProjectName:" + projectname)
                Else
                    xlWorkSheet.Cells(6, excelcol) = "No"
                    xlWorkSheet.Cells(6, excelcol).AddComment("ProjectName:" + projectname)
                End If

                If revision = revisionlevel Then
                    xlWorkSheet.Cells(7, excelcol) = "Yes"
                    xlWorkSheet.Cells(7, excelcol).AddComment("Revision Number:" + revision)
                Else
                    xlWorkSheet.Cells(7, excelcol) = "No"
                    xlWorkSheet.Cells(7, excelcol).AddComment("Revision Number:" + revision)
                End If

                '210-07306-1.par:1
                If documentno = fileNameWithoutExt Then
                    xlWorkSheet.Cells(8, excelcol) = "Yes"
                    xlWorkSheet.Cells(8, excelcol).AddComment("Document Number:" + documentno)
                ElseIf fileNameWithoutExt.Contains(documentno) Then
                    xlWorkSheet.Cells(8, excelcol) = "Yes"
                    xlWorkSheet.Cells(8, excelcol).AddComment("Document Number:" + documentno)
                Else
                    xlWorkSheet.Cells(8, excelcol) = "No"
                    xlWorkSheet.Cells(8, excelcol).AddComment("Document Number:" + documentno)
                End If

                If Authorlst.Contains(author) Then
                    xlWorkSheet.Cells(9, excelcol) = "Yes"
                    xlWorkSheet.Cells(9, excelcol).AddComment("Author:" + author)
                Else
                    xlWorkSheet.Cells(9, excelcol) = "No"
                    xlWorkSheet.Cells(9, excelcol).AddComment("Author:" + author)
                End If


                If author = Nothing Or Title = Nothing Or materialused = Nothing Or matlspec = Nothing Or projectname = Nothing Or density = Nothing Or revision = Nothing Or documentno = Nothing Or UomProperty = Nothing Or category = Nothing Or comments = Nothing Or keywords = Nothing Then
                    xlWorkSheet.Cells(10, excelcol) = "No"
                Else
                    xlWorkSheet.Cells(10, excelcol) = "Yes"
                End If

                If density = "" Then
                    xlWorkSheet.Cells(11, excelcol) = "No"
                    xlWorkSheet.Cells(11, excelcol).AddComment("Density:" + density)
                Else
                    xlWorkSheet.Cells(11, excelcol) = "Yes"
                    xlWorkSheet.Cells(11, excelcol).AddComment("Density:" + density)
                End If
                If fullpath.EndsWith(".asm") Then
                    xlWorkSheet.Cells(12, excelcol) = "Yes"
                Else
                    xlWorkSheet.Cells(12, excelcol) = "No"
                End If

                If matlspec = "" Then
                    xlWorkSheet.Cells(13, excelcol) = "No"
                    xlWorkSheet.Cells(13, excelcol).AddComment("Material Specification:" + matlspec)
                Else
                    xlWorkSheet.Cells(13, excelcol) = "Yes"
                    xlWorkSheet.Cells(13, excelcol).AddComment("Material Specification:" + matlspec)
                End If

                If materialused = "" Then
                    xlWorkSheet.Cells(14, excelcol) = "No"
                    xlWorkSheet.Cells(14, excelcol).AddComment("Material Used:" + materialused)
                Else
                    xlWorkSheet.Cells(14, excelcol) = "Yes"
                    xlWorkSheet.Cells(14, excelcol).AddComment("Material Used:" + materialused)
                End If
                xlWorkSheet.Cells(15, excelcol) = gageeexcelfile

                If Documenttype = "BaseLined" And hardwarepart = True Then
                    xlWorkSheet.Cells(19, excelcol) = "Yes"
                    xlWorkSheet.Cells(19, excelcol).AddComment("Status & HardwarePart:" + Documenttype + "&" + hardwarepart.ToString())
                Else
                    xlWorkSheet.Cells(19, excelcol) = "No"
                    xlWorkSheet.Cells(19, excelcol).AddComment("Status & HardwarePart:" + Documenttype + "&" + hardwarepart.ToString())
                End If


                If M2Mmeasure = UomProperty Then
                    xlWorkSheet.Cells(23, excelcol) = "Yes"
                    xlWorkSheet.Cells(23, excelcol).AddComment("UOM Property:" + UomProperty)
                Else
                    xlWorkSheet.Cells(23, excelcol) = "No"
                    xlWorkSheet.Cells(23, excelcol).AddComment("UOM Property:" + UomProperty)
                End If

                If M2Mdescript = Title Then

                    xlWorkSheet.Cells(22, excelcol) = "Yes"
                    xlWorkSheet.Cells(22, excelcol).AddComment("Title:" + Title)
                Else
                    xlWorkSheet.Cells(22, excelcol) = "No"
                    xlWorkSheet.Cells(22, excelcol).AddComment("Title:" + Title)
                End If


                If commentformate.Replace(" ", "").Trim().ToUpper() = comments.Replace(" ", "").Trim().ToUpper() Then

                    xlWorkSheet.Cells(43, excelcol) = "Yes"
                    xlWorkSheet.Cells(43, excelcol).AddComment("Comments:" + comments)
                Else
                    xlWorkSheet.Cells(43, excelcol) = "No"
                    xlWorkSheet.Cells(43, excelcol).AddComment("Comments:" + comments)
                End If
                'If M2Mkeywords = keywords Then

                '        xlWorkSheet.Cells(44, excelcol) = "Yes"
                '        xlWorkSheet.Cells(44, excelcol).AddComment("Keywords:" + Title)
                '    Else
                '        xlWorkSheet.Cells(44, excelcol) = "No"
                '        xlWorkSheet.Cells(44, excelcol).AddComment("Keywords:" + Title)
                '    End If
                If statusfile = True And interference = True Then
                    xlWorkSheet.Cells(45, excelcol) = "Yes"
                Else
                    xlWorkSheet.Cells(45, excelcol) = "No"
                End If
                If statusfile = False Then
                    xlWorkSheet.Cells(45, excelcol).AddComment("Please Run InteferenceTool")
                End If



                xlWorkSheet.Cells(44, excelcol) = partlistcount.ToString()
                xlWorkSheet.Cells(46, excelcol) = SEfeatures
                xlWorkSheet.Cells(47, excelcol) = "Yes"
                xlWorkSheet.Cells(16, excelcol) = isflatpattern
                xlWorkSheet.Cells(18, excelcol) = matlspec
                xlWorkSheet.Cells(20, excelcol) = isadjustable
                ' xlWorkSheet.Cells(32, excelcol) = partlistcount
                xlWorkSheet.Cells(38, excelcol) = iscutout



                'Check features

                xlWorkSheet.Cells(48, excelcol) = allinterpartcopycheck
                xlWorkSheet.Cells(49, excelcol) = checkAssemblyFeature
                'checkAssemblyFeature
                If issupress = "Yes" Then
                    xlWorkSheet.Cells(50, excelcol) = "Yes"
                Else
                    xlWorkSheet.Cells(50, excelcol) = "No"

                End If

                xlWorkSheet.Cells(51, excelcol) = sketchisfullydefined
                excelcol = excelcol + 1

            Catch ex As Exception
                MsgBox($"{ex.Message}{vbNewLine}{ex.StackTrace}")
            End Try
        Next




        'edit the cell with new value

        Dim Asmname As String = IO.Path.GetFileNameWithoutExtension(objAssemblyDocument.FullName)

        xlWorkSheet.Columns.AutoFit()
        Dim fileName As String = $"{Asmname}_MTC_Report_{excelName}"
        Dim newname As String = IO.Path.Combine(exportDirectoryLocation, $"{fileName}.xlsx") '  "C:\Users\milipatel\Downloads\MTC11111.xlsx"
        xlWorkBook.SaveAs(newname)
        xlWorkBook.Close()
        xlApp.Quit()

    End Sub

    Private Function SetVariablesData(ByRef dr As DataRow, ByVal mtcReviewObj As MTCReview, ByVal excelName As String) As MTCReview

#Region "Set Variables Data"

        '210-07293.psm:1
        If Not dr("Item Number") Is Nothing Then
            mtcReviewObj.itemnumber = dr("Item Number").ToString()
        End If

        If Not dr("File Name (no extension)") Is Nothing Then
            mtcReviewObj.fileNameWithoutExt = dr("File Name (no extension)").ToString()
            If mtcReviewObj.fileNameWithoutExt.Contains("_") Then
                Dim myDelims As String() = New String() {"_"}
                Dim splitfilename = mtcReviewObj.fileNameWithoutExt.Split(myDelims, StringSplitOptions.None)
                ' Dim splitfilename As String() = fileNameWithoutExt.Split("_")
                mtcReviewObj.fileNameWithoutExt = splitfilename(0)
                mtcReviewObj.revisionlevel = splitfilename(1)
                'mtcReviewObj.revisionlevel = GetRevisionLevel(mtcReviewObj.fileNameWithoutExt)
            End If
            'If fileNameWithoutExt.Contains("_") Then
            '    Dim splitfilename As String() = fileNameWithoutExt.Split("_")
            '    fileNameWithoutExt = splitfilename(0)
            '    revisionlevel = splitfilename(1)
            'End If
        End If

        If Not dr("Revision Number") Is Nothing Then
            mtcReviewObj.revision = dr("Revision Number").ToString()
        End If

        If Not dr("Title") Is Nothing Then
            mtcReviewObj.Title = dr("Title").ToString()
        End If

        If Not dr("Author") Is Nothing Then

            If excelName = "DGS" Then
                mtcReviewObj.author = "DGS"
            Else
                mtcReviewObj.author = dr("Author").ToString()
            End If

        End If

        If Not dr("Document Number") Is Nothing Then
            mtcReviewObj.documentno = dr("Document Number").ToString()
        End If

        If Not dr("Comments") Is Nothing Then
            mtcReviewObj.comments = dr("Comments").ToString()
        End If

        If Not dr("Category") Is Nothing Then
            mtcReviewObj.category = dr("Category").ToString()
        End If

        If Not dr("Material Used") Is Nothing Then
            mtcReviewObj.materialused = dr("Material Used").ToString()
        End If

        If Not dr("MATL SPEC") Is Nothing Then
            mtcReviewObj.matlspec = dr("MATL SPEC").ToString()
        End If

        If mtcReviewObj.matlspec.ToUpper() = "PURCHASED" Or mtcReviewObj.materialused.ToUpper() = "PURCHASED" Then
            mtcReviewObj.isBaseline = True
        End If

        If Not dr("File Name (full path)") Is Nothing Then
            mtcReviewObj.fullpath = dr("File Name (full path)").ToString()
        End If

        If Not dr("Last Author") Is Nothing Then
            mtcReviewObj.lastsaved = dr("Last Author").ToString()
        End If

        If Not dr("Density") Is Nothing Then
            mtcReviewObj.density = dr("Density").ToString()
        End If

        If Not dr("Project") Is Nothing Then
            mtcReviewObj.projectname = dr("Project").ToString()
        End If

        If mtcReviewObj.projectname = "BROOKVILLE EQUIPMENT CORP" Then
            mtcReviewObj.isBrookVilleProject_Baseline = "Yes"
        End If

        If Not dr("Status Text") Is Nothing Then
            mtcReviewObj.Documenttype = dr("Status Text").ToString()
        End If

        If Not dr("UOM") Is Nothing Then
            mtcReviewObj.UomProperty = dr("UOM").ToString()
        End If

        If Not dr("Keywords") Is Nothing Then
            mtcReviewObj.keywords = dr("Keywords").ToString()
        End If

        If Not dr("ECO/SOW") Is Nothing Then
            mtcReviewObj.ECO = dr("ECO/SOW").ToString()
        End If

#End Region

        Return mtcReviewObj
    End Function

    Private Function SetAdjustableDetails(ByVal mtcReviewObj As MTCReview) As MTCReview
        If mtcReviewObj.adjustable = "True" Then
            mtcReviewObj.isadjustable = "Yes"
        ElseIf mtcReviewObj.adjustable = "False" Then
            mtcReviewObj.isadjustable = "No"
        End If
        Return mtcReviewObj
    End Function

    Private Function SetInterferenceDetails(ByVal mtcReviewObj As MTCReview) As MTCReview

        Try

            Dim activedoc As SolidEdgeAssembly.AssemblyDocument = objApplication.ActiveDocument

            Dim asmpathdir As String = IO.Path.GetDirectoryName(activedoc.FullName)
            Dim files() As String = IO.Directory.GetFiles(asmpathdir, "*.txt")
            For Each file As String In files

                Dim interferencefile As String = IO.Path.GetFileName(file)
                If interferencefile.Contains("Status") Then
                    mtcReviewObj.statusfile = True
                End If

                If interferencefile.Contains("InterferenceReport") Then
                    mtcReviewObj.interference = True
                End If

            Next

        Catch ex As Exception

        End Try
        Return mtcReviewObj
    End Function

    Private Sub SetHorizontalAlignment(ByRef xlWorkSheet As Excel.Worksheet)
        Try
            xlWorkSheet.Columns.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

        Catch ex As Exception

        End Try
    End Sub

    Private Sub SetExcelColumnColorAndBorder(ByRef xlWorkSheet As Excel.Worksheet, ByVal maxRowCnt As Integer, ByVal excelcol As Integer)
        Try

            If Not CLng(excelcol) > 3 Then
                Exit Sub
            End If

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
                Catch ex As Exception
                End Try
                rowCount += 1
            Next
#End Region

        Catch ex As Exception

        End Try
    End Sub

    Private Sub BaseLineErrorColor(ByRef xlWorkSheet As Excel.Worksheet, ByVal maxRowCnt As Integer, ByVal excelcol As Integer)
        Try

            If Not CLng(excelcol) > 3 Then
                Exit Sub
            End If

            Dim excelRange As Range = xlWorkSheet.Cells(1, excelcol)
            excelRange.Interior.Color = Color.Red

            xlWorkSheet.Cells(1, excelcol).AddComment("Comments:" + "Invalid baseline directory Path.")

        Catch ex As Exception

        End Try

    End Sub



    Public Class MTCReview
        Public itemnumber As String = String.Empty
        Public fileNameWithoutExt As String = String.Empty
        Public revision As String = String.Empty
        Public author As String = String.Empty
        Public documentno As String = String.Empty
        Public materialused As String = String.Empty
        Public matlspec As String = String.Empty
        Public fullpath As String = String.Empty
        Public lastsaved As String = String.Empty
        Public density As String = String.Empty
        Public projectname As String = String.Empty
        Public isadjustable As String = String.Empty
        Public properties As String = "Yes"
        Public customprop As String = String.Empty
        Public columname As String = String.Empty
        Public Documenttype As String = String.Empty
        Public adjustable As Boolean
        Public hardwarepart As Boolean
        Public isflatpattern As String = String.Empty
        Public iscutout As String = String.Empty
        Public revisionlevel As String = String.Empty
        Public authorcheck As String = String.Empty
        Public UomProperty As String = String.Empty
        Public Title As String = String.Empty
        Public comments As String = String.Empty
        Public category As String = String.Empty
        Public keywords As String = String.Empty
        Public SEfeatures As String = String.Empty
        Public partlistcount As Integer = 0
        Public interferencereport As String = String.Empty
        Public checkPartFeature As String = String.Empty
        Public checkAssemblyFeature As String = String.Empty
        Public gageeexcelfile As String = String.Empty
        Public issupress As String = "No"
        Public sketchisfullydefined As String = String.Empty
        Public allinterpartcopycheck As String = String.Empty
        'Public statusfile

        Public statusfile As Boolean = False
        Public interference As Boolean = False

        Public ispartfound As Boolean = False

        Public ECO As String = String.Empty

        Public partCopiesDetected As String = "No"
        Public interPartCopiesDetected As String = "No"
        Public documentLinkBroken As String = "No"
        Public isBaseline As Boolean = False

        Public interPartLink As String = String.Empty


        Public isBrookVilleProject_Baseline As String = "No"

        Public isThreadExists As String = "No"
        Public isGeometryBroken As String = "Yes"

        'Public baseLineDirectoryPath As String = String.Empty
        Public isValidBaseLineDirectoryPath As Boolean = False

    End Class
    Public Class M2MData
        Public M2Mdescript As String = String.Empty
        Public M2Mmeasure As String = String.Empty
        Public M2Msource As String = String.Empty
        Public M2Mvendorname As String = String.Empty
        Public M2Mkeywords As String = String.Empty
        Public M2Mpartno As String = String.Empty
        Public commentformate As String = String.Empty
        Public M2Mflocation As String = String.Empty
        Public M2MFbin As String = String.Empty
    End Class

    Public Function SetM2mData(ByVal m2MDataObj As M2MData, ByRef dv As DataView) As M2MData

        For Each drv As DataRowView In dv

            m2MDataObj.M2Mpartno = drv(ExcelUtil.ExcelMtcReview.fpartno.ToString)
            m2MDataObj.M2Mdescript = drv(ExcelUtil.ExcelMtcReview.fdescript.ToString)
            m2MDataObj.M2Mmeasure = drv(ExcelUtil.ExcelMtcReview.fmeasure.ToString)
            m2MDataObj.M2Msource = drv(ExcelUtil.ExcelMtcReview.fsource.ToString)
            m2MDataObj.M2Mvendorname = drv(ExcelUtil.ExcelMtcReview.VendorName.ToString)

            m2MDataObj.M2Mflocation = drv(ExcelUtil.ExcelMtcReview.flocation.ToString)
            m2MDataObj.M2MFbin = drv(ExcelUtil.ExcelMtcReview.Fbin.ToString)
        Next
        m2MDataObj.commentformate = m2MDataObj.M2Mvendorname + " = " + m2MDataObj.M2Mpartno
        Return m2MDataObj
    End Function

    Private Sub SetPSMWorkSheet_MTC(ByRef mTCReviewObj As MTCReview, ByRef xlWorkSheet As Excel.Worksheet, ByRef objSheetMetalDocument As SolidEdgePart.SheetMetalDocument)

        If mTCReviewObj.itemnumber.Contains("*") Then

            mTCReviewObj.documentLinkBroken = "Yes"

            Exit Sub

        End If


        'Check sketch fullydefined or not 
        mTCReviewObj.sketchisfullydefined = sketchdefined(objSheetMetalDocument)

        Dim excelfilecount As Integer = getgagename(objSheetMetalDocument)
        If excelfilecount = 0 Then
            mTCReviewObj.gageeexcelfile = "No"
        ElseIf excelfilecount = 1 Then
            mTCReviewObj.gageeexcelfile = "Yes"
        End If

        Dim flatpatternmodel As SolidEdgePart.FlatPatternModels = objSheetMetalDocument.FlatPatternModels
        If flatpatternmodel.Count = 0 Then
            mTCReviewObj.isflatpattern = "No"
        Else
            mTCReviewObj.isflatpattern = "Yes"
        End If

        Dim models As SolidEdgePart.Models = objSheetMetalDocument.Models
        Dim model As SolidEdgePart.Model = Nothing
        objSheetMetalDocument.UpdateOnFileSave = False
        objSheetMetalDocument.Models.Item(1).Recompute()


        'Check supress 
        Dim features As SolidEdgePart.Features = objSheetMetalDocument.Models.Item(1).Features

        For i = 1 To features.Count

            Dim obj As Object = features.Item(i)
            Dim supressvariable As Boolean = obj.suppress
            If supressvariable = True Then
                mTCReviewObj.issupress = "Yes"
                Exit For
            End If

        Next
        'models = objSheetMetalDocument.Models
        mTCReviewObj.iscutout = ""
        Try
            model = models.Item(1)
            Try
                If model.ExtrudedCutouts.Count > 0 Then
                    Dim cutouts As SolidEdgePart.ExtrudedCutouts = model.ExtrudedCutouts
                    For Each cutout As SolidEdgePart.ExtrudedCutout In cutouts
                        Dim profile As Profile = cutout.Profile
                        If profile.Circles2d.Count = "0" Then
                            mTCReviewObj.iscutout = "Yes"
                        Else
                            mTCReviewObj.iscutout = "No"
                        End If
                        ' Dim circle As SolidEdgeFrameworkSupport.Circles2d = profile.Circles2d.co
                    Next
                ElseIf model.NormalCutouts.Count > 0 Then
                    Dim cutouts As SolidEdgePart.NormalCutouts = model.NormalCutouts
                    For Each cutout As SolidEdgePart.NormalCutout In cutouts
                        'Try
                        '    ' Dim noOfProfile As Object
                        '    '  Dim prfArray As Object
                        '    '  cutout.GetProfiles(noOfProfile, prfArray)
                        'Catch ex As Exception

                        'End Try

                        Dim profile As Profile = cutout.Profile
                        If profile.Circles2d.Count = "0" Then
                            mTCReviewObj.iscutout = "Yes"
                        Else
                            mTCReviewObj.iscutout = "No"
                        End If
                        ' Dim circle As SolidEdgeFrameworkSupport.Circles2d = profile.Circles2d.co
                    Next
                End If
            Catch ex As Exception

            End Try


        Catch ex As Exception
        End Try


        'Dim status As SolidEdgeAssembly.OccurrenceStatusConstants = Nothing
        'Dim occur As SolidEdgeAssembly.Occurrence = objSheetMetalDocument

        'Dim cutout As SolidEdgePart.ExtrudedCutout = cutouts.Item(1)
        Try
            mTCReviewObj.adjustable = objSheetMetalDocument.IsAdjustablePart
        Catch ex As Exception

            Debug.Print("aaaa")
        End Try

        'Check SEFeatures
        Try
            Dim SEModels As SimplifiedModels = objSheetMetalDocument.SimplifiedModels
            If SEModels.Count > 0 Then
                mTCReviewObj.SEfeatures = "Yes"
            Else
                mTCReviewObj.SEfeatures = "No"
            End If


        Catch ex As Exception


        End Try

        'check all-interpartlink,copypart,geomtry broken
        mTCReviewObj.allinterpartcopycheck = interpartcopycheck(objSheetMetalDocument, mTCReviewObj)

        mTCReviewObj.checkPartFeature = CheckPartConstructionFeature2(objSheetMetalDocument)

        objSheetMetalDocument.Close(SaveChanges:=False)



    End Sub

    Private Sub SetPSMWorkSheet_MTR(ByRef mTCReviewObj As MTCReview, ByRef xlWorkSheet As Excel.Worksheet, ByRef objSheetMetalDocument As SolidEdgePart.SheetMetalDocument)

        If mTCReviewObj.itemnumber.Contains("*") Then

            mTCReviewObj.documentLinkBroken = "Yes"

            Exit Sub

        End If


        'Check sketch fullydefined or not 
        mTCReviewObj.sketchisfullydefined = sketchdefined(objSheetMetalDocument)

        Dim excelfilecount As Integer = getgagename(objSheetMetalDocument)
        If excelfilecount = 0 Then
            mTCReviewObj.gageeexcelfile = "No"
        ElseIf excelfilecount = 1 Then
            mTCReviewObj.gageeexcelfile = "Yes"
        End If

        Dim flatpatternmodel As SolidEdgePart.FlatPatternModels = objSheetMetalDocument.FlatPatternModels
        If flatpatternmodel.Count = 0 Then
            mTCReviewObj.isflatpattern = "No"
        Else
            mTCReviewObj.isflatpattern = "Yes"
        End If

        Dim models As SolidEdgePart.Models = objSheetMetalDocument.Models
        Dim model As SolidEdgePart.Model = Nothing
        objSheetMetalDocument.UpdateOnFileSave = False
        objSheetMetalDocument.Models.Item(1).Recompute()


        'Check supress 
        Dim features As SolidEdgePart.Features = objSheetMetalDocument.Models.Item(1).Features

        For i = 1 To features.Count

            Dim obj As Object = features.Item(i)
            Dim supressvariable As Boolean = obj.suppress
            If supressvariable = True Then
                mTCReviewObj.issupress = "Yes"
                Exit For
            End If

        Next
        'models = objSheetMetalDocument.Models
        mTCReviewObj.iscutout = ""
        Try
            model = models.Item(1)
            Try
                If model.ExtrudedCutouts.Count > 0 Then
                    Dim cutouts As SolidEdgePart.ExtrudedCutouts = model.ExtrudedCutouts
                    For Each cutout As SolidEdgePart.ExtrudedCutout In cutouts
                        Dim profile As Profile = cutout.Profile
                        If profile.Circles2d.Count = "0" Then
                            mTCReviewObj.iscutout = "Yes"
                        Else
                            mTCReviewObj.iscutout = "No"
                        End If
                        ' Dim circle As SolidEdgeFrameworkSupport.Circles2d = profile.Circles2d.co
                    Next
                ElseIf model.NormalCutouts.Count > 0 Then
                    Dim cutouts As SolidEdgePart.NormalCutouts = model.NormalCutouts
                    For Each cutout As SolidEdgePart.NormalCutout In cutouts
                        'Try
                        '    ' Dim noOfProfile As Object
                        '    '  Dim prfArray As Object
                        '    '  cutout.GetProfiles(noOfProfile, prfArray)
                        'Catch ex As Exception

                        'End Try

                        Dim profile As Profile = cutout.Profile
                        If profile.Circles2d.Count = "0" Then
                            mTCReviewObj.iscutout = "Yes"
                        Else
                            mTCReviewObj.iscutout = "No"
                        End If
                        ' Dim circle As SolidEdgeFrameworkSupport.Circles2d = profile.Circles2d.co
                    Next
                End If
            Catch ex As Exception

            End Try


        Catch ex As Exception
        End Try


        'Dim status As SolidEdgeAssembly.OccurrenceStatusConstants = Nothing
        'Dim occur As SolidEdgeAssembly.Occurrence = objSheetMetalDocument

        'Dim cutout As SolidEdgePart.ExtrudedCutout = cutouts.Item(1)
        Try
            mTCReviewObj.adjustable = objSheetMetalDocument.IsAdjustablePart
        Catch ex As Exception

            Debug.Print("aaaa")
        End Try

        'Check SEFeatures
        Try
            Dim SEModels As SimplifiedModels = objSheetMetalDocument.SimplifiedModels
            If SEModels.Count > 0 Then
                mTCReviewObj.SEfeatures = "Yes"
            Else
                mTCReviewObj.SEfeatures = "No"
            End If


        Catch ex As Exception
        End Try


        mTCReviewObj.isGeometryBroken = IsGeomtryBroken(objSheetMetalDocument, mTCReviewObj)

        'check all-interpartlink,copypart,geomtry broken
        mTCReviewObj.allinterpartcopycheck = interpartcopycheck(objSheetMetalDocument, mTCReviewObj)

        mTCReviewObj.checkPartFeature = CheckPartConstructionFeature2(objSheetMetalDocument)

        mTCReviewObj.interPartLink = CheckInterPartLinksPSM()

        objSheetMetalDocument.Close(SaveChanges:=False)



    End Sub

    Private Sub SetPSMBaseLineWorkSheet_MTC(ByRef mTCReviewObj As MTCReview, ByRef xlWorkSheet As Excel.Worksheet, ByRef objSheetMetalDocument As SolidEdgePart.SheetMetalDocument)

        If mTCReviewObj.itemnumber.Contains("*") Then

            mTCReviewObj.documentLinkBroken = "Yes"

            'xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
            'xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
            'xlWorkSheet.Cells(1, excelcol).font.Bold = True
            'xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.fileNameWithoutExt
            'xlWorkSheet.Cells(48, excelcol) = mTCReviewObj.allinterpartcopycheck
            'xlWorkSheet.Cells(48, excelcol).AddComment("Document Link Broken")

            Exit Sub

        End If


        'Check sketch fullydefined or not 

        mTCReviewObj.sketchisfullydefined = sketchdefined(objSheetMetalDocument)

        Dim excelfilecount As Integer = getgagename(objSheetMetalDocument)
        If excelfilecount = 0 Then
            mTCReviewObj.gageeexcelfile = "No"
        ElseIf excelfilecount = 1 Then
            mTCReviewObj.gageeexcelfile = "Yes"
        End If

        Dim flatpatternmodel As SolidEdgePart.FlatPatternModels = objSheetMetalDocument.FlatPatternModels
        If flatpatternmodel.Count = 0 Then
            mTCReviewObj.isflatpattern = "No"
        Else
            mTCReviewObj.isflatpattern = "Yes"
        End If

        Dim models As SolidEdgePart.Models = objSheetMetalDocument.Models
        Dim model As SolidEdgePart.Model = Nothing
        objSheetMetalDocument.UpdateOnFileSave = False
        objSheetMetalDocument.Models.Item(1).Recompute()


        'Check supress 
        Dim features As SolidEdgePart.Features = objSheetMetalDocument.Models.Item(1).Features

        For i = 1 To features.Count

            Dim obj As Object = features.Item(i)
            Dim supressvariable As Boolean = obj.suppress
            If supressvariable = True Then
                mTCReviewObj.issupress = "Yes"
                Exit For
            End If

        Next
        'models = objSheetMetalDocument.Models
        mTCReviewObj.iscutout = ""
        Try
            model = models.Item(1)
            Try
                If model.ExtrudedCutouts.Count > 0 Then
                    Dim cutouts As SolidEdgePart.ExtrudedCutouts = model.ExtrudedCutouts
                    For Each cutout As SolidEdgePart.ExtrudedCutout In cutouts
                        Dim profile As Profile = cutout.Profile
                        If profile.Circles2d.Count = "0" Then
                            mTCReviewObj.iscutout = "Yes"
                        Else
                            mTCReviewObj.iscutout = "No"
                        End If
                        ' Dim circle As SolidEdgeFrameworkSupport.Circles2d = profile.Circles2d.co
                    Next
                ElseIf model.NormalCutouts.Count > 0 Then
                    Dim cutouts As SolidEdgePart.NormalCutouts = model.NormalCutouts
                    For Each cutout As SolidEdgePart.NormalCutout In cutouts

                        'Try
                        '    ' Dim noOfProfile As Object
                        '    '  Dim prfArray As Object
                        '    '  cutout.GetProfiles(noOfProfile, prfArray)
                        'Catch ex As Exception
                        'End Try


                        Dim profile As Profile = cutout.Profile

                        If profile.Circles2d.Count = "0" Then
                            mTCReviewObj.iscutout = "Yes"
                        Else
                            mTCReviewObj.iscutout = "No"
                        End If

                        ' Dim circle As SolidEdgeFrameworkSupport.Circles2d = profile.Circles2d.co
                    Next
                End If
            Catch ex As Exception

            End Try


        Catch ex As Exception
        End Try


        'Dim status As SolidEdgeAssembly.OccurrenceStatusConstants = Nothing
        'Dim occur As SolidEdgeAssembly.Occurrence = objSheetMetalDocument

        'Dim cutout As SolidEdgePart.ExtrudedCutout = cutouts.Item(1)
        Try
            mTCReviewObj.adjustable = objSheetMetalDocument.IsAdjustablePart
        Catch ex As Exception

            Debug.Print("aaaa")
        End Try

        'Check SEFeatures
        Try
            Dim SEModels As SimplifiedModels = objSheetMetalDocument.SimplifiedModels
            If SEModels.Count > 0 Then
                mTCReviewObj.SEfeatures = "Yes"
            Else
                mTCReviewObj.SEfeatures = "No"
            End If


        Catch ex As Exception


        End Try

        'check all-interpartlink,copypart,geomtry broken
        mTCReviewObj.allinterpartcopycheck = interpartcopycheck(objSheetMetalDocument, mTCReviewObj)

        mTCReviewObj.checkPartFeature = CheckPartConstructionFeature2(objSheetMetalDocument)

        objSheetMetalDocument.Close(SaveChanges:=False)



    End Sub

    Private Sub SetPSMWorkSheet_OtherDetails_MTC(ByRef xlWorkSheet As Excel.Worksheet, ByRef mTCReviewObj As MTCReview, ByRef m2MDataObj As M2MData, ByRef excelcol As Integer)

        Dim cnt As Integer = 1
        '0. Assembly Name
        xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
        xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
        xlWorkSheet.Cells(1, excelcol).font.Bold = True
        xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.fileNameWithoutExt

        SetExcelColumnColorAndBorder(xlWorkSheet, 23, excelcol)

        '1. Eco Number

        xlWorkSheet.Cells(cnt + 1, excelcol) = mTCReviewObj.ECO
        If mTCReviewObj.revision = "0" Then
            xlWorkSheet.Cells(cnt + 5, excelcol).AddComment($"SOW = {mTCReviewObj.ECO}")
        Else
            xlWorkSheet.Cells(cnt + 5, excelcol).AddComment($"ECO = {mTCReviewObj.ECO}")
        End If

        '2. Part Number
        If mTCReviewObj.ispartfound = True Then
            xlWorkSheet.Cells(cnt + 2, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 2, excelcol).AddComment("PartName:" + mTCReviewObj.fileNameWithoutExt)
        Else
            xlWorkSheet.Cells(cnt + 2, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 2, excelcol).AddComment("PartName:" + mTCReviewObj.fileNameWithoutExt)
        End If

        '3. Revision Level
        xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.revision

        '4. Author
        xlWorkSheet.Cells(cnt + 4, excelcol) = mTCReviewObj.author

        '5. Project Name
        Try
            If projectnamelst.Contains(mTCReviewObj.projectname) Then
                xlWorkSheet.Cells(cnt + 5, excelcol) = "Yes"
                xlWorkSheet.Cells(cnt + 5, excelcol).AddComment("ProjectName:" + mTCReviewObj.projectname)
            Else
                xlWorkSheet.Cells(cnt + 5, excelcol) = "No"
                xlWorkSheet.Cells(cnt + 5, excelcol).AddComment("ProjectName:" + mTCReviewObj.projectname)
            End If
        Catch ex As Exception
            'MsgBox(ex.Message)
        End Try


        '6. Revision Number Correct
        If mTCReviewObj.revision = mTCReviewObj.revisionlevel Then
            xlWorkSheet.Cells(cnt + 6, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revision)
        Else
            xlWorkSheet.Cells(cnt + 6, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revision)
        End If

        '7. Document Number correct
        If mTCReviewObj.documentno = mTCReviewObj.fileNameWithoutExt Then
            xlWorkSheet.Cells(cnt + 7, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Document Number:" + mTCReviewObj.documentno)
        ElseIf mTCReviewObj.fileNameWithoutExt.Contains(mTCReviewObj.documentno) Then
            xlWorkSheet.Cells(cnt + 7, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Document Number:" + mTCReviewObj.documentno)
        Else
            xlWorkSheet.Cells(cnt + 7, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Document Number:" + mTCReviewObj.documentno)
        End If

        '8. Is the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) 
        If Authorlst.Contains(mTCReviewObj.author) Then
            xlWorkSheet.Cells(cnt + 8, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Author:" + mTCReviewObj.author)
        Else
            xlWorkSheet.Cells(cnt + 8, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Author:" + mTCReviewObj.author)
        End If

        '9. Do all technically unused properties have a "dash" populated? 
        If mTCReviewObj.author = Nothing Or mTCReviewObj.Title = Nothing Or mTCReviewObj.materialused = Nothing _
            Or mTCReviewObj.matlspec = Nothing Or mTCReviewObj.projectname = Nothing _
            Or mTCReviewObj.density = Nothing Or mTCReviewObj.revision = Nothing Or mTCReviewObj.documentno = Nothing Or mTCReviewObj.UomProperty = Nothing Or mTCReviewObj.category = Nothing Or mTCReviewObj.comments = Nothing Or mTCReviewObj.keywords = Nothing Then
            xlWorkSheet.Cells(cnt + 9, excelcol) = "No"
        Else
            xlWorkSheet.Cells(cnt + 9, excelcol) = "Yes"
        End If


        '10. Does the Model Title MATCH the M2M Item Master (INV) Description Field?
        If m2MDataObj.M2Mdescript = mTCReviewObj.Title Then

            xlWorkSheet.Cells(cnt + 10, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.Title)
        Else
            xlWorkSheet.Cells(cnt + 10, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.Title)
        End If

        '11. Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
        If mTCReviewObj.UomProperty.Contains($"({m2MDataObj.M2Mmeasure})") Then
            xlWorkSheet.Cells(cnt + 11, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
        Else
            xlWorkSheet.Cells(cnt + 11, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
        End If


        '12. Is the mat'l spec field populated? (indicated by RED in metadata audit file)
        If mTCReviewObj.matlspec = "" Then
            xlWorkSheet.Cells(cnt + 12, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 12, excelcol).AddComment("Material Specification:" + mTCReviewObj.matlspec)
        Else
            xlWorkSheet.Cells(cnt + 12, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 12, excelcol).AddComment("Material Specification:" + mTCReviewObj.matlspec)
        End If

        '13. Is the material used field populated? *
        If mTCReviewObj.materialused = "" Then
            xlWorkSheet.Cells(cnt + 13, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 13, excelcol).AddComment("Material Used:" + mTCReviewObj.materialused)
        Else
            xlWorkSheet.Cells(cnt + 13, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 13, excelcol).AddComment("Material Used:" + mTCReviewObj.materialused)
        End If

        '14. Is the bend radius of the part equal to or above the ASTM minimum? *
        xlWorkSheet.Cells(cnt + 14, excelcol) = mTCReviewObj.gageeexcelfile


        '15. . Is the flat pattern turned ON? (Applicable only to parts needing dxf or cut length) *
        xlWorkSheet.Cells(cnt + 15, excelcol) = mTCReviewObj.isflatpattern

        '16. Did I use the
        'tool for ALL holes requiring fasteners? (include clearance holes for        hardware, tapped holes, And Slots) *
        xlWorkSheet.Cells(cnt + 16, excelcol) = mTCReviewObj.iscutout


        '17. Is the part "Adjustable"? (part should NOT be adjustable) *
        xlWorkSheet.Cells(cnt + 17, excelcol) = mTCReviewObj.isadjustable

        '18. Does the M2M Source reflect part creation (Make child of stk/pur parent, Project        procurement strategy, material types Like composite, etc.)? *
        xlWorkSheet.Cells(cnt + 18, excelcol) = GetM2MSource(m2MDataObj.M2Msource)

        excelcol = excelcol + 1
    End Sub


    Private Sub SetPARWorkSheet_MTC(ByRef mTCReviewObj As MTCReview, ByRef xlWorkSheet As Excel.Worksheet, ByRef objPartDocument As SolidEdgePart.PartDocument)
        'If mTCReviewObj.itemnumber.Contains("*") Then
        '    mTCReviewObj.allinterpartcopycheck = "No"
        '    xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
        '    xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
        '    xlWorkSheet.Cells(1, excelcol).font.Bold = True
        '    xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.fileNameWithoutExt
        '    xlWorkSheet.Cells(48, excelcol) = mTCReviewObj.allinterpartcopycheck
        '    xlWorkSheet.Cells(48, excelcol).AddComment("Document Link Broken")
        '    Exit Sub

        'End If
        If mTCReviewObj.itemnumber.Contains("*") Then
            mTCReviewObj.documentLinkBroken = "Yes"
            Exit Sub
        End If


        mTCReviewObj.adjustable = objPartDocument.IsAdjustablePart
        mTCReviewObj.hardwarepart = objPartDocument.HardwareFile

        Try
            mTCReviewObj.isflatpattern = "No"
            Dim flatpatternmodel As SolidEdgePart.FlatPatternModels = objPartDocument.FlatPatternModels
            If flatpatternmodel.Count = 0 Then
                mTCReviewObj.isflatpattern = "No"
            Else
                mTCReviewObj.isflatpattern = "Yes"
            End If
        Catch ex As Exception

        End Try



        Dim models As SolidEdgePart.Models = Nothing
        Dim model As SolidEdgePart.Model = Nothing
        models = objPartDocument.Models

        'check skect is fully defined or not
        mTCReviewObj.sketchisfullydefined = sketchdefined(objPartDocument)

        'Check Supress
        Dim features As SolidEdgePart.Features = objPartDocument.Models.Item(1).Features

        Try
            mTCReviewObj.issupress = "No"
            For i = 1 To features.Count

                Dim obj As Object = features.Item(i)
                Dim supressvariable As Boolean = obj.suppress
                If supressvariable = True Then
                    mTCReviewObj.issupress = "Yes"
                    Exit For
                End If

            Next
        Catch ex As Exception

        End Try


        mTCReviewObj.iscutout = "Yes"
        Try
            model = models.Item(1)
            Dim cutouts As SolidEdgePart.ExtrudedCutouts = model.ExtrudedCutouts
            For Each cutout As SolidEdgePart.ExtrudedCutout In cutouts

                Try
                    Dim profile As Profile = cutout.Profile
                    If profile.Circles2d.Count = "0" Then
                        mTCReviewObj.iscutout = "Yes"
                    Else
                        mTCReviewObj.iscutout = "No"
                    End If
                Catch ex As Exception

                End Try
            Next
        Catch ex As Exception

        End Try
        objPartDocument.UpdateOnFileSave = False
        objPartDocument.Models.Item(1).Recompute()
        ' Dim occurances As SolidEdgeAssembly.Occurrences = objAssemblyDocument.Occurrences
        '  Dim a As Integer = objPartDocument.OccurrenceID
        ' Dim occur As SolidEdgeAssembly.Occurrence = occurances.GetOccurrence(a)

        'check all-interpartlink,copypart,geomtry broken
        mTCReviewObj.allinterpartcopycheck = interpartcopycheck(objPartDocument, mTCReviewObj)

        mTCReviewObj.checkPartFeature = CheckPartConstructionFeature(objPartDocument)
        objPartDocument.Close(SaveChanges:=False)


    End Sub

    Private Sub SetPSMWorkSheet_OtherDetails_MTR(ByRef xlWorkSheet As Excel.Worksheet, ByRef mTCReviewObj As MTCReview, ByRef m2MDataObj As M2MData, ByRef excelcol As Integer)

        Dim cnt As Integer = 1
        '0. Assembly Name
        xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
        xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
        xlWorkSheet.Cells(1, excelcol).font.Bold = True
        xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.fileNameWithoutExt

        SetExcelColumnColorAndBorder(xlWorkSheet, 24, excelcol)

        '1. Verify that ALL features have been fully constrained
        xlWorkSheet.Cells(cnt + 1, excelcol) = mTCReviewObj.sketchisfullydefined

        '2. Verify that ALL suppressed and unused features have been removed
        xlWorkSheet.Cells(cnt + 2, excelcol) = mTCReviewObj.issupress

        '3. Verify that the part model is NOT adjustable
        xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.isadjustable

        '4. Verify that the inter-part copies are broken when released
        xlWorkSheet.Cells(cnt + 4, excelcol) = mTCReviewObj.interPartCopiesDetected

        '5.  Verify that the part copies are broken when released
        xlWorkSheet.Cells(cnt + 5, excelcol) = mTCReviewObj.partCopiesDetected

        '6. Verify the all categories below have been populated with relative Information Or, at a minimum, a dash (Summary, project, custom)
        If mTCReviewObj.author = Nothing Or mTCReviewObj.Title = Nothing Or mTCReviewObj.materialused = Nothing _
            Or mTCReviewObj.matlspec = Nothing Or mTCReviewObj.projectname = Nothing _
            Or mTCReviewObj.density = Nothing Or mTCReviewObj.revision = Nothing Or mTCReviewObj.documentno = Nothing Or mTCReviewObj.UomProperty = Nothing Or mTCReviewObj.category = Nothing Or mTCReviewObj.comments = Nothing Or mTCReviewObj.keywords = Nothing Then
            xlWorkSheet.Cells(cnt + 6, excelcol) = "No"
        Else
            xlWorkSheet.Cells(cnt + 6, excelcol) = "Yes"
        End If

        '7. Verify that the weight and mass has been applied
        xlWorkSheet.Cells(cnt + 7, excelcol) = mTCReviewObj.density

        '8. Verify that the “Update on File Save” is UNCHECKED
        xlWorkSheet.Cells(cnt + 8, excelcol) = ""

        '9. Verify that the included geometry is broken when released
        xlWorkSheet.Cells(cnt + 9, excelcol) = mTCReviewObj.isGeometryBroken 'CheckInterPartLinksPSM()

        excelcol = excelcol + 1

    End Sub

    Private Sub SetPARWorkSheet_MTR(ByRef mTCReviewObj As MTCReview, ByRef xlWorkSheet As Excel.Worksheet, ByRef objPartDocument As SolidEdgePart.PartDocument)
        'If mTCReviewObj.itemnumber.Contains("*") Then
        '    mTCReviewObj.allinterpartcopycheck = "No"
        '    xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
        '    xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
        '    xlWorkSheet.Cells(1, excelcol).font.Bold = True
        '    xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.fileNameWithoutExt
        '    xlWorkSheet.Cells(48, excelcol) = mTCReviewObj.allinterpartcopycheck
        '    xlWorkSheet.Cells(48, excelcol).AddComment("Document Link Broken")
        '    Exit Sub

        'End If
        If mTCReviewObj.itemnumber.Contains("*") Then
            mTCReviewObj.documentLinkBroken = "Yes"
            Exit Sub
        End If


        mTCReviewObj.adjustable = objPartDocument.IsAdjustablePart
        mTCReviewObj.hardwarepart = objPartDocument.HardwareFile

        Try
            mTCReviewObj.isflatpattern = "No"
            Dim flatpatternmodel As SolidEdgePart.FlatPatternModels = objPartDocument.FlatPatternModels
            If flatpatternmodel.Count = 0 Then
                mTCReviewObj.isflatpattern = "No"
            Else
                mTCReviewObj.isflatpattern = "Yes"
            End If
        Catch ex As Exception

        End Try



        Dim models As SolidEdgePart.Models = Nothing
        Dim model As SolidEdgePart.Model = Nothing
        models = objPartDocument.Models

        'check skect is fully defined or not
        mTCReviewObj.sketchisfullydefined = sketchdefined(objPartDocument)

        'Check Supress
        Dim features As SolidEdgePart.Features = objPartDocument.Models.Item(1).Features

        For i = 1 To features.Count

            Dim obj As Object = features.Item(i)
            Dim supressvariable As Boolean = obj.suppress
            If supressvariable = True Then
                mTCReviewObj.issupress = "Yes"
                Exit For
            End If

        Next

        mTCReviewObj.iscutout = "Yes"
        Try
            model = models.Item(1)
            Dim cutouts As SolidEdgePart.ExtrudedCutouts = model.ExtrudedCutouts
            For Each cutout As SolidEdgePart.ExtrudedCutout In cutouts

                Try
                    Dim profile As Profile = cutout.Profile
                    If profile.Circles2d.Count = "0" Then
                        mTCReviewObj.iscutout = "Yes"
                    Else
                        mTCReviewObj.iscutout = "No"
                    End If
                Catch ex As Exception

                End Try
            Next
        Catch ex As Exception

        End Try
        objPartDocument.UpdateOnFileSave = False
        objPartDocument.Models.Item(1).Recompute()
        ' Dim occurances As SolidEdgeAssembly.Occurrences = objAssemblyDocument.Occurrences
        '  Dim a As Integer = objPartDocument.OccurrenceID
        ' Dim occur As SolidEdgeAssembly.Occurrence = occurances.GetOccurrence(a)

        'check all-interpartlink,copypart,geomtry broken
        mTCReviewObj.allinterpartcopycheck = interpartcopycheck(objPartDocument, mTCReviewObj)

        mTCReviewObj.checkPartFeature = CheckPartConstructionFeature(objPartDocument)

        mTCReviewObj.interPartLink = CheckInterPartLinksPAR()


        objPartDocument.Close(SaveChanges:=False)


    End Sub

    Private Sub SetParWorkSheet_OtherDetails_MTR(ByRef xlWorkSheet As Excel.Worksheet, ByRef mTCReviewObj As MTCReview, ByRef m2MDataObj As M2MData, ByRef excelcol As Integer)

        Dim cnt As Integer = 1
        '0. Assembly Name
        xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
        xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
        xlWorkSheet.Cells(1, excelcol).font.Bold = True
        xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.fileNameWithoutExt

        SetExcelColumnColorAndBorder(xlWorkSheet, 17, excelcol)

        '1. Verify that ALL features have been fully constrained
        xlWorkSheet.Cells(cnt + 1, excelcol) = mTCReviewObj.sketchisfullydefined

        '2. Verify that ALL suppressed and unused features have been removed
        xlWorkSheet.Cells(cnt + 2, excelcol) = mTCReviewObj.issupress

        '3. Verify that the part model is NOT adjustable
        xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.isadjustable

        '4. Verify the all categories below have been populated with relative  Information Or, at a minimum, a dash (Summary, project, custom)
        If mTCReviewObj.author = Nothing Or mTCReviewObj.Title = Nothing Or mTCReviewObj.materialused = Nothing _
            Or mTCReviewObj.matlspec = Nothing Or mTCReviewObj.projectname = Nothing _
            Or mTCReviewObj.density = Nothing Or mTCReviewObj.revision = Nothing Or mTCReviewObj.documentno = Nothing Or mTCReviewObj.UomProperty = Nothing Or mTCReviewObj.category = Nothing Or mTCReviewObj.comments = Nothing Or mTCReviewObj.keywords = Nothing Then
            xlWorkSheet.Cells(cnt + 4, excelcol) = "No"
        Else
            xlWorkSheet.Cells(cnt + 4, excelcol) = "Yes"
        End If

        '5. If the model is a fastener then the HARDWARE PART box should be checked
        If mTCReviewObj.hardwarepart Then
            xlWorkSheet.Cells(cnt + 5, excelcol) = "Yes"
        Else
            xlWorkSheet.Cells(cnt + 5, excelcol) = "No"
        End If



        '6. Verify that the weight and mass has been applied
        xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.density

        '7. Verify that the “Update on File Save” is UNCHECKED
        xlWorkSheet.Cells(cnt + 7, excelcol) = ""


        excelcol = excelcol + 1
    End Sub

    Private Sub SetASMWorkSheet_MTR(ByRef mTCReviewObj As MTCReview, ByRef asmdoc As SolidEdgeAssembly.AssemblyDocument)

        Dim occurrences As SolidEdgeAssembly.Occurrences = asmdoc.Occurrences
        Dim partlst As List(Of String) = New List(Of String)

        For Each occur As SolidEdgeAssembly.Occurrence In occurrences
            If Not partlst.Contains(occur.Name) Then
                partlst.Add(occur.Name)
                mTCReviewObj.partlistcount = mTCReviewObj.partlistcount + 1
            End If

        Next
        mTCReviewObj.checkAssemblyFeature = CheckAssemblyFeatureExistence(asmdoc)

        mTCReviewObj.interPartLink = CheckInterPartLinksASM()

        mTCReviewObj.isGeometryBroken = IsGeomtryBroken_assembly(asmdoc)

        asmdoc.Close(False)


    End Sub

    Private Sub SetPARBaseLineWorkSheet_MTC(ByRef mTCReviewObj As MTCReview, ByRef xlWorkSheet As Excel.Worksheet, ByRef objPartDocument As SolidEdgePart.PartDocument)
        'If mTCReviewObj.itemnumber.Contains("*") Then
        '    mTCReviewObj.allinterpartcopycheck = "No"
        '    xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
        '    xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
        '    xlWorkSheet.Cells(1, excelcol).font.Bold = True
        '    xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.fileNameWithoutExt
        '    xlWorkSheet.Cells(48, excelcol) = mTCReviewObj.allinterpartcopycheck
        '    xlWorkSheet.Cells(48, excelcol).AddComment("Document Link Broken")
        '    Exit Sub

        'End If
        If mTCReviewObj.itemnumber.Contains("*") Then
            mTCReviewObj.documentLinkBroken = "Yes"
            Exit Sub
        End If


        mTCReviewObj.adjustable = objPartDocument.IsAdjustablePart
        mTCReviewObj.hardwarepart = objPartDocument.HardwareFile

        Try
            mTCReviewObj.isflatpattern = "No"
            Dim flatpatternmodel As SolidEdgePart.FlatPatternModels = objPartDocument.FlatPatternModels
            If flatpatternmodel.Count = 0 Then
                mTCReviewObj.isflatpattern = "No"
            Else
                mTCReviewObj.isflatpattern = "Yes"
            End If
        Catch ex As Exception

        End Try



        Dim models As SolidEdgePart.Models = Nothing
        Dim model As SolidEdgePart.Model = Nothing
        models = objPartDocument.Models

        'check skect is fully defined or not
        mTCReviewObj.sketchisfullydefined = sketchdefined(objPartDocument)

        mTCReviewObj.isThreadExists = CheckThread(objPartDocument)

        'Check Supress
        Dim features As SolidEdgePart.Features = objPartDocument.Models.Item(1).Features

        For i = 1 To features.Count

            Dim obj As Object = features.Item(i)
            Dim supressvariable As Boolean = obj.suppress
            If supressvariable = True Then
                mTCReviewObj.issupress = "Yes"
                Exit For
            End If

        Next

        mTCReviewObj.iscutout = "Yes"
        Try
            model = models.Item(1)
            Dim cutouts As SolidEdgePart.ExtrudedCutouts = model.ExtrudedCutouts
            For Each cutout As SolidEdgePart.ExtrudedCutout In cutouts

                Try
                    Dim profile As Profile = cutout.Profile
                    If profile.Circles2d.Count = "0" Then
                        mTCReviewObj.iscutout = "Yes"
                    Else
                        mTCReviewObj.iscutout = "No"
                    End If
                Catch ex As Exception

                End Try
            Next
        Catch ex As Exception

        End Try
        objPartDocument.UpdateOnFileSave = False
        objPartDocument.Models.Item(1).Recompute()
        ' Dim occurances As SolidEdgeAssembly.Occurrences = objAssemblyDocument.Occurrences
        '  Dim a As Integer = objPartDocument.OccurrenceID
        ' Dim occur As SolidEdgeAssembly.Occurrence = occurances.GetOccurrence(a)

        'check all-interpartlink,copypart,geomtry broken
        mTCReviewObj.allinterpartcopycheck = interpartcopycheck(objPartDocument, mTCReviewObj)

        mTCReviewObj.checkPartFeature = CheckPartConstructionFeature(objPartDocument)
        objPartDocument.Close(SaveChanges:=False)


    End Sub

    Private Sub SetParWorkSheet_OtherDetails_MTC(ByRef xlWorkSheet As Excel.Worksheet, ByRef mTCReviewObj As MTCReview, ByRef m2MDataObj As M2MData, ByRef excelcol As Integer)

        Dim cnt As Integer = 1
        '0. Assembly Name
        xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
        xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
        xlWorkSheet.Cells(1, excelcol).font.Bold = True
        xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.fileNameWithoutExt

        SetExcelColumnColorAndBorder(xlWorkSheet, 21, excelcol)

        '1. Eco Number

        xlWorkSheet.Cells(cnt + 1, excelcol) = mTCReviewObj.ECO
        If mTCReviewObj.revision = "0" Then
            xlWorkSheet.Cells(cnt + 1, excelcol).AddComment($"SOW = {mTCReviewObj.ECO}")
        Else
            xlWorkSheet.Cells(cnt + 1, excelcol).AddComment($"ECO = {mTCReviewObj.ECO}")
        End If

        '2. Part Number
        If mTCReviewObj.ispartfound = True Then
            xlWorkSheet.Cells(cnt + 2, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 2, excelcol).AddComment("PartName:" + mTCReviewObj.fileNameWithoutExt)
        Else
            xlWorkSheet.Cells(cnt + 2, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 2, excelcol).AddComment("PartName:" + mTCReviewObj.fileNameWithoutExt)
        End If

        '3. Revision Level
        xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.revision

        '4. Author
        xlWorkSheet.Cells(cnt + 4, excelcol) = mTCReviewObj.author

        '5. Project Name
        Try
            If projectnamelst.Contains(mTCReviewObj.projectname) Then
                xlWorkSheet.Cells(cnt + 5, excelcol) = "Yes"
                xlWorkSheet.Cells(cnt + 5, excelcol).AddComment("ProjectName:" + mTCReviewObj.projectname)
            Else
                xlWorkSheet.Cells(cnt + 5, excelcol) = "No"
                xlWorkSheet.Cells(cnt + 5, excelcol).AddComment("ProjectName:" + mTCReviewObj.projectname)
            End If
        Catch ex As Exception

        End Try


        '6. Revision Number Correct
        If mTCReviewObj.revision = mTCReviewObj.revisionlevel Then
            xlWorkSheet.Cells(cnt + 6, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revision)
        Else
            xlWorkSheet.Cells(cnt + 6, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revision)
        End If

        '7. Document Number correct
        If mTCReviewObj.documentno = mTCReviewObj.fileNameWithoutExt Then
            xlWorkSheet.Cells(cnt + 7, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Document Number:" + mTCReviewObj.documentno)
        ElseIf mTCReviewObj.fileNameWithoutExt.Contains(mTCReviewObj.documentno) Then
            xlWorkSheet.Cells(cnt + 7, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Document Number:" + mTCReviewObj.documentno)
        Else
            xlWorkSheet.Cells(cnt + 7, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Document Number:" + mTCReviewObj.documentno)
        End If

        '8. Is the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) 
        If Authorlst.Contains(mTCReviewObj.author) Then
            xlWorkSheet.Cells(cnt + 8, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Author:" + mTCReviewObj.author)
        Else
            xlWorkSheet.Cells(cnt + 8, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Author:" + mTCReviewObj.author)
        End If

        '9. Do all technically unused properties have a "dash" populated? 
        If mTCReviewObj.author = Nothing Or mTCReviewObj.Title = Nothing Or mTCReviewObj.materialused = Nothing _
            Or mTCReviewObj.matlspec = Nothing Or mTCReviewObj.projectname = Nothing _
            Or mTCReviewObj.density = Nothing Or mTCReviewObj.revision = Nothing Or mTCReviewObj.documentno = Nothing Or mTCReviewObj.UomProperty = Nothing Or mTCReviewObj.category = Nothing Or mTCReviewObj.comments = Nothing Or mTCReviewObj.keywords = Nothing Then
            xlWorkSheet.Cells(cnt + 9, excelcol) = "No"
        Else
            xlWorkSheet.Cells(cnt + 9, excelcol) = "Yes"
        End If


        '10. Does the Model Title MATCH the M2M Item Master (INV) Description Field?
        If m2MDataObj.M2Mdescript = mTCReviewObj.Title Then

            xlWorkSheet.Cells(cnt + 10, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.Title)
        Else
            xlWorkSheet.Cells(cnt + 10, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.Title)
        End If

        '11. Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
        If mTCReviewObj.UomProperty.Contains($"({m2MDataObj.M2Mmeasure})") Then 'If m2MDataObj.M2Mmeasure = mTCReviewObj.UomProperty Then
            xlWorkSheet.Cells(cnt + 11, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
        Else
            xlWorkSheet.Cells(cnt + 11, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
        End If


        '12. Is the mat'l spec field populated? (indicated by RED in metadata audit file)
        If mTCReviewObj.matlspec = "" Then
            xlWorkSheet.Cells(cnt + 12, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 12, excelcol).AddComment("Material Specification:" + mTCReviewObj.matlspec)
        Else
            xlWorkSheet.Cells(cnt + 12, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 12, excelcol).AddComment("Material Specification:" + mTCReviewObj.matlspec)
        End If

        '13. Is the material used field populated? *
        If mTCReviewObj.materialused = "" Then
            xlWorkSheet.Cells(cnt + 13, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 13, excelcol).AddComment("Material Used:" + mTCReviewObj.materialused)
        Else
            xlWorkSheet.Cells(cnt + 13, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 13, excelcol).AddComment("Material Used:" + mTCReviewObj.materialused)
        End If



        '14. Did I use the hole tool for ALL holes requiring fasteners? (include clearance holes for hardware, tapped holes, And Slots) *     
        '37
        xlWorkSheet.Cells(cnt + 14, excelcol) = mTCReviewObj.iscutout

        '15. Is sketch fully constrain?
        xlWorkSheet.Cells(cnt + 15, excelcol) = $"{mTCReviewObj.sketchisfullydefined}"


        '16. Have any suppressed (unused) features been removed from the model Pathfinder? *
        '49
        If mTCReviewObj.issupress = "Yes" Then
            xlWorkSheet.Cells(cnt + 16, excelcol) = "Yes"
        Else
            xlWorkSheet.Cells(cnt + 16, excelcol) = "No"
        End If

        '17. Is the part "Adjustable"? (part should NOT be adjustable) *
        '19
        xlWorkSheet.Cells(cnt + 17, excelcol) = mTCReviewObj.isadjustable


        '18. Does the M2M Source reflect part creation (Make child of stk/pur parent, Project  procurement strategy, material types Like composite, etc.)? *
        '23
        xlWorkSheet.Cells(cnt + 18, excelcol) = GetM2MSource(m2MDataObj.M2Msource)


        '20.
#Region "Old"


        'If mTCReviewObj.author = Nothing Or mTCReviewObj.Title = Nothing Or mTCReviewObj.materialused = Nothing _
        '    Or mTCReviewObj.matlspec = Nothing Or mTCReviewObj.projectname = Nothing _
        '    Or mTCReviewObj.density = Nothing Or mTCReviewObj.revision = Nothing Or mTCReviewObj.documentno = Nothing Or mTCReviewObj.UomProperty = Nothing Or mTCReviewObj.category = Nothing Or mTCReviewObj.comments = Nothing Or mTCReviewObj.keywords = Nothing Then
        '    xlWorkSheet.Cells(10, excelcol) = "No"
        'Else
        '    xlWorkSheet.Cells(10, excelcol) = "Yes"
        'End If

        'If mTCReviewObj.density = "" Then
        '    xlWorkSheet.Cells(11, excelcol) = "No"
        '    xlWorkSheet.Cells(11, excelcol).AddComment("Density:" + mTCReviewObj.density)
        'Else
        '    xlWorkSheet.Cells(11, excelcol) = "Yes"
        '    xlWorkSheet.Cells(11, excelcol).AddComment("Density:" + mTCReviewObj.density)
        'End If

        'If mTCReviewObj.fullpath.EndsWith(".asm") Then
        '    xlWorkSheet.Cells(12, excelcol) = "Yes"
        'Else
        '    xlWorkSheet.Cells(12, excelcol) = "No"
        'End If

        'xlWorkSheet.Cells(15, excelcol) = mTCReviewObj.gageeexcelfile

        'If mTCReviewObj.Documenttype = "BaseLined" And mTCReviewObj.hardwarepart = True Then
        '    xlWorkSheet.Cells(19, excelcol) = "Yes"
        '    xlWorkSheet.Cells(19, excelcol).AddComment("Status & HardwarePart:" + mTCReviewObj.Documenttype + "&" + mTCReviewObj.hardwarepart.ToString())
        'Else
        '    xlWorkSheet.Cells(19, excelcol) = "No"
        '    xlWorkSheet.Cells(19, excelcol).AddComment("Status & HardwarePart:" + mTCReviewObj.Documenttype + "&" + mTCReviewObj.hardwarepart.ToString())
        'End If

        'If m2MDataObj.M2Mmeasure = mTCReviewObj.UomProperty Then
        '    xlWorkSheet.Cells(23, excelcol) = "Yes"
        '    xlWorkSheet.Cells(23, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
        'Else
        '    xlWorkSheet.Cells(23, excelcol) = "No"
        '    xlWorkSheet.Cells(23, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
        'End If

        'If m2MDataObj.M2Mdescript = mTCReviewObj.Title Then

        '    xlWorkSheet.Cells(22, excelcol) = "Yes"
        '    xlWorkSheet.Cells(22, excelcol).AddComment("Title:" + mTCReviewObj.Title)
        'Else
        '    xlWorkSheet.Cells(22, excelcol) = "No"
        '    xlWorkSheet.Cells(22, excelcol).AddComment("Title:" + mTCReviewObj.Title)
        'End If

        'If m2MDataObj.commentformate.Replace(" ", "").Trim().ToUpper() = mTCReviewObj.comments.Replace(" ", "").Trim().ToUpper() Then

        '    xlWorkSheet.Cells(43, excelcol) = "Yes"
        '    xlWorkSheet.Cells(43, excelcol).AddComment("Comments:" + mTCReviewObj.comments)
        'Else
        '    xlWorkSheet.Cells(43, excelcol) = "No"
        '    xlWorkSheet.Cells(43, excelcol).AddComment("Comments:" + mTCReviewObj.comments)
        'End If

        'If mTCReviewObj.statusfile = True And mTCReviewObj.interference = True Then
        '    xlWorkSheet.Cells(45, excelcol) = "Yes"
        'Else
        '    xlWorkSheet.Cells(45, excelcol) = "No"
        'End If

        'If mTCReviewObj.statusfile = False Then
        '    xlWorkSheet.Cells(45, excelcol).AddComment("Please Run InteferenceTool")
        'End If

        'xlWorkSheet.Cells(44, excelcol) = mTCReviewObj.partlistcount.ToString()
        'xlWorkSheet.Cells(46, excelcol) = mTCReviewObj.SEfeatures
        'xlWorkSheet.Cells(47, excelcol) = "Yes"
        'xlWorkSheet.Cells(16, excelcol) = mTCReviewObj.isflatpattern
        'xlWorkSheet.Cells(18, excelcol) = mTCReviewObj.matlspec





        ''Check features

        'xlWorkSheet.Cells(48, excelcol) = mTCReviewObj.allinterpartcopycheck
        'xlWorkSheet.Cells(49, excelcol) = mTCReviewObj.checkAssemblyFeature

        ''checkAssemblyFeature
        'xlWorkSheet.Cells(51, excelcol) = mTCReviewObj.sketchisfullydefined
#End Region

        excelcol = excelcol + 1
    End Sub

    Private Function GetM2MSource(ByVal m2mSource As String) As String
        Dim m2mSource2 As String = String.Empty
        'B-BUY, M-MAKE, P-PHANTOM,S-STOCK

        If m2mSource = "S" Then
            m2mSource2 = "STOCK"
        ElseIf m2mSource = "B" Then
            m2mSource2 = "BUY"
        ElseIf m2mSource = "M" Then
            m2mSource2 = "MAKE"
        ElseIf m2mSource = "P" Then
            m2mSource2 = "PHANTOM"
        End If

        Return m2mSource2

    End Function

    Private Sub SetASMWorkSheet_MTC(ByRef mTCReviewObj As MTCReview, ByRef asmdoc As SolidEdgeAssembly.AssemblyDocument)

        Try
            Dim occurrences As SolidEdgeAssembly.Occurrences = asmdoc.Occurrences
            Dim partlst As List(Of String) = New List(Of String)

            For Each occur As SolidEdgeAssembly.Occurrence In occurrences
                If Not partlst.Contains(occur.Name) Then
                    partlst.Add(occur.Name)
                    mTCReviewObj.partlistcount = mTCReviewObj.partlistcount + 1
                End If

            Next
            mTCReviewObj.checkAssemblyFeature = CheckAssemblyFeatureExistence(asmdoc)

            mTCReviewObj.interPartLink = CheckInterPartLinksPSM()

        Catch ex As Exception

        End Try

        asmdoc.Close(False)


    End Sub

    Private Sub SetASMWorkSheet_OtherDetails_MTC(ByRef xlWorkSheet As Excel.Worksheet, ByRef mTCReviewObj As MTCReview, ByRef m2MDataObj As M2MData, ByRef excelcol As Integer)
        Dim cnt As Integer = 1
        '0. Assembly Name
        xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
        xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
        xlWorkSheet.Cells(1, excelcol).font.Bold = True
        xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.fileNameWithoutExt

        SetExcelColumnColorAndBorder(xlWorkSheet, 22, excelcol)

        '1. Eco Number

        xlWorkSheet.Cells(cnt + 1, excelcol) = mTCReviewObj.ECO
        If mTCReviewObj.revision = "0" Then
            xlWorkSheet.Cells(cnt + 1, excelcol).AddComment($"SOW = {mTCReviewObj.ECO}")
        Else
            xlWorkSheet.Cells(cnt + 1, excelcol).AddComment($"ECO = {mTCReviewObj.ECO}")
        End If

        '2. Part Number
        If mTCReviewObj.ispartfound = True Then
            xlWorkSheet.Cells(cnt + 2, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 2, excelcol).AddComment("PartName:" + mTCReviewObj.fileNameWithoutExt)
        Else
            xlWorkSheet.Cells(cnt + 2, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 2, excelcol).AddComment("PartName:" + mTCReviewObj.fileNameWithoutExt)
        End If

        '3. Revision Level
        xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.revision

        '4. Author
        xlWorkSheet.Cells(cnt + 4, excelcol) = mTCReviewObj.author

        '5. Project Name
        Try
            If projectnamelst.Contains(mTCReviewObj.projectname) Then
                xlWorkSheet.Cells(cnt + 5, excelcol) = "Yes"
                xlWorkSheet.Cells(cnt + 5, excelcol).AddComment("ProjectName:" + mTCReviewObj.projectname)
            Else
                xlWorkSheet.Cells(cnt + 5, excelcol) = "No"
                xlWorkSheet.Cells(cnt + 5, excelcol).AddComment("ProjectName:" + mTCReviewObj.projectname)
            End If
        Catch ex As Exception

        End Try


        '6. Revision Number Correct
        If mTCReviewObj.revision = mTCReviewObj.revisionlevel Then
            xlWorkSheet.Cells(cnt + 6, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revision)
        Else
            xlWorkSheet.Cells(cnt + 6, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 6, excelcol).AddComment("Revision Number:" + mTCReviewObj.revision)
        End If

        '7. Document Number correct
        If mTCReviewObj.documentno = mTCReviewObj.fileNameWithoutExt Then
            xlWorkSheet.Cells(cnt + 7, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Document Number:" + mTCReviewObj.documentno)
        ElseIf mTCReviewObj.fileNameWithoutExt.Contains(mTCReviewObj.documentno) Then
            xlWorkSheet.Cells(cnt + 7, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Document Number:" + mTCReviewObj.documentno)
        Else
            xlWorkSheet.Cells(cnt + 7, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 7, excelcol).AddComment("Document Number:" + mTCReviewObj.documentno)
        End If

        '8. Is the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) 
        If Authorlst.Contains(mTCReviewObj.author) Then
            xlWorkSheet.Cells(cnt + 8, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Author:" + mTCReviewObj.author)
        Else
            xlWorkSheet.Cells(cnt + 8, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Author:" + mTCReviewObj.author)
        End If

        '9. Do all technically unused properties have a "dash" populated? 
        If mTCReviewObj.author = Nothing Or mTCReviewObj.Title = Nothing Or mTCReviewObj.materialused = Nothing _
            Or mTCReviewObj.matlspec = Nothing Or mTCReviewObj.projectname = Nothing _
            Or mTCReviewObj.density = Nothing Or mTCReviewObj.revision = Nothing Or mTCReviewObj.documentno = Nothing Or mTCReviewObj.UomProperty = Nothing Or mTCReviewObj.category = Nothing Or mTCReviewObj.comments = Nothing Or mTCReviewObj.keywords = Nothing Then
            xlWorkSheet.Cells(cnt + 9, excelcol) = "No"
        Else
            xlWorkSheet.Cells(cnt + 9, excelcol) = "Yes"
        End If


        '10. Does the Model Title MATCH the M2M Item Master (INV) Description Field?
        If m2MDataObj.M2Mdescript = mTCReviewObj.Title Then

            xlWorkSheet.Cells(cnt + 10, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.Title)
        Else
            xlWorkSheet.Cells(cnt + 10, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.Title)
        End If

        '11. Do all UOMs match M2M? (i.e. EA, sqft, in, etc.)
        If mTCReviewObj.UomProperty.Contains($"({m2MDataObj.M2Mmeasure})") Then 'If m2MDataObj.M2Mmeasure = mTCReviewObj.UomProperty Then
            xlWorkSheet.Cells(cnt + 11, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
        Else
            xlWorkSheet.Cells(cnt + 11, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
        End If

        '12. Perform a Parts List Report. How many total BOM items are in this assembly/weldment? (Not qty of all parts, only line items that will show up on draft PL) *
        xlWorkSheet.Cells(cnt + 12, excelcol) = $"{mTCReviewObj.partlistcount}"


        '13. Is interfernces found in assembly?
        If mTCReviewObj.statusfile = True And mTCReviewObj.interference = True Then
            xlWorkSheet.Cells(cnt + 13, excelcol) = "Yes"
        Else
            xlWorkSheet.Cells(cnt + 13, excelcol) = "No"
        End If

        '14. inter-part copies detected
        xlWorkSheet.Cells(cnt + 14, excelcol) = mTCReviewObj.interPartCopiesDetected

        '15. part copies detected
        xlWorkSheet.Cells(cnt + 15, excelcol) = mTCReviewObj.partCopiesDetected

        '16. broken  file Path detected
        xlWorkSheet.Cells(cnt + 16, excelcol) = mTCReviewObj.documentLinkBroken

        '17. isadjustable
        xlWorkSheet.Cells(cnt + 17, excelcol) = mTCReviewObj.isadjustable


        excelcol = excelcol + 1
    End Sub

    Private Sub SetBaseLineWorkSheet_OtherDetails_MTC(ByRef xlWorkSheet As Excel.Worksheet, ByRef mTCReviewObj As MTCReview, ByRef m2MDataObj As M2MData, ByRef excelcol As Integer)

        Dim cnt As Integer = 1

        '0. Assembly Name
        xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
        xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
        xlWorkSheet.Cells(1, excelcol).font.Bold = True
        xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.fileNameWithoutExt

        SetExcelColumnColorAndBorder(xlWorkSheet, 39, excelcol)

        If Not mTCReviewObj.isValidBaseLineDirectoryPath Then
            BaseLineErrorColor(xlWorkSheet, 39, excelcol)
        End If

        '1. Is the part number match with M2M? *
        If mTCReviewObj.ispartfound = True Then
            xlWorkSheet.Cells(cnt + 1, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 1, excelcol).AddComment("PartName:" + mTCReviewObj.fileNameWithoutExt)
        Else
            xlWorkSheet.Cells(cnt + 1, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 1, excelcol).AddComment("PartName:" + mTCReviewObj.fileNameWithoutExt)
        End If

        '2. What is the revision level? *
        xlWorkSheet.Cells(cnt + 2, excelcol) = mTCReviewObj.revision

        '3. Who is the Author of the file? (For BEC authors, type 3 capitalized initials of author)
        xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.author

        '4. What type of component?
        xlWorkSheet.Cells(cnt + 4, excelcol) = mTCReviewObj.category

        '5. Virtual thread applied for Fasteners?
        xlWorkSheet.Cells(cnt + 5, excelcol) = mTCReviewObj.isThreadExists

        '6. Are ALL features fully defined in the model Pathfinder? * (Sketch fully constraint?)
        xlWorkSheet.Cells(cnt + 6, excelcol) = mTCReviewObj.sketchisfullydefined

        '7. Any suppressed feature found?
        xlWorkSheet.Cells(cnt + 7, excelcol) = mTCReviewObj.issupress

        '8. Is the "Mat'l Spec" field populated? (PURCHASED for library components) *
        If mTCReviewObj.matlspec = "PURCHASED" Then
            xlWorkSheet.Cells(cnt + 8, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Material Specification:" + mTCReviewObj.matlspec)
        Else
            xlWorkSheet.Cells(cnt + 8, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 8, excelcol).AddComment("Material Specification:" + mTCReviewObj.matlspec)
        End If

        '9. Is the "Material Used" field populated? (PURCHASED for library components) *
        If mTCReviewObj.materialused = "PURCHASED" Then
            xlWorkSheet.Cells(cnt + 9, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 9, excelcol).AddComment("Material Used:" + mTCReviewObj.materialused)
        Else
            xlWorkSheet.Cells(cnt + 9, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 9, excelcol).AddComment("Material Used:" + mTCReviewObj.materialused)
        End If

        '10. Is the "Title" field populated with the correct part description? (This should MATCH the M2M Item Master (INV) Description field) *
        If m2MDataObj.M2Mdescript = mTCReviewObj.Title Then

            xlWorkSheet.Cells(cnt + 10, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.Title)
        Else
            xlWorkSheet.Cells(cnt + 10, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 10, excelcol).AddComment("Title:" + mTCReviewObj.Title)
        End If

        '11. Is the "Author" field populated with the full name? (DGS for DGSTS models regardless of individual name) *
        If Authorlst.Contains(mTCReviewObj.author) Then
            xlWorkSheet.Cells(cnt + 11, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("Author:" + mTCReviewObj.author)
        Else
            xlWorkSheet.Cells(cnt + 11, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 11, excelcol).AddComment("Author:" + mTCReviewObj.author)
        End If

        '12. Is the "Keywords" field populated with the FULL M2M Item Master Description as shown in the Comments field? *
        xlWorkSheet.Cells(cnt + 12, excelcol) = String.Empty


        '13. Is the "Comments" field populated with the Vendor name and Vendor part number? (It should appear as VENDOR NAME = VENDOR PART NUMBER) *
        If m2MDataObj.commentformate.Replace(" ", "").Trim().ToUpper() = mTCReviewObj.comments.Replace(" ", "").Trim().ToUpper() Then

            xlWorkSheet.Cells(cnt + 13, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 13, excelcol).AddComment("Comments:" + mTCReviewObj.comments)
        Else
            xlWorkSheet.Cells(cnt + 13, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 13, excelcol).AddComment("Comments:" + mTCReviewObj.comments)
        End If

        '14. Is the "Document Number" field populated with the correct part number? (This should  MATCH the M2M Item Master Part Number field) *
        If mTCReviewObj.documentno = mTCReviewObj.fileNameWithoutExt Then
            xlWorkSheet.Cells(cnt + 14, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 14, excelcol).AddComment("Document Number:" + mTCReviewObj.documentno)
        ElseIf mTCReviewObj.fileNameWithoutExt.Contains(mTCReviewObj.documentno) Then
            xlWorkSheet.Cells(cnt + 14, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 14, excelcol).AddComment("Document Number:" + mTCReviewObj.documentno)
        Else
            xlWorkSheet.Cells(cnt + 14, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 14, excelcol).AddComment("Document Number:" + mTCReviewObj.documentno)
        End If

        '15. Is the "Revision" field populated with the correct revision number? *
        If mTCReviewObj.revision = mTCReviewObj.revisionlevel Then
            xlWorkSheet.Cells(cnt + 15, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 15, excelcol).AddComment("Revision Number:" + mTCReviewObj.revision)
        Else
            xlWorkSheet.Cells(cnt + 15, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 15, excelcol).AddComment("Revision Number:" + mTCReviewObj.revision)
        End If

        '16. "Is the ""Project"" field populated with the correct Project Name? (The Project should be BROOKVILLE EQUIPMENT CORP) *"
        xlWorkSheet.Cells(cnt + 16, excelcol) = mTCReviewObj.isBrookVilleProject_Baseline

        '17. Is the "Hardware Part" box checked for hardware components? (This should be for ALL  nuts, bolts, washers, screws, etc.) *
        If mTCReviewObj.Documenttype = "BaseLined" And mTCReviewObj.hardwarepart = True Then
            xlWorkSheet.Cells(cnt + 17, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 17, excelcol).AddComment("Status & HardwarePart:" + mTCReviewObj.Documenttype + "&" + mTCReviewObj.hardwarepart.ToString())
        Else
            xlWorkSheet.Cells(cnt + 17, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 17, excelcol).AddComment("Status & HardwarePart:" + mTCReviewObj.Documenttype + "&" + mTCReviewObj.hardwarepart.ToString())
        End If

        '18. Do all other unused property fields have a "dash" (-) populated? *
        If mTCReviewObj.author = Nothing Or mTCReviewObj.Title = Nothing _
            Or mTCReviewObj.materialused = Nothing Or mTCReviewObj.matlspec = Nothing _
            Or mTCReviewObj.projectname = Nothing Or mTCReviewObj.density = Nothing _
            Or mTCReviewObj.revision = Nothing Or mTCReviewObj.documentno = Nothing _
            Or mTCReviewObj.UomProperty = Nothing Or mTCReviewObj.category = Nothing _
            Or mTCReviewObj.comments = Nothing Or mTCReviewObj.keywords = Nothing Then

            xlWorkSheet.Cells(cnt + 18, excelcol) = "No"
        Else
            xlWorkSheet.Cells(cnt + 18, excelcol) = "Yes"
        End If

        '19. Does the unit of measures (UOM) match M2M "For Inventory UOM"? (i.e. EA, sqft, in, etc.) *
        If mTCReviewObj.UomProperty.Contains($"({m2MDataObj.M2Mmeasure})") Then 'If m2MDataObj.M2Mmeasure = mTCReviewObj.UomProperty Then
            xlWorkSheet.Cells(cnt + 19, excelcol) = "Yes"
            xlWorkSheet.Cells(cnt + 19, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
        Else
            xlWorkSheet.Cells(cnt + 19, excelcol) = "No"
            xlWorkSheet.Cells(cnt + 19, excelcol).AddComment("UOM Property:" + mTCReviewObj.UomProperty)
        End If


        '20. Is the M2M Source marked stock/purchased? *

        If m2MDataObj.M2Msource = "S" Then
            xlWorkSheet.Cells(cnt + 20, excelcol) = "Yes"
        Else
            xlWorkSheet.Cells(cnt + 20, excelcol) = "No"
        End If


        '21. Did I use the hole tool for ALL holes requiring fasteners and they match the Vendor supplied Data?(This include clearances holes For hardware, tapped holes, And Slots) *
        xlWorkSheet.Cells(cnt + 21, excelcol) = mTCReviewObj.iscutout

        '22. Are all inter-part copies, part copies and included geometry broken? *
        xlWorkSheet.Cells(cnt + 22, excelcol) = String.Empty

        '23. inter-part copies detected?
        xlWorkSheet.Cells(cnt + 23, excelcol) = mTCReviewObj.interPartCopiesDetected

        '24. part copies detected?
        xlWorkSheet.Cells(cnt + 24, excelcol) = mTCReviewObj.partCopiesDetected

        '25. broken  file Path detected?
        xlWorkSheet.Cells(cnt + 25, excelcol) = mTCReviewObj.documentLinkBroken

        '26. Is the model a SE Adjustable Part? (Models should NOT be adjustable) *
        xlWorkSheet.Cells(cnt + 26, excelcol) = mTCReviewObj.isadjustable

        '27. Does this model have any SE Simplified Features? (NO model should contain SE Simplified Features) *
        xlWorkSheet.Cells(cnt + 27, excelcol) = mTCReviewObj.SEfeatures

        '28. Do all hardware AND library components in this asm model have an SE Status of Baselined?(If Available Or Released status hardware/library component exists, select "NO" ) *
        If mTCReviewObj.Documenttype.ToUpper() = "BASELINED" Then
            xlWorkSheet.Cells(cnt + 28, excelcol) = "Yes"
        Else
            xlWorkSheet.Cells(cnt + 28, excelcol) = "No"
        End If


        'mTCReviewObj.Documenttype = "BaseLined"

        excelcol = excelcol + 1
    End Sub

    Private Function IsValidBaseLineDirectoryPath(ByVal baselineDirPath As String, ByVal docFullPath As String)
        Dim resValidDir As Boolean = False
        Try
            Dim docDirPath As String = IO.Path.GetDirectoryName(docFullPath)

            If docDirPath.Contains(baselineDirPath) Then
                resValidDir = True
            End If

        Catch ex As Exception

        End Try
        Return resValidDir
    End Function

    Public Sub SaveAsExcelSplit_MTC(ByVal dt As System.Data.DataTable, ByVal exportDirectoryLocation As String, ByVal baseLineDirectoryLocation As String, ByVal excelName As String)

        'excelcol = 4

        Dim excelAsmCol As Integer = 4
        Dim excelPartCol As Integer = 4
        Dim excelSheetMetalCol As Integer = 4
        Dim excelBaseLineCol As Integer = 4
        Dim excelElectricalCol As Integer = 4

        Dim asmExcelRowCnt As Integer = 23
        Dim partExcelRowCnt As Integer = 22
        Dim sheetMetalExcelRowCnt As Integer = 24
        Dim baseLineExcelRowCnt As Integer = 41
        Dim electricalExcelRowCnt As Integer = 12


        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook


        Dim xlWorkSheetAssembly As Excel.Worksheet
        Dim xlWorkSheetPart As Excel.Worksheet
        Dim xlWorkSheetSheetMetal As Excel.Worksheet
        Dim xlWorkSheetBaseline As Excel.Worksheet
        Dim xlWorkSheetElectrical As Excel.Worksheet

        xlApp = New Application
        Dim aPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location
        Dim dirPath As String = IO.Path.GetDirectoryName(aPath)
        'Dim mtcExcelPath As String = IO.Path.Combine(dirPath, $"MTC.xlsx")
        Dim mtcExcelPath As String = IO.Path.Combine(dirPath, $"MTC_BEC.xlsx")
        '

        xlWorkBook = xlApp.Workbooks.Open(mtcExcelPath)
        'xlWorkSheet = xlWorkBook.Worksheets("MTC")

        xlWorkSheetAssembly = xlWorkBook.Worksheets("Assembly")
        xlWorkSheetPart = xlWorkBook.Worksheets("Part")
        xlWorkSheetSheetMetal = xlWorkBook.Worksheets("Sheetmetal")
        xlWorkSheetBaseline = xlWorkBook.Worksheets("Baseline")
        xlWorkSheetElectrical = xlWorkBook.Worksheets("Electrical")

        'SetHorizontalAlignment(xlWorkSheet)
        SetHorizontalAlignment(xlWorkSheetAssembly)
        SetHorizontalAlignment(xlWorkSheetPart)
        SetHorizontalAlignment(xlWorkSheetSheetMetal)
        SetHorizontalAlignment(xlWorkSheetBaseline)
        SetHorizontalAlignment(xlWorkSheetElectrical)


        'Dim uniqueCols As String() = dt.DefaultView.ToTable(True, "Name").AsEnumerable().[Select](Function(r) r.Field(Of String)("Name")).ToArray()
        Dim lstCompletedFileName As List(Of String) = New List(Of String)()

        Dim rCnt As Integer = 1
        For Each dr As DataRow In dt.Rows

            Dim filePath As String = dr("File Name (no extension)").ToString()
            log.Info($"{rCnt.ToString()}. {filePath}")

            'If Not dr("Select").ToString().ToUpper() = "TRUE" Then
            '    Continue For
            'End If

            Dim mTCReviewObj As MTCReview = New MTCReview()
            mTCReviewObj = SetVariablesData(dr, mTCReviewObj, excelName)

            mTCReviewObj.isValidBaseLineDirectoryPath = IsValidBaseLineDirectoryPath(baseLineDirectoryLocation, mTCReviewObj.fullpath)

            If lstCompletedFileName.Contains(mTCReviewObj.fileNameWithoutExt) Then
                Continue For
            Else
                lstCompletedFileName.Add(mTCReviewObj.fileNameWithoutExt)
            End If

            Try
                Dim dv As DataView = New DataView(dtM2M)
                Dim partno As String = ExcelUtil.ExcelMtcReview.fpartno.ToString()
                dv.RowFilter = $"{partno}='{mTCReviewObj.fileNameWithoutExt}'"

                Dim m2MDataObj As M2MData = New M2MData()
                m2MDataObj = SetM2mData(m2MDataObj, dv)


                If dv.Count = 0 Then
                    mTCReviewObj.ispartfound = False
                Else
                    mTCReviewObj.ispartfound = True
                End If

                'SetExcelColumnColorAndBorder(xlWorkSheetAssembly, 22)
                'SetExcelColumnColorAndBorder(xlWorkSheetPart, 21)
                'SetExcelColumnColorAndBorder(xlWorkSheetSheetMetal, 23)
                'SetExcelColumnColorAndBorder(xlWorkSheetBaseline, 39)
                'SetExcelColumnColorAndBorder(xlWorkSheetElectrical, 9)

                If dr("Select").ToString().ToUpper() = "TRUE" Then

                    Dim documents As SolidEdgeFramework.Documents = objApp.Documents

                    'If mTCReviewObj.Documenttype.ToUpper() = "BASELINE" Then

                    If mTCReviewObj.fullpath.EndsWith(".psm") Then

                        If mTCReviewObj.fullpath.Contains("210-04782.psm") Then
                            Debug.Print("aa")

                        End If
                        Dim objSheetMetalDocument As SolidEdgePart.SheetMetalDocument = DirectCast(documents.Open(mTCReviewObj.fullpath), SolidEdgeFramework.SolidEdgeDocument)

                        Threading.Thread.Sleep(2000)

                        If mTCReviewObj.isBaseline = False Then
                            SetPSMWorkSheet_MTC(mTCReviewObj, xlWorkSheetSheetMetal, objSheetMetalDocument)
                        Else
                            SetPSMBaseLineWorkSheet_MTC(mTCReviewObj, xlWorkSheetBaseline, objSheetMetalDocument)
                        End If
                    ElseIf mTCReviewObj.fullpath.EndsWith(".par") Then
                        Dim objPartDocument As SolidEdgePart.PartDocument = Nothing
                        Try
                            objPartDocument = DirectCast(documents.Open(mTCReviewObj.fullpath), SolidEdgeFramework.SolidEdgeDocument)
                            Threading.Thread.Sleep(2000)

                            If mTCReviewObj.isBaseline = False Then
                                SetPARWorkSheet_MTC(mTCReviewObj, xlWorkSheetPart, objPartDocument)
                            Else
                                SetPARBaseLineWorkSheet_MTC(mTCReviewObj, xlWorkSheetBaseline, objPartDocument)
                            End If
                        Catch ex As Exception
                            Debug.Print($"#######################Error in open part document {mTCReviewObj.fullpath} {vbNewLine} {ex.Message} {vbNewLine }{ex.StackTrace}")
                        End Try


                    ElseIf mTCReviewObj.fullpath.EndsWith(".asm") Then

                        Dim asmdoc As SolidEdgeAssembly.AssemblyDocument = Nothing
                        asmdoc = DirectCast(documents.Open(mTCReviewObj.fullpath), SolidEdgeFramework.SolidEdgeDocument)
                        Threading.Thread.Sleep(2000)
                        SetASMWorkSheet_MTC(mTCReviewObj, asmdoc)

                    End If

                    mTCReviewObj = SetAdjustableDetails(mTCReviewObj)

                End If

                mTCReviewObj = SetInterferenceDetails(mTCReviewObj)

                If mTCReviewObj.fullpath.EndsWith(".psm") Then

                    If mTCReviewObj.isBaseline = False Then
                        SetPSMWorkSheet_OtherDetails_MTC(xlWorkSheetSheetMetal, mTCReviewObj, m2MDataObj, excelSheetMetalCol)
                    Else
                        SetBaseLineWorkSheet_OtherDetails_MTC(xlWorkSheetBaseline, mTCReviewObj, m2MDataObj, excelBaseLineCol)
                    End If

                ElseIf mTCReviewObj.fullpath.EndsWith(".par") Then

                    If mTCReviewObj.isBaseline = False Then
                        SetParWorkSheet_OtherDetails_MTC(xlWorkSheetPart, mTCReviewObj, m2MDataObj, excelPartCol)
                    Else
                        SetBaseLineWorkSheet_OtherDetails_MTC(xlWorkSheetBaseline, mTCReviewObj, m2MDataObj, excelBaseLineCol)
                    End If

                ElseIf mTCReviewObj.fullpath.EndsWith(".asm") Then
                    SetASMWorkSheet_OtherDetails_MTC(xlWorkSheetAssembly, mTCReviewObj, m2MDataObj, excelAsmCol)
                End If

            Catch ex As Exception
                'MsgBox($"{ex.Message}{vbNewLine}{ex.StackTrace}")
                log.Error($"{filePath} Msg: {ex.Message} Trace: {ex.StackTrace}")
            End Try

            rCnt = rCnt + 1
        Next
        Try
            'SetExcelColumnColorAndBorder(xlWorkSheetAssembly, 22)
            'SetExcelColumnColorAndBorder(xlWorkSheetPart, 21)
            'SetExcelColumnColorAndBorder(xlWorkSheetSheetMetal, 23)
            'SetExcelColumnColorAndBorder(xlWorkSheetBaseline, 39)
            'SetExcelColumnColorAndBorder(xlWorkSheetElectrical, 9)
        Catch ex As Exception

        End Try
        xlWorkSheetAssembly.Columns.AutoFit()
        xlWorkSheetPart.Columns.AutoFit()
        xlWorkSheetSheetMetal.Columns.AutoFit()
        xlWorkSheetBaseline.Columns.AutoFit()
        xlWorkSheetElectrical.Columns.AutoFit()

        Dim Asmname As String = IO.Path.GetFileNameWithoutExtension(objAssemblyDocument.FullName)
        Dim fileName As String = $"{Asmname}_MTC_Report_{excelName}"
        Dim newname As String = IO.Path.Combine(exportDirectoryLocation, $"{fileName}.xlsx")
        xlWorkBook.SaveAs(newname)
        xlWorkBook.Close()
        xlApp.Quit()

    End Sub

    Private Sub SetASMWorkSheet_OtherDetails_MTR(ByRef xlWorkSheet As Excel.Worksheet, ByRef mTCReviewObj As MTCReview, ByRef m2MDataObj As M2MData, ByRef excelcol As Integer)

        Dim cnt As Integer = 1

        '0. Assembly Name
        xlWorkSheet.Cells(1, excelcol).Interior.Color = System.Drawing.Color.Gray
        xlWorkSheet.Cells(1, excelcol).font.Color = System.Drawing.Color.White
        xlWorkSheet.Cells(1, excelcol).font.Bold = True
        xlWorkSheet.Cells(1, excelcol) = mTCReviewObj.fileNameWithoutExt

        SetExcelColumnColorAndBorder(xlWorkSheet, 26, excelcol)

        '1. Verify that there are NO adjustable parts present in the assembly model
        xlWorkSheet.Cells(cnt + 1, excelcol) = mTCReviewObj.isadjustable


        '2. Verify that the inter-part copies are broken when released
        xlWorkSheet.Cells(cnt + 2, excelcol) = mTCReviewObj.interPartCopiesDetected

        '3. Verify that the part copies are broken when released
        xlWorkSheet.Cells(cnt + 3, excelcol) = mTCReviewObj.partCopiesDetected

        '4. Verify the all categories below have been populated with relative  Information Or, at a minimum, a dash (Summary, project, custom)
        If mTCReviewObj.author = Nothing Or mTCReviewObj.Title = Nothing Or mTCReviewObj.materialused = Nothing _
            Or mTCReviewObj.matlspec = Nothing Or mTCReviewObj.projectname = Nothing _
            Or mTCReviewObj.density = Nothing Or mTCReviewObj.revision = Nothing Or mTCReviewObj.documentno = Nothing Or mTCReviewObj.UomProperty = Nothing Or mTCReviewObj.category = Nothing Or mTCReviewObj.comments = Nothing Or mTCReviewObj.keywords = Nothing Then
            xlWorkSheet.Cells(cnt + 4, excelcol) = "No"
        Else
            xlWorkSheet.Cells(cnt + 4, excelcol) = "Yes"
        End If

        '5. Do assembly features exist within the assembly model?  If present, can they be removed?
        xlWorkSheet.Cells(cnt + 5, excelcol) = mTCReviewObj.checkAssemblyFeature


        ' If mTCReviewObj.statusfile = True And mTCReviewObj.interference = True Then
        'xlWorkSheet.Cells(cnt + 13, excelcol) = "Yes"
        'Else
        'xlWorkSheet.Cells(cnt + 13, excelcol) = "No"
        'End If
        '6. Verify that mating parts have been checked for interferences
        If mTCReviewObj.statusfile = True Then
            xlWorkSheet.Cells(cnt + 6, excelcol) = "Yes"
        Else
            xlWorkSheet.Cells(cnt + 6, excelcol) = "No"
        End If


        '7. Verify that there are no interferences with objects in the environment
        If mTCReviewObj.interference = True Then
            xlWorkSheet.Cells(cnt + 7, excelcol) = "Yes"
        Else
            xlWorkSheet.Cells(cnt + 7, excelcol) = "No"
        End If


        '8. Verify that the “Update on File Save” is UNCHECKED
        xlWorkSheet.Cells(cnt + 8, excelcol) = ""

        '9. Verify that the included geometry is broken when released
        xlWorkSheet.Cells(cnt + 9, excelcol) = mTCReviewObj.isGeometryBroken 'CheckInterPartLinksASM()

        excelcol = excelcol + 1
    End Sub

    Public Sub SaveAsExcelSplit_MTR(ByVal dt As System.Data.DataTable, ByVal exportDirectoryLocation As String, ByVal excelName As String)

        'excelcol = 4

        Dim excelAsmCol As Integer = 4
        Dim excelPartCol As Integer = 4
        Dim excelSheetMetalCol As Integer = 4
        Dim excelBaseLineCol As Integer = 4
        Dim excelElectricalCol As Integer = 4

        Dim asmExcelRowCnt As Integer = 23
        Dim partExcelRowCnt As Integer = 22
        Dim sheetMetalExcelRowCnt As Integer = 24
        Dim baseLineExcelRowCnt As Integer = 41
        Dim electricalExcelRowCnt As Integer = 12


        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook


        Dim xlWorkSheetAssembly As Excel.Worksheet
        Dim xlWorkSheetPart As Excel.Worksheet
        Dim xlWorkSheetSheetMetal As Excel.Worksheet


        xlApp = New Application
        Dim aPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location
        Dim dirPath As String = IO.Path.GetDirectoryName(aPath)
        Dim mtcExcelPath As String = IO.Path.Combine(dirPath, $"MTR_BEC.xlsx")
        xlWorkBook = xlApp.Workbooks.Open(mtcExcelPath)

        xlWorkSheetAssembly = xlWorkBook.Worksheets("Assembly")
        xlWorkSheetPart = xlWorkBook.Worksheets("Part")
        xlWorkSheetSheetMetal = xlWorkBook.Worksheets("Sheetmetal")


        SetHorizontalAlignment(xlWorkSheetAssembly)
        SetHorizontalAlignment(xlWorkSheetPart)
        SetHorizontalAlignment(xlWorkSheetSheetMetal)

        Dim lstCompletedFileName As List(Of String) = New List(Of String)()
        Dim rCnt As Integer = 1
        For Each dr As DataRow In dt.Rows

            Dim filePath As String = dr("File Name (no extension)").ToString()
            log.Info($"{rCnt.ToString()}. {filePath}")

            Dim mTCReviewObj As MTCReview = New MTCReview()

            mTCReviewObj = SetVariablesData(dr, mTCReviewObj, excelName)

            If lstCompletedFileName.Contains(mTCReviewObj.fileNameWithoutExt) Then
                Continue For
            Else
                lstCompletedFileName.Add(mTCReviewObj.fileNameWithoutExt)
            End If

            Try
                Dim dv As DataView = New DataView(dtM2M)
                Dim partno As String = ExcelUtil.ExcelMtcReview.fpartno.ToString()
                dv.RowFilter = $"{partno}='{mTCReviewObj.fileNameWithoutExt}'"

                Dim m2MDataObj As M2MData = New M2MData()
                m2MDataObj = SetM2mData(m2MDataObj, dv)


                If dv.Count = 0 Then
                    mTCReviewObj.ispartfound = False
                Else
                    mTCReviewObj.ispartfound = True
                End If

                'SetExcelColumnColorAndBorder(xlWorkSheetAssembly, 22)
                'SetExcelColumnColorAndBorder(xlWorkSheetPart, 21)
                'SetExcelColumnColorAndBorder(xlWorkSheetSheetMetal, 23)


                If dr("Select").ToString().ToUpper() = "TRUE" Then

                    Dim documents As SolidEdgeFramework.Documents = objApp.Documents

                    If mTCReviewObj.fullpath.EndsWith(".psm") Then

                        Dim objSheetMetalDocument As SolidEdgePart.SheetMetalDocument = DirectCast(documents.Open(mTCReviewObj.fullpath), SolidEdgeFramework.SolidEdgeDocument)

                        Threading.Thread.Sleep(2000)

                        SetPSMWorkSheet_MTR(mTCReviewObj, xlWorkSheetSheetMetal, objSheetMetalDocument)

                    ElseIf mTCReviewObj.fullpath.EndsWith(".par") Then

                        Dim objPartDocument As SolidEdgePart.PartDocument = Nothing

                        objPartDocument = DirectCast(documents.Open(mTCReviewObj.fullpath), SolidEdgeFramework.SolidEdgeDocument)

                        Threading.Thread.Sleep(2000)

                        SetPARWorkSheet_MTR(mTCReviewObj, xlWorkSheetPart, objPartDocument)

                    ElseIf mTCReviewObj.fullpath.EndsWith(".asm") Then

                        Dim asmdoc As SolidEdgeAssembly.AssemblyDocument = Nothing

                        asmdoc = DirectCast(documents.Open(mTCReviewObj.fullpath), SolidEdgeFramework.SolidEdgeDocument)

                        Threading.Thread.Sleep(2000)

                        SetASMWorkSheet_MTR(mTCReviewObj, asmdoc)

                    End If

                    mTCReviewObj = SetAdjustableDetails(mTCReviewObj)

                End If

                mTCReviewObj = SetInterferenceDetails(mTCReviewObj)

                If mTCReviewObj.fullpath.EndsWith(".psm") Then

                    SetPSMWorkSheet_OtherDetails_MTR(xlWorkSheetSheetMetal, mTCReviewObj, m2MDataObj, excelSheetMetalCol)

                ElseIf mTCReviewObj.fullpath.EndsWith(".par") Then

                    SetParWorkSheet_OtherDetails_MTR(xlWorkSheetPart, mTCReviewObj, m2MDataObj, excelPartCol)

                ElseIf mTCReviewObj.fullpath.EndsWith(".asm") Then

                    SetASMWorkSheet_OtherDetails_MTR(xlWorkSheetAssembly, mTCReviewObj, m2MDataObj, excelAsmCol)

                End If

            Catch ex As Exception
                'MsgBox($"{ex.Message}{vbNewLine}{ex.StackTrace}")
                log.Error($"{filePath} Msg: {ex.Message} Trace: {ex.StackTrace}")
            End Try

            rCnt = rCnt + 1
        Next

        Try
            'SetExcelColumnColorAndBorder(xlWorkSheetAssembly, 22)
            'SetExcelColumnColorAndBorder(xlWorkSheetPart, 21)
            'SetExcelColumnColorAndBorder(xlWorkSheetSheetMetal, 23)
        Catch ex As Exception
        End Try

        xlWorkSheetAssembly.Columns.AutoFit()
        xlWorkSheetPart.Columns.AutoFit()
        xlWorkSheetSheetMetal.Columns.AutoFit()

        Dim Asmname As String = IO.Path.GetFileNameWithoutExtension(objAssemblyDocument.FullName)
        Dim fileName As String = $"{Asmname}_MTR_Report_{excelName}"
        Dim newname As String = IO.Path.Combine(exportDirectoryLocation, $"{fileName}.xlsx")

        xlWorkBook.SaveAs(newname)
        xlWorkBook.Close()
        xlApp.Quit()

    End Sub

    Public Shared Sub releaseObject(ByVal obj As Object)
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

    Public Function getpartlist() As System.Data.DataTable

        Dim objDocuments As SolidEdgeFramework.Documents = Nothing
        Dim objDraft As SolidEdgeDraft.DraftDocument = Nothing
        Dim objSheet As SolidEdgeDraft.Sheet = Nothing
        Dim objModelLinks As SolidEdgeDraft.ModelLinks = Nothing
        Dim objModelLink As SolidEdgeDraft.ModelLink = Nothing
        Dim objDrawingViews As SolidEdgeDraft.DrawingViews = Nothing
        Dim objDrawingView As SolidEdgeDraft.DrawingView = Nothing

        Try
            '  OleMessageFilter.Register()

            ' Connect to a running instance of Solid Edge
            objApplication = Marshal.GetActiveObject("SolidEdge.Application")

            ' Get a reference to the documents collection
            objDocuments = objApplication.Documents

            ' Add a Draft document      
            objDraft = objDocuments.Add("SolidEdge.DraftDocument")

            ' Get a reference to the active sheet
            objSheet = objDraft.ActiveSheet

            ' Get a reference to the model links collection
            objModelLinks = objDraft.ModelLinks
            Dim filename As String
            Dim file As String = objAssemblyDocument.FullName
            filename = System.IO.Path.Combine(SolidEdgeCommunity.SolidEdgeUtils.GetTrainingFolderPath(), file)

            ' Add a new model link
            objModelLink = objModelLinks.Add(filename)

            ' Get a reference to the drawing views collection
            objDrawingViews = objSheet.DrawingViews


            objDrawingView = objDrawingViews.AddAssemblyView([From]:=objModelLink, Orientation:=SolidEdgeDraft.ViewOrientationConstants.igDimetricTopBackLeftView, Scale:=1.0, x:=0.4, y:=0.4, ViewType:=SolidEdgeDraft.AssemblyDrawingViewTypeConstants.seAssemblyDesignedView)

            ' Add a new drawing view
            'objDrawingView = objDrawingViews.AddAssemblyView(
            '  objModelLink,
            '  SolidEdgeDraft.ViewOrientationConstants.igFrontView,
            '  1,
            '  0.1,
            '  0.1,
            '  SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView)

            ' Assign a caption
            objDrawingView.Caption = "My New Drawing View"

            ' Ensure caption is displayed
            objDrawingView.DisplayCaption = False

        Catch ex As Exception
            Console.WriteLine(ex.Message)
        Finally
            ' OleMessageFilter.Revoke()
        End Try


        '  Dim objApp As SolidEdgeFramework.Application = Nothing
        Dim objDoc As SolidEdgeDraft.DraftDocument = Nothing
        Dim objPartsLists As SolidEdgeDraft.PartsLists = Nothing
        Dim objPartsList As SolidEdgeDraft.PartsList = Nothing

        ' objApp = GetObject(, "SolidEdge.Application")
        objDoc = objApp.ActiveDocument
        objPartsLists = objDoc.PartsLists
        objPartsList = objPartsLists.Add(objDrawingView, "BEC", 1, 1)
        ' objSheet.Activate()

        ' objAsm.LinkedDocuments(DesignManager.LinkTypeConstants.seLinkTypeAll)

        objPartsList = objPartsLists.Item(1)
        '  Dim dt As DataTable = objPartsList
        Dim tableCell As SolidEdgeDraft.TableCell = Nothing
        Dim dt As System.Data.DataTable = New Data.DataTable()
        Dim myDataColumn1 As DataColumn = New DataColumn()
        myDataColumn1 = New DataColumn()
        myDataColumn1.ColumnName = "Select"
        myDataColumn1.DefaultValue = "0"
        myDataColumn1.DataType = System.Type.GetType("System.Boolean")
        dt.Columns.Add(myDataColumn1)


        Dim cols As TableColumns = objPartsList.Columns
        Dim rows As TableRows = objPartsList.Rows
        For Each tableColumn In cols.OfType(Of SolidEdgeDraft.TableColumn)()
            Dim dtcolums As DataColumn = New DataColumn()
            dtcolums.ColumnName = tableColumn.Header
            dt.Columns.Add(dtcolums)
            Debug.Print(tableColumn.Header)

        Next tableColumn


        'For j As Integer = 1 To rows.Count
        '    Dim objrow As TableRow = rows.Item(j)

        '    tableCell = PartsList.Cell(TableRow.Index, TableColumn.Index)


        'Next

        Dim dtrows As DataRow = dt.NewRow()
        dt.Rows.Add(dtrows)








        For Each tableRow In rows.OfType(Of SolidEdgeDraft.TableRow)()
            For Each tableColumn In cols.OfType(Of SolidEdgeDraft.TableColumn)()
                If tableColumn.Show Then
                    'visibleRowCount += 1
                    tableCell = objPartsList.Cell(tableRow.Index, tableColumn.Index)
                    Dim rowindex As Integer = tableRow.Index - 1
                    Dim colindex As Integer = tableColumn.Index - 1 + 1
                    Dim tabvalue As String = tableCell.value

                    dt.Rows(rowindex).Item(colindex) = tabvalue
                    ' MsgBox(tableCell.value)
                    'excelRange = excelCells.Item(tableRow.Index + 1, tableColumn.Index)
                    'excelRange.Value = tableCell.value
                End If
            Next tableColumn
            If Not tableRow.Index = rows.Count Then
                dtrows = dt.NewRow()
                dt.Rows.Add(dtrows)

            End If
            ' visibleRowCount = 0


        Next tableRow





        ' objPartsList.CopyToClipboard()
        objApp.Documents.CloseDocument(objDoc.FullName, False, "", False, False)
        Return dt

    End Function

    Private Function CheckThread(ByVal ObjPartDoc As SolidEdgeFramework.SolidEdgeDocument) As String

        Dim isThreadExist As String = "No"
        Try
            Dim objApp As SolidEdgeFramework.Application = Nothing
            objApp = Marshal.GetActiveObject("SolidEdge.Application")
            ObjPartDoc = objApp.ActiveDocument

            Dim models As SolidEdgePart.Models = Nothing
            Dim model As SolidEdgePart.Model = Nothing
            models = ObjPartDoc.Models
            model = models.Item(1)
            If model.Threads.Count > 0 Then
                isThreadExist = "Yes"
            End If
        Catch ex As Exception

        End Try
        Return isThreadExist
    End Function

    Private Function sketchdefined(ByVal ObjPartDoc As SolidEdgeFramework.SolidEdgeDocument) As String
        Dim objApp As SolidEdgeFramework.Application = Nothing
        'Dim ObjPartDoc As SolidEdgeFramework.SolidEdgeDocument = Nothing
        Dim objProfiles As SolidEdgePart.Profiles = Nothing
        Dim objProfilesets As SolidEdgePart.ProfileSets = Nothing

        objApp = Marshal.GetActiveObject("SolidEdge.Application")
        ObjPartDoc = objApp.ActiveDocument

        Dim models As SolidEdgePart.Models = Nothing
        Dim model As SolidEdgePart.Model = Nothing
        models = ObjPartDoc.Models
        model = models.Item(1)
        objProfilesets = ObjPartDoc.ProfileSets
        Dim sketchfullydefined As String = "Yes"
        For Each profileset As ProfileSet In objProfilesets

            Dim isunderdefined As Boolean = profileset.IsUnderDefined

            If isunderdefined = True Then
                sketchfullydefined = "No"
                Return sketchfullydefined
            End If

        Next

        Return sketchfullydefined
    End Function

    Private Sub MTCReviewForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TestAssemblyFeatureExistence()
        'TestPartConstructionFeature()
        'TestReferencePart()
        'TestAssemblyFeature()
        'Me.Close()
        'Exit Sub

        Dim aPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location
        Dim dirPath As String = IO.Path.GetDirectoryName(aPath)
        Dim mtcExcelPath As String = IO.Path.Combine(dirPath, "M2MData.xlsx")
        dicData = ExcelUtil.ReadM2Mfile(mtcExcelPath)
        Dim ds As DataSet = dicData("Sheet1")
        dtM2M = ds.Tables(0)

        SetControls(dgvDocumentDetails)

        Me.Text = $"{Me.Text} ({GlobalEntity.Version}) "

        Dim Propseedfilepath As String = System.Reflection.Assembly.GetExecutingAssembly.Location
        Dim Propseeddirpath As String = IO.Path.GetDirectoryName(aPath)
        Dim propseedfile As String = IO.Path.Combine(dirPath, "propseed.txt")
        projectnamelst = readfile(propseedfile)
        Authorlst = readfileAuthor(propseedfile)
    End Sub

    Private Sub TestPartConstructionFeature()
        objApp = Marshal.GetActiveObject("SolidEdge.Application")
        Dim partDoc As PartDocument = objApp.ActiveDocument


        Dim copyConstructionCount As Integer = partDoc.Constructions.CopyConstructions.Count
        'For index = 0 To constructionsObj.Count - 1
        '    Dim constructionModelObj As ConstructionModel = constructionsObj.Item(index)
        '    Dim constructionModelName As String = constructionModelObj.BodyName
        'Next
    End Sub

    Private Sub TestAssemblyFeatureExistence()

        Try
            objApp = Marshal.GetActiveObject("SolidEdge.Application")

            Dim asssemblyDocObj As AssemblyDocument = objApp.ActiveDocument
            Dim asmModelObj As SolidEdgePart.Model = asssemblyDocObj.AssemblyModel
            Dim asmModelFeatures As Features = asmModelObj.Features
            Dim cnt As Integer = asmModelFeatures.Count
            MsgBox(cnt.ToString())


        Catch ex As Exception

        End Try

    End Sub

    Private Function CheckInterPartLinksPSM() As String
        Dim resCheckInterPartLinkBroken As String = "No"
        Try
            Dim sheetMetalDoc As SheetMetalDocument = objApp.ActiveDocument
            Dim resHasInterPartLinks As Boolean
            sheetMetalDoc.HasInterpartLinks(resHasInterPartLinks)
            If resHasInterPartLinks Then
                resCheckInterPartLinkBroken = "No"
            Else
                resCheckInterPartLinkBroken = "Yes"
            End If
        Catch ex As Exception

        End Try

        Return resCheckInterPartLinkBroken
    End Function

    Private Function CheckInterPartLinksPAR() As String
        Dim resCheckInterPartLinkBroken As String = "No"
        'Dim sheetMetalDoc As PartDocument = objApp.ActiveDocument
        'Dim resHasInterPartLinks As Boolean
        'sheetMetalDoc.HasInterpartLinks(resHasInterPartLinks)
        'If resHasInterPartLinks Then
        '    resCheckInterPartLinkBroken = "No"
        'Else
        '    resCheckInterPartLinkBroken = "Yes"
        'End If
        Return resCheckInterPartLinkBroken
    End Function

    Private Function CheckInterPartLinksASM() As String
        Dim resCheckInterPartLinkBroken As String = "No"
        Dim sheetMetalDoc As AssemblyDocument = objApp.ActiveDocument
        Dim resHasInterPartLinks As Boolean
        sheetMetalDoc.HasInterpartLinks(resHasInterPartLinks)
        If resHasInterPartLinks Then
            resCheckInterPartLinkBroken = "No"
        Else
            resCheckInterPartLinkBroken = "Yes"
        End If
        Return resCheckInterPartLinkBroken
    End Function

    Private Sub TestReferencePart()
        objApp = Marshal.GetActiveObject("SolidEdge.Application")
        Dim sheetMetalDoc As SheetMetalDocument = objApp.ActiveDocument
        Dim resHasInterPartLinks As Boolean
        sheetMetalDoc.HasInterpartLinks(resHasInterPartLinks) 'InterpartLinks

        'Dim bodyMembersCnt As Integer
        'Dim bodyMembers As Array
        'Dim assemblyExists As Boolean
        'Dim assemblyFileName As String
        Dim asmStatus As MultiBodyPublishStatusConstants = MultiBodyPublishStatusConstants.seMBPStatusUnknown
        'sheetMetalDoc.GetMultiBodyPublishMembers(bodyMembersCnt, bodyMembers, assemblyExists, assemblyFileName, asmStatus)
        Dim pObj As Object = Nothing
        Dim noOdPar As Integer
        Dim noOfDepen As Integer
        sheetMetalDoc.GetNumberOfParentsAndDependents(pObj, noOdPar, noOfDepen)
        Debug.Print("aaaa")
        'For Each part As Part In assemblyParts
        '    Dim a = part.ReferenceOnly
        '    Dim b = part.Relations3d

        'Next
    End Sub

    Private Sub TestAssemblyFeature()
        objApp = Marshal.GetActiveObject("SolidEdge.Application")
        objAssemblyDocument = objApp.ActiveDocument

        Dim assemblyDrivenPartFeatures As AssemblyDrivenPartFeatures = objAssemblyDocument.AssemblyDrivenPartFeatures

        Dim extCutOut As AssemblyDrivenPartFeaturesExtrudedCutouts = assemblyDrivenPartFeatures.AssemblyDrivenPartFeaturesExtrudedCutouts
        If extCutOut.Count > 0 Then
            MsgBox("Feature Exists")
        End If

        Dim assemblyDrivenPartFeatures1 As AssemblyFeatures = objAssemblyDocument.AssemblyFeatures

        Dim extCutOut1 As AssemblyDrivenPartFeaturesExtrudedCutouts = assemblyDrivenPartFeatures1.AssemblyFeaturesExtrudedCutouts
        For Each a As AssemblyFeaturesExtrudedCutout In extCutOut1
            MsgBox("Feature Exists@@@")
        Next
        If extCutOut1.Count > 0 Then
            MsgBox("Feature Exists")
        End If
    End Sub

    Private Sub SetControls(ByRef DataGridViewComments As DataGridView)
        For Each col As DataGridViewColumn In DataGridViewComments.Columns
            col.SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        DataGridViewComments.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.GhostWhite
        DataGridViewComments.AlternatingRowsDefaultCellStyle.ForeColor = System.Drawing.Color.DarkBlue

        DataGridViewComments.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(6, 150, 215) ' Drawing.Color.WhiteSmoke
        DataGridViewComments.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.WhiteSmoke 'Drawing.Color.OrangeRed 'FromArgb(6, 150, 215) '

        DataGridViewComments.AllowUserToAddRows = False
        DataGridViewComments.AllowUserToDeleteRows = False
        DataGridViewComments.AllowUserToResizeRows = False
        DataGridViewComments.AllowUserToResizeColumns = False
        DataGridViewComments.RowHeadersVisible = False
        DataGridViewComments.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect 'CellSelect
        DataGridViewComments.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        'DataGridViewComments.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        DataGridViewComments.EnableHeadersVisualStyles = True
        DataGridViewComments.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
        DataGridViewComments.MultiSelect = False
        DataGridViewComments.ReadOnly = False
    End Sub

    Private Sub btnGetData_Click(sender As Object, e As EventArgs) Handles btnGetCurrentAssembly.Click
        waitStartSave()
        objApp = Marshal.GetActiveObject("SolidEdge.Application")
        objAssemblyDocument = objApp.ActiveDocument

        'dgvDocumentDetails.Columns.Add(dr)
        dtAssemblyData = getpartlist()

        'DataTable distinctTable = originalTable.DefaultView.ToTable( /*distinct*/ true);
        dtAssemblyData = dtAssemblyData.DefaultView.ToTable(True)
        dgvDocumentDetails.RowHeadersVisible = False
        dgvDocumentDetails.DataSource = dtAssemblyData

        FillSearchCombo()
        btn_Alldata_Click(sender, e)
        WaitEndSave()
    End Sub

    Private Sub FillSearchCombo()
        ComboBoxFields.Items.Clear()
        ComboBoxFields.Items.Add("Select")

        For Each columns As System.Windows.Forms.DataGridViewColumn In dgvDocumentDetails.Columns
            ComboBoxFields.Items.Add(columns.Name.ToString())
        Next
        If ComboBoxFields.Items.Count > 0 Then
            ComboBoxFields.SelectedItem = ComboBoxFields.Items(0) 'ComboBoxFields.Items.Count - 1)
        End If
    End Sub

    Private Sub ComboBoxFields_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxFields.SelectedIndexChanged
        Try
            dgvDocumentDetails.DataSource = dtfilter
            SetAutoComplete_CSV()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("Auto complete." + vbNewLine + ex.Message, "Search", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub SetAutoComplete_CSV()
        Try
            txtSearch.Text = String.Empty
            Dim lst As New List(Of String)
            Dim MySource As New AutoCompleteStringCollection()

            If ComboBoxFields.SelectedIndex > 0 Then
                For Each outerList As DataGridViewRow In dgvDocumentDetails.Rows
                    lst.Add(outerList.Cells((ComboBoxFields.SelectedIndex) - 1).Value.ToString()) '.Cells(ComboBoxFields.SelectedIndex))
                Next
            End If

            lst = lst.Distinct().ToList()
            MySource.AddRange(lst.ToArray)
            txtSearch.AutoCompleteCustomSource = MySource
            txtSearch.AutoCompleteMode = AutoCompleteMode.SuggestAppend
            txtSearch.AutoCompleteSource = AutoCompleteSource.CustomSource
            'lst.Clear()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtSearch_KeyUp(sender As Object, e As KeyEventArgs) Handles txtSearch.KeyUp
        Try
            If (e.KeyValue = Keys.Enter) Then
                btnSearchFile_Click(sender, e)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSearchFile_Click(sender As Object, e As EventArgs) Handles btnSearchFile.Click
        TestRemoveCode()
        TestRemoveCode()
    End Sub

    Private Sub TestRemoveCode()
        Try

            If txtSearch.Text.Trim = String.Empty Then
                ' dgvDocumentDetails.DataSource = Nothing
                dgvDocumentDetails.DataSource = dtfilter
                'Exit Sub
            End If

            If txtSearch.Text = String.Empty Then
                'Dim dt As DataTable = DataGridViewComments.DataSource
                'Dim DV As DataView = New DataView(dt)
                'Try
                '    DV.RowFilter = String.Format("" + ComboBoxFields.Text + " = ''")
                'Catch ex As Exception
                'End Try
                'DataGridViewComments.DataSource = DV
            Else
                Dim dt As System.Data.DataTable = dtfilter
                Dim DV As DataView = New DataView(dt)
                Try
                    DV.RowFilter = String.Format("[" + ComboBoxFields.Text + "] LIKE '%{0}%'", txtSearch.Text)
                Catch ex As Exception
                    DV.RowFilter = String.Format("[" + ComboBoxFields.Text + "]={0}", txtSearch.Text)
                End Try
                dgvDocumentDetails.DataSource = DV.ToTable()
            End If

        Catch ex As Exception
            MessageBox.Show("Unable to search." + vbNewLine + ex.Message, "Search", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

        End Try
    End Sub

    Private Sub btnAssemblyCheck_Click(sender As Object, e As EventArgs) Handles btnAssemblyCheck.Click

        Dim extension As String = ".asm"
        Documentfilter(extension)

    End Sub

    Public Sub Documentfilter(ByVal ext As String)
        'Dim ext As String = extenstion
        Dim dt As System.Data.DataTable = dtAssemblyData
        Dim DV As DataView = New DataView(dt)
        Try


            DV.RowFilter = "[File Name (full path)] LIKE '%" + ext + "%' or [File Name (full path)] LIKE '%" + ext.ToUpper() + "%'" ' String.Format("" + "File Name (full path)" + " LIKE '%{0}%'", a)
        Catch ex As Exception
            ' DV.RowFilter = String.Format("" + ComboBoxFields.Text + "={0}", txtSearch.Text)
        End Try
        dgvDocumentDetails.DataSource = DV.ToTable()
        dtfilter = DV.ToTable
    End Sub

    Private Sub btnPartCheck_Click(sender As Object, e As EventArgs) Handles btnPartCheck.Click
        Dim extension As String = ".par"
        Documentfilter(extension)
    End Sub

    Private Sub btnSheetMetalCheck_Click(sender As Object, e As EventArgs) Handles btnSheetMetalCheck.Click
        Dim extension As String = ".psm"
        Documentfilter(extension)
    End Sub

    Private Sub btn_Alldata_Click(sender As Object, e As EventArgs) Handles btn_Alldata.Click
        btnGetCurrentAssembly.Enabled = False
        dgvDocumentDetails.DataSource = dtAssemblyData
    End Sub

    Dim waitFormObj As Wait

    Public Sub waitStartSave()
        '==Processing==
        Dim waitThread As System.Threading.Thread
        waitThread = New System.Threading.Thread(AddressOf launchWaitSave)
        waitThread.Start()
        Threading.Thread.Sleep(1000)
        waitFormObj.SetWaitMessage("In progress..")

        waitFormObj.SetProgressInformationVisibility(True)
        waitFormObj.SetProgressInformationMessage("")

        waitFormObj.SetProgressCountVisibility(True)
        waitFormObj.SetProgressCountMessage("0/0")
        '
    End Sub

    Public Sub launchWaitSave()
        waitFormObj = New Wait()
        waitFormObj.ShowDialog()
    End Sub

    Public Sub WaitEndSave()
        If Not waitFormObj Is Nothing Then
            waitFormObj.dispose2()
            waitFormObj = Nothing
        End If
    End Sub

    Private Function GetBECAuthorAssemblyData() As System.Data.DataTable

        Dim filterStr As String = String.Empty
        Dim length As Int16 = Authorlst.Count
        Dim cnt As Integer = 1
        For Each str As String In Authorlst
            If cnt < length Then
                filterStr = filterStr + " Author = '" + str + "' Or "
            Else
                filterStr = filterStr + " Author = '" + str + "'"
            End If
            cnt = cnt + 1

        Next
        filterStr = filterStr.Trim()

        Dim dtAssemblyData As System.Data.DataTable = DirectCast(dgvDocumentDetails.DataSource, System.Data.DataTable)
        Dim dv As DataView = New DataView(dtAssemblyData)
        dv.RowFilter = filterStr
        Dim dtBECAuthor As System.Data.DataTable = dv.ToTable()
        Return dtBECAuthor
    End Function

    Private Function GetNonBECAuthorAssemblyData() As System.Data.DataTable
        Dim filterStr As String = String.Empty
        Dim length As Int16 = Authorlst.Count
        Dim cnt As Integer = 1
        For Each str As String In Authorlst
            If cnt < length Then
                filterStr = filterStr + " Author <> '" + str + "' And "
            Else
                filterStr = filterStr + " Author <> '" + str + "'"
            End If
            cnt = cnt + 1

        Next
        filterStr = filterStr.Trim()

        Dim dtAssemblyData As System.Data.DataTable = DirectCast(dgvDocumentDetails.DataSource, System.Data.DataTable)
        Dim dv As DataView = New DataView(dtAssemblyData)
        dv.RowFilter = filterStr
        Dim dtNonBECAuthor As System.Data.DataTable = dv.ToTable()
        Return dtNonBECAuthor
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles ButtonExportExcelMTC.Click

        log.Info($"MTC report Start")
        log.Info($"======================================")

        Dim dtBECAuthor As System.Data.DataTable = GetBECAuthorAssemblyData()
        Dim dtNonBECAuthor As System.Data.DataTable = GetNonBECAuthorAssemblyData()
        Dim exportDirectoryLocation As String = GetExportDirectoryLocation()
        Dim baseLineDirectoryLocation As String = txtBaseLineDirPath.Text

        Dim dtData As System.Data.DataTable = DirectCast(dgvDocumentDetails.DataSource, System.Data.DataTable)
        waitStartSave()
        killTasks()
        'SaveAsExcel2(dtBECAuthor, exportDirectoryLocation, "BEC")

        SaveAsExcelSplit_MTC(dtBECAuthor, exportDirectoryLocation, baseLineDirectoryLocation, "BEC")
        killTasks()
        'SaveAsExcel2(dtNonBECAuthor, exportDirectoryLocation, "DGS")

        SaveAsExcelSplit_MTC(dtNonBECAuthor, exportDirectoryLocation, baseLineDirectoryLocation, "DGS")
        killTasks()
        'SaveAsExcel2(dtData, exportDirectoryLocation, "MTC")

        WaitEndSave()

        log.Info($"======================================")
        log.Info($"MTC report End")


    End Sub
    Private Shared Sub killTasks()

        Try
            Dim pProcess() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
            For Each p As Process In pProcess
                p.Kill()
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            Dim UnCheckedItems = From Rows In dgvDocumentDetails.Rows.Cast(Of DataGridViewRow)()
                                 Where CBool(Rows.Cells("Select").Value) = False


            For Each item In UnCheckedItems
                item.Cells("Select").Value = True

            Next
        Else
            Dim UnCheckedItems = From Rows In dgvDocumentDetails.Rows.Cast(Of DataGridViewRow)()
                                 Where CBool(Rows.Cells("Select").Value) = True


            For Each item In UnCheckedItems
                item.Cells("Select").Value = False
            Next
        End If
    End Sub

    Public Function readfile(ByVal file As String) As List(Of String)

        Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(file)
        Dim a As String

        Dim lst As List(Of String) = New List(Of String)
        Do
            a = reader.ReadLine
            If lst.Contains("Begin Project") Then
                Dim split As String() = a.Split(";")

                lst.Add(split(0))
            End If
            If a = "Begin Project" Then
                Dim split As String() = a.Split(";")

                lst.Add(split(0))
            End If

            '
            ' Code here
            '
        Loop Until a = "End Project"

        reader.Close()
        Return lst



    End Function

    Public Function readfileAuthor(ByVal file As String) As List(Of String)

        Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(file)
        Dim a As String

        Dim lst As List(Of String) = New List(Of String)
        Do
            a = reader.ReadLine
            If lst.Contains("Begin Author") Then
                Dim split As String() = a.Split(";")

                lst.Add(split(0))
            End If
            If a = "Begin Author" Then
                Dim split As String() = a.Split(";")

                lst.Add(split(0))
            End If

            '
            ' Code here
            '
        Loop Until a = "End Author"

        reader.Close()
        Return lst



    End Function

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles ButtonClose.Click
        Me.Close()
    End Sub

    Private Sub BtnExportExcelMTR_Click(sender As Object, e As EventArgs) Handles BtnExportExcelMTR.Click
        log.Info($"MTR report Start")
        log.Info($"======================================")

        Dim dtBECAuthor As System.Data.DataTable = GetBECAuthorAssemblyData()

        Dim dtNonBECAuthor As System.Data.DataTable = GetNonBECAuthorAssemblyData()

        Dim exportDirectoryLocation As String = GetExportDirectoryLocation()

        Dim dtData As System.Data.DataTable = DirectCast(dgvDocumentDetails.DataSource, System.Data.DataTable)

        waitStartSave()

        killTasks()

        SaveAsExcelSplit_MTR(dtBECAuthor, exportDirectoryLocation, "BEC")

        killTasks()

        SaveAsExcelSplit_MTR(dtNonBECAuthor, exportDirectoryLocation, "DGS")

        killTasks()

        WaitEndSave()

        log.Info($"======================================")
        log.Info($"MTR report End")


    End Sub

    Private Sub btnBrowseBaselinePath_Click(sender As Object, e As EventArgs) Handles btnBrowseBaselinePath.Click
        txtBaseLineDirPath.Text = GetExportDirectoryLocation()
    End Sub
End Class