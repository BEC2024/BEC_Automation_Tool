Imports System.IO
Imports System.Runtime.InteropServices
Imports ExcelDataReader
Imports NLog
Imports SolidEdgeDraft

Public Class MTC_MTR_BL
    Public Function ReadAuthorExcelData(ByVal dt As DataTable) As DataTable
        Dim Path = Config.configObj.EmployeeExcelPath
        If (System.IO.File.Exists(Path)) Then
            Dim ds As New DataSet("Data")

            Using stream As FileStream = File.Open(Path, FileMode.Open, FileAccess.Read)
                Using reader As IExcelDataReader = ExcelReaderFactory.CreateOpenXmlReader(stream)
                    Dim conf = New ExcelDataSetConfiguration With
                        {
                            .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration With
                            {
                                .UseHeaderRow = True
                            }
                        }
                    ds = reader.AsDataSet(conf)


                End Using
            End Using
            dt = ds.Tables(0)

        Else
            MessageBox.Show("3 MTC_MTR -- Employee Excel Path Missing or not Exist in Config")
            Exit Function
        End If

        Return dt
    End Function

    Public Function GetBOMCount(ByRef dtCurrentAssemblyData As DataTable) As Integer

        Dim Count As Integer = dtCurrentAssemblyData.Rows.Count() - 1
        Dim value As String
        Dim BOMCount As Integer
        Dim Max As Integer = 0
        For i = 0 To Count
            value = dtCurrentAssemblyData.Rows(i)(1)
            If value.ToString.Contains("*") Then
                value = value.ToString.Replace("*", "")
                value.ToString.Trim()
            End If
            BOMCount = Convert.ToInt32(value)
            If BOMCount > Max Then
                Max = BOMCount
            End If
        Next
        BOMCount = Max
        Return BOMCount
    End Function

    Public Function GetCurrentAssemblyData() As System.Data.DataTable
        '9th Sep 2024
        Try
            Dim objApp As SolidEdgeFramework.Application = Nothing
            Dim objAssemblyDocument As SolidEdgeAssembly.AssemblyDocument = Nothing
            Dim dtAssemblyData As System.Data.DataTable

            objApp = Marshal.GetActiveObject("SolidEdge.Application")

            objAssemblyDocument = objApp.ActiveDocument

            dtAssemblyData = GetPartList()



            dtAssemblyData = dtAssemblyData.DefaultView.ToTable(True)

            Return dtAssemblyData
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Function


    '17th Sep 2024
    Public Function GetCurrentPartData() As System.Data.DataTable
        '9th Sep 2024
        Try
            Dim objApp As SolidEdgeFramework.Application = Nothing
            Dim objPartDocument As SolidEdgePart.PartDocument = Nothing
            Dim dtAssemblyData As System.Data.DataTable

            objApp = Marshal.GetActiveObject("SolidEdge.Application")

            objPartDocument = objApp.ActiveDocument

            dtAssemblyData = GetPartListPartDoc()

            dtAssemblyData = dtAssemblyData.DefaultView.ToTable(True)

            Return dtAssemblyData
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Function


    '17th Sep 2024
    Public Function GetCurrentSheetMetalData() As System.Data.DataTable
        '9th Sep 2024
        Try
            Dim objApp As SolidEdgeFramework.Application = Nothing
            Dim objSheetMetalDocument As SolidEdgePart.SheetMetalDocument = Nothing
            Dim dtAssemblyData As System.Data.DataTable

            objApp = Marshal.GetActiveObject("SolidEdge.Application")

            objSheetMetalDocument = objApp.ActiveDocument

            dtAssemblyData = GetPartListSheetMetalDoc()

            dtAssemblyData = dtAssemblyData.DefaultView.ToTable(True)

            Return dtAssemblyData
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Function


    '17th Sep 2024
    Private Function GetPartListPartDoc() As System.Data.DataTable

        Dim objApp As SolidEdgeFramework.Application = Nothing
        Dim objAssemblyDocument As SolidEdgeAssembly.AssemblyDocument = Nothing
        Dim objDocuments As SolidEdgeFramework.Documents = Nothing
        Dim objDraft As SolidEdgeDraft.DraftDocument = Nothing
        Dim objSheet As SolidEdgeDraft.Sheet = Nothing
        Dim objModelLinks As SolidEdgeDraft.ModelLinks = Nothing
        Dim objModelLink As SolidEdgeDraft.ModelLink = Nothing
        Dim objDrawingViews As SolidEdgeDraft.DrawingViews = Nothing
        Dim objDrawingView As SolidEdgeDraft.DrawingView = Nothing

        '17th Sep 2024
        Dim objPartDocument As SolidEdgePart.PartDocument = Nothing

        Try
            '  OleMessageFilter.Register()

            ' Connect to a running instance of Solid Edge
            objApp = Marshal.GetActiveObject("SolidEdge.Application")

            objPartDocument = objApp.ActiveDocument

            ' Get a reference to the documents collection
            objDocuments = objApp.Documents

            ' Add a Draft document
            objDraft = objDocuments.Add("SolidEdge.DraftDocument")

            ' Get a reference to the active sheet
            objSheet = objDraft.ActiveSheet

            ' Get a reference to the model links collection
            objModelLinks = objDraft.ModelLinks
            Dim filename As String
            Dim file As String = objPartDocument.FullName
            filename = System.IO.Path.Combine(SolidEdgeCommunity.SolidEdgeUtils.GetTrainingFolderPath(), file)

            ' Add a new model link
            objModelLink = objModelLinks.Add(filename)

            ' Get a reference to the drawing views collection
            objDrawingViews = objSheet.DrawingViews

            objDrawingView = objDrawingViews.AddPartView([From]:=objModelLink, Orientation:=SolidEdgeDraft.ViewOrientationConstants.igDimetricTopBackLeftView, Scale:=1.0, x:=0.4, y:=0.4, ViewType:=SolidEdgeDraft.AssemblyDrawingViewTypeConstants.seAssemblyDesignedView)

#Region "Add drawing view"

            ' Add a new drawing view
            'objDrawingView = objDrawingViews.AddAssemblyView(
            '  objModelLink,
            '  SolidEdgeDraft.ViewOrientationConstants.igFrontView,
            '  1,
            '  0.1,
            '  0.1,
            '  SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView)

#End Region

            ' Assign a caption
            objDrawingView.Caption = "My New Drawing View"

            ' Ensure caption is displayed
            objDrawingView.DisplayCaption = False
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            MTC_MTR_ReviewForm2.log.Error($"{ex.Message}{vbNewLine}{ex.StackTrace}")
        Finally
            ' OleMessageFilter.Revoke()
        End Try

        '  Dim objApp As SolidEdgeFramework.Application = Nothing
        Dim objDoc As SolidEdgeDraft.DraftDocument = Nothing
        Dim objPartsLists As SolidEdgeDraft.PartsLists = Nothing
        Dim objPartsList As SolidEdgeDraft.PartsList = Nothing

        objDoc = objApp.ActiveDocument
        objPartsLists = objDoc.PartsLists
        objPartsList = objPartsLists.Add(objDrawingView, "BEC", 1, 1)

        objPartsList = objPartsLists.Item(1)
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

        Dim dtrows As DataRow = dt.NewRow()
        dt.Rows.Add(dtrows)

        For Each tableRow In rows.OfType(Of SolidEdgeDraft.TableRow)()
            For Each tableColumn In cols.OfType(Of SolidEdgeDraft.TableColumn)()
                If tableColumn.Show Then
                    tableCell = objPartsList.Cell(tableRow.Index, tableColumn.Index)
                    Dim rowindex As Integer = tableRow.Index - 1
                    Dim colindex As Integer = tableColumn.Index - 1 + 1
                    Dim tabvalue As String = tableCell.value

                    dt.Rows(rowindex).Item(colindex) = tabvalue

                End If
            Next tableColumn
            If Not tableRow.Index = rows.Count Then
                dtrows = dt.NewRow()
                dt.Rows.Add(dtrows)

            End If

        Next tableRow

        objApp.Documents.CloseDocument(objDraft.FullName, False, "", False, False)

        Return dt

    End Function


    '17th Sep 2024
    Private Function GetPartListSheetMetalDoc() As System.Data.DataTable

        Dim objApp As SolidEdgeFramework.Application = Nothing
        Dim objAssemblyDocument As SolidEdgeAssembly.AssemblyDocument = Nothing
        Dim objDocuments As SolidEdgeFramework.Documents = Nothing
        Dim objDraft As SolidEdgeDraft.DraftDocument = Nothing
        Dim objSheet As SolidEdgeDraft.Sheet = Nothing
        Dim objModelLinks As SolidEdgeDraft.ModelLinks = Nothing
        Dim objModelLink As SolidEdgeDraft.ModelLink = Nothing
        Dim objDrawingViews As SolidEdgeDraft.DrawingViews = Nothing
        Dim objDrawingView As SolidEdgeDraft.DrawingView = Nothing

        '17th Sep 2024
        Dim objSheetMetalDocument As SolidEdgePart.SheetMetalDocument = Nothing

        Try
            '  OleMessageFilter.Register()

            ' Connect to a running instance of Solid Edge
            objApp = Marshal.GetActiveObject("SolidEdge.Application")

            objSheetMetalDocument = objApp.ActiveDocument

            ' Get a reference to the documents collection
            objDocuments = objApp.Documents

            ' Add a Draft document
            objDraft = objDocuments.Add("SolidEdge.DraftDocument")

            ' Get a reference to the active sheet
            objSheet = objDraft.ActiveSheet

            ' Get a reference to the model links collection
            objModelLinks = objDraft.ModelLinks
            Dim filename As String
            Dim file As String = objSheetMetalDocument.FullName
            filename = System.IO.Path.Combine(SolidEdgeCommunity.SolidEdgeUtils.GetTrainingFolderPath(), file)

            ' Add a new model link
            objModelLink = objModelLinks.Add(filename)

            ' Get a reference to the drawing views collection
            objDrawingViews = objSheet.DrawingViews

            objDrawingView = objDrawingViews.AddSheetMetalView([From]:=objModelLink, Orientation:=SolidEdgeDraft.ViewOrientationConstants.igDimetricTopBackLeftView, Scale:=1.0, x:=0.4, y:=0.4, ViewType:=SolidEdgeDraft.AssemblyDrawingViewTypeConstants.seAssemblyDesignedView)

#Region "Add drawing view"

            ' Add a new drawing view
            'objDrawingView = objDrawingViews.AddAssemblyView(
            '  objModelLink,
            '  SolidEdgeDraft.ViewOrientationConstants.igFrontView,
            '  1,
            '  0.1,
            '  0.1,
            '  SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView)

#End Region

            ' Assign a caption
            objDrawingView.Caption = "My New Drawing View"

            ' Ensure caption is displayed
            objDrawingView.DisplayCaption = False
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            MTC_MTR_ReviewForm2.log.Error($"{ex.Message}{vbNewLine}{ex.StackTrace}")
        Finally
            ' OleMessageFilter.Revoke()
        End Try

        '  Dim objApp As SolidEdgeFramework.Application = Nothing
        Dim objDoc As SolidEdgeDraft.DraftDocument = Nothing
        Dim objPartsLists As SolidEdgeDraft.PartsLists = Nothing
        Dim objPartsList As SolidEdgeDraft.PartsList = Nothing

        objDoc = objApp.ActiveDocument
        objPartsLists = objDoc.PartsLists
        objPartsList = objPartsLists.Add(objDrawingView, "BEC", 1, 1)

        objPartsList = objPartsLists.Item(1)
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

        Dim dtrows As DataRow = dt.NewRow()
        dt.Rows.Add(dtrows)

        For Each tableRow In rows.OfType(Of SolidEdgeDraft.TableRow)()
            For Each tableColumn In cols.OfType(Of SolidEdgeDraft.TableColumn)()
                If tableColumn.Show Then
                    tableCell = objPartsList.Cell(tableRow.Index, tableColumn.Index)
                    Dim rowindex As Integer = tableRow.Index - 1
                    Dim colindex As Integer = tableColumn.Index - 1 + 1
                    Dim tabvalue As String = tableCell.value

                    dt.Rows(rowindex).Item(colindex) = tabvalue

                End If
            Next tableColumn
            If Not tableRow.Index = rows.Count Then
                dtrows = dt.NewRow()
                dt.Rows.Add(dtrows)

            End If

        Next tableRow

        objApp.Documents.CloseDocument(objDraft.FullName, False, "", False, False)

        Return dt

    End Function

    Private Function GetPartList() As System.Data.DataTable
        Dim objApp As SolidEdgeFramework.Application = Nothing
        Dim objAssemblyDocument As SolidEdgeAssembly.AssemblyDocument = Nothing
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
            objApp = Marshal.GetActiveObject("SolidEdge.Application")

            objAssemblyDocument = objApp.ActiveDocument

            ' Get a reference to the documents collection
            objDocuments = objApp.Documents

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

#Region "Add drawing view"

            ' Add a new drawing view
            'objDrawingView = objDrawingViews.AddAssemblyView(
            '  objModelLink,
            '  SolidEdgeDraft.ViewOrientationConstants.igFrontView,
            '  1,
            '  0.1,
            '  0.1,
            '  SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView)

#End Region

            ' Assign a caption
            objDrawingView.Caption = "My New Drawing View"

            ' Ensure caption is displayed
            objDrawingView.DisplayCaption = False
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            MTC_MTR_ReviewForm2.log.Error($"{ex.Message}{vbNewLine}{ex.StackTrace}")
        Finally
            ' OleMessageFilter.Revoke()
        End Try

        '  Dim objApp As SolidEdgeFramework.Application = Nothing
        Dim objDoc As SolidEdgeDraft.DraftDocument = Nothing
        Dim objPartsLists As SolidEdgeDraft.PartsLists = Nothing
        Dim objPartsList As SolidEdgeDraft.PartsList = Nothing

        objDoc = objApp.ActiveDocument
        objPartsLists = objDoc.PartsLists
        objPartsList = objPartsLists.Add(objDrawingView, "BEC", 1, 1)

        objPartsList = objPartsLists.Item(1)
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

        Dim dtrows As DataRow = dt.NewRow()
        dt.Rows.Add(dtrows)

        For Each tableRow In rows.OfType(Of SolidEdgeDraft.TableRow)()
            For Each tableColumn In cols.OfType(Of SolidEdgeDraft.TableColumn)()
                If tableColumn.Show Then
                    tableCell = objPartsList.Cell(tableRow.Index, tableColumn.Index)
                    Dim rowindex As Integer = tableRow.Index - 1
                    Dim colindex As Integer = tableColumn.Index - 1 + 1
                    Dim tabvalue As String = tableCell.value

                    dt.Rows(rowindex).Item(colindex) = tabvalue

                End If
            Next tableColumn
            If Not tableRow.Index = rows.Count Then
                dtrows = dt.NewRow()
                dt.Rows.Add(dtrows)

            End If

        Next tableRow

        objApp.Documents.CloseDocument(objDraft.FullName, False, "", False, False)

        Return dt

    End Function

    Public Function GetBECAuthorAssemblyData(ByVal mtcMtrModelObj As MTC_MTR_Model) As System.Data.DataTable
        '9th Sep 2024 'Added try..catch
        Try
            Dim Authorlst As List(Of String) = mtcMtrModelObj.authorList
            Dim dtAssemblyData As System.Data.DataTable = mtcMtrModelObj.dtFilteredAssemblyData
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

            'Dim dtAssemblyData As System.Data.DataTable = DirectCast(dgvDocumentDetails.DataSource, System.Data.DataTable)
            Dim dv As DataView = New DataView(dtAssemblyData)
            dv.RowFilter = filterStr
            Dim dtBECAuthor As System.Data.DataTable = dv.ToTable()

            Return dtBECAuthor
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

    Public Function GetNonBECAuthorAssemblyData(ByVal mtcMtrModelObj As MTC_MTR_Model) As System.Data.DataTable
        '9th Sep 2024 'Added try..catch
        Try
            Dim Authorlst As List(Of String) = mtcMtrModelObj.authorList
            Dim dtAssemblyData As System.Data.DataTable = mtcMtrModelObj.dtFilteredAssemblyData

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

            'Dim dtAssemblyData As System.Data.DataTable = DirectCast(dgvDocumentDetails.DataSource, System.Data.DataTable)
            Dim dv As DataView = New DataView(dtAssemblyData)
            dv.RowFilter = filterStr
            Dim dtNonBECAuthor As System.Data.DataTable = dv.ToTable()
            Return dtNonBECAuthor
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

End Class