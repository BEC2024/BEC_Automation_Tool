Imports System.IO
Imports System.Runtime.InteropServices
Imports ExcelDataReader
Imports SolidEdgeCommunity
Imports SolidEdgeDraft
Imports SolidEdgeFileProperties

Public Class Author_Updation
    Dim objApp As SolidEdgeFramework.Application
    Dim _applicationEvents As SolidEdgeFramework.ISEApplicationEvents_Event
    Dim Objdocuemnts As SolidEdgeFramework.Documents = Nothing
    Dim ObjdraftDoc As SolidEdgeDraft.DraftDocument = Nothing
    Dim ObjAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = Nothing
    Dim ObjPartdoc As SolidEdgePart.PartDocument = Nothing
    Dim ObjSheetMetalDoc As SolidEdgePart.SheetMetalDocument = Nothing
    Dim Objsheet As SolidEdgeDraft.Sheet = Nothing
    Dim ObjdrwViews As SolidEdgeDraft.DrawingViews = Nothing
    Dim ObjdrwView As SolidEdgeDraft.DrawingView = Nothing
    Dim ObjpartLists As SolidEdgeDraft.PartsLists = Nothing
    Dim ObjpartList As SolidEdgeDraft.PartsList = Nothing
    Dim objModelLink As SolidEdgeDraft.ModelLink = Nothing
    Dim objModelLinks As SolidEdgeDraft.ModelLinks = Nothing
    Dim ObjTableCell As SolidEdgeDraft.TableCell = Nothing
    Dim ObjCols As SolidEdgeDraft.TableColumns = Nothing
    Dim ObjRows As SolidEdgeDraft.TableRows = Nothing
    Dim objDoc As SolidEdgeAssembly.AssemblyDocument
    Dim objParts As SolidEdgeAssembly.Occurrences

    Dim PartdocPath As String = String.Empty
    Dim SheetMetalDocPath As String = String.Empty
    Dim AssemblyDocPath As String = String.Empty
    Dim DocumentPath1 As String = Nothing
    Dim activeDocPath As String
    Dim MainAssemblyPath As String
    Dim dt As DataTable = New DataTable()
    Private Sub Author_Updation_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtEmployeeExcelPath.Text = Config.configObj.EmployeeExcelPath
        CheckSEInstance()
        _applicationEvents = CType(objApp.ApplicationEvents, SolidEdgeFramework.ISEApplicationEvents_Event)

        AddHandler _applicationEvents.AfterDocumentSave, AddressOf _applicationEvents_AfterDocumentSave
        AddHandler _applicationEvents.BeforeDocumentSave, AddressOf _applicationEvents_BeforeDocumentSave

        Dim dtUsers As DataTable = ReadAuthorExcelData(dtUsers)
        dt = dtUsers
    End Sub

    Private Sub _applicationEvents_AfterDocumentSave(ByVal theDocument As Object)
        ' Code to execute after a document is saved


    End Sub

    Private Sub _applicationEvents_BeforeDocumentSave(ByVal theDocument As Object)
        Try
            UpdateSummaryInfoAuthor(dt)
        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        Dim IsActive = CheckSEInstance()
        If IsActive = False Then
            MessageBox.Show("Please Open Solid Edge", "Message")
        Else
            Dim IsValid = CheckDocument()

            IsValid = True
            If (IsValid = True) Then

                'Dim dt = GetPartList()
                Dim dtUsers As DataTable = ReadAuthorExcelData(dtUsers)
                UpdateSummaryInfoAuthor(dtUsers)
                'Dim i = 0
                'For i = 0 To dt.Rows.Count
                '    activeDocPath = dt.Rows(i)(0).ToString()
                '    OpenDocument(i, dt)
                '    UpdateSummaryInfoAuthor(dtUsers)
                '    'CloseDocument(i, activeDocPath)
                'Next


            Else
                MessageBox.Show("Please Open Any Document in Solid Edge", "Message")
            End If

        End If
    End Sub

    Public Function CheckSEInstance() As Boolean

        OleMessageFilter.Register()
        Try
            objApp = Marshal.GetActiveObject("SolidEdge.Application")
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    Public Function CheckDocument() As Boolean
        Dim FilePath As String = objApp.ActiveDocument.FullName
        If (FilePath.EndsWith(".asm")) Then
            MainAssemblyPath = FilePath
            Return True
        ElseIf (FilePath.EndsWith(".psm")) Then
            Return True
        ElseIf (FilePath.EndsWith(".par")) Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function GetPartList() As DataTable


#Region "Not Working"
        'Try

        '    Objdocuemnts = objApp.Documents
        '    ObjdraftDoc = Objdocuemnts.Add("SolidEdge.DraftDocument")
        '    objSheet = ObjdraftDoc.ActiveSheet
        '    objModelLinks = ObjdraftDoc.ModelLinks
        '    Dim FileName As String = Nothing
        '    Dim File As String = DocumentPath1
        '    FileName = File
        '    objModelLink = objModelLinks.Add(FileName)
        '    ObjdrwViews = objSheet.DrawingViews
        '    ObjdrwView = ObjdrwViews.AddAssemblyView(From:=objModelLink, Orientation:=SolidEdgeDraft.ViewOrientationConstants.igDimetricTopBackLeftView, Scale:=1.0, x:=0.4, y:=0.4, ViewType:=SolidEdgeDraft.AssemblyDrawingViewTypeConstants.seAssemblyDesignedView)
        '    ObjdrwView.Caption = "BEC"
        '    ObjdrwView.DisplayCaption = False
        'Catch ex As Exception
        '    MessageBox.Show("Error While Getting Part Lists", ex.Message)
        'End Try

        'ObjdraftDoc = objApp.ActiveDocument
        'ObjpartLists = ObjdraftDoc.PartsLists
        'ObjpartList = ObjpartLists.Add(ObjdrwView, "BEC", 1, 1)
        'ObjpartList = ObjpartLists.Item(1)
        'ObjCols = ObjpartList.Columns
        'ObjRows = ObjpartList.Rows

        'For Each tableColumn In ObjCols.OfType(Of SolidEdgeDraft.TableColumn)()
        '    Dim dtcolums As DataColumn = New DataColumn(tableColumn.Header)
        '    dt.Columns.Add(dtcolums.ColumnName)
        'Next

        'Dim dtrows As DataRow = dt.NewRow()
        'dt.Rows.Add(dtrows)

        'For Each tableRow In ObjRows.OfType(Of SolidEdgeDraft.TableRow)()

        '    For Each tableColumn In ObjCols.OfType(Of SolidEdgeDraft.TableColumn)()

        '        If tableColumn.Show Then
        '            ObjTableCell = ObjpartList.Cell(tableRow.Index, tableColumn.Index)
        '            Dim rowindex As Integer = tableRow.Index - 1
        '            Dim colindex As Integer = tableColumn.Index - 1
        '            Dim tablevalue As String = ObjTableCell.value
        '            dt.Rows(rowindex)(colindex) = tablevalue
        '        End If
        '    Next

        '    dtrows = dt.NewRow()
        '    dt.Rows.Add(dtrows)
        'Next

        'Dim rowCnt As Integer = dt.Rows.Count

        'If dt.Rows.Count > 0 Then
        '    dt.Rows.RemoveAt(rowCnt - 1)
        'End If
#End Region
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

        objApp.Documents.CloseDocument(objDoc.FullName, False, "", False, False)


        Dim i = 0
        Dim count = dt.Columns.Count - 1
        For i = 0 To count
            If i <= count Then
            Else
                i = 0
            End If
            If (Not dt.Columns(i).ColumnName = "File Name (full path)") Then
                dt.Columns.RemoveAt(i)
                i = 0
                count = dt.Columns.Count - 1
                If (count = 0) Then
                    Exit For
                End If
            End If
        Next




        Return dt
    End Function
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
            MessageBox.Show("2 Author Updation.vb -- Employee Excel Path Missing or not Exist in Config")
            Exit Function
        End If

        Return dt
    End Function
    Public Sub OpenDocument(ByRef i As Integer, ByRef dt As DataTable)
        Try
            OleMessageFilter.Register()
            activeDocPath = dt.Rows(i)(0).ToString()
            objApp = CType(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
            Dim documents = objApp.Documents
            OpenSolidEdgeDocument(documents, activeDocPath, i, dt)
        Catch ex As System.Exception
            Console.WriteLine(ex)
        Finally
            OleMessageFilter.Unregister()
        End Try
    End Sub

    Public Sub OpenSolidEdgeDocument(ByVal documents As SolidEdgeFramework.Documents, ByVal path As String, ByRef i As Integer, ByRef dt As DataTable)
        Dim document As SolidEdgeFramework.SolidEdgeDocument = Nothing
        System.Threading.Thread.Sleep(3000)
        objApp.DisplayAlerts = False
        document = CType(documents.Open(path), SolidEdgeFramework.SolidEdgeDocument)

        Try

            For Each column As DataColumn In dt.Columns

                Select Case column.ColumnName
                    Case "File Name (full path)"

                        If dt.Rows(i)(column.ColumnName).ToString().EndsWith(".par") Then
                            PartdocPath = dt.Rows(i)(column.ColumnName).ToString()

                            Try
                                ObjPartdoc = CType(document, SolidEdgePart.PartDocument)
                                activeDocPath = PartdocPath
                            Catch __unusedException1__ As Exception
                                Dim name As String = dt.Rows(i)(column.ColumnName).ToString()
                                SheetMetalDocPath = name
                                PartdocPath = SheetMetalDocPath
                                System.Threading.Thread.Sleep(3000)
                                ObjSheetMetalDoc = CType(document, SolidEdgePart.SheetMetalDocument)
                                activeDocPath = SheetMetalDocPath
                            End Try
                        ElseIf dt.Rows(i)(column.ColumnName).ToString().EndsWith(".psm") Then
                            SheetMetalDocPath = dt.Rows(i)(column.ColumnName).ToString()
                            ObjSheetMetalDoc = CType(document, SolidEdgePart.SheetMetalDocument)
                            activeDocPath = SheetMetalDocPath
                        ElseIf dt.Rows(i)(column.ColumnName).ToString().EndsWith(".asm") Then
                            AssemblyDocPath = dt.Rows(i)(column.ColumnName).ToString()
                            ObjAssemblyDoc = CType(document, SolidEdgeAssembly.AssemblyDocument)
                            activeDocPath = AssemblyDocPath
                        End If
                End Select
            Next

        Catch ex As Exception
            MessageBox.Show("Open Part" & ex.Message + ex.StackTrace)
        End Try
    End Sub
    Public Sub UpdateSummaryInfoAuthor(ByRef dtUsers As DataTable)
        'Dim objDocument As SolidEdgeFramework.SolidEdgeDocument = objApp.ActiveDocument
        'Dim objParts As SolidEdgeAssembly.Occurrences

        'Dim propSets As SolidEdgeFramework.PropertySets = objApp.ActiveDocument.Properties

        'Dim custProps As SolidEdgeFramework.Properties = propSets.Item("SummaryInformation")

        ' Getting the parts objects of the AssemblyDocument object.
        'objParts = objDocument.Occurrences
        'Dim a As Integer

        'For a = 0 To dtUsers.Rows.Count - 1

        Try
            Dim objApp As SolidEdgeFramework.Application = System.Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application")
            Dim objPropSets As SolidEdgeFramework.PropertySets = Nothing
            Dim objProp As SolidEdgeFramework.Property = Nothing
            Dim objProps As SolidEdgeFramework.Properties = Nothing
            Dim objDocument As SolidEdgeFramework.SolidEdgeDocument = Nothing
            objDocument = objApp.ActiveDocument
            objPropSets = objDocument.Properties

            Dim finished As Integer = 0
            For Each objProps In objPropSets
                If objProps.Name = "SummaryInformation" Then
                    For Each objProp In objProps
                        If objProp.Name = "Author" Then
                            Dim dv As New DataView(dtUsers)
                            Dim WindowsAuthorName As String = "Windows Usernames"
                            Dim value1 As String = objProp.Value.ToString()
                            Dim filter As String = $"Convert([{WindowsAuthorName}], 'System.String') = '{value1}'"
                            dv.RowFilter = filter '$"{materialUsedCol}='{cmbMaterialUsed2_Mw.Text}'"
                            For Each drv As DataRowView In dv
                                'mtcMtrModelObj.authorList.Remove(OldValue)
                                objProp.Value = If(drv("Full Name") = Nothing, "", drv("Full Name"))
                                objProps.Save()
                                'objDocument.Save()
                                'MessageBox.Show("AUTHOR NAME UPDATED : " + objProp.Value, "Message")
                                Exit Sub

                            Next

                        End If
                    Next
                End If
            Next

            'objDocument.Close()

            'Next
        Catch ex As Exception

        End Try

    End Sub
    Private Function GetPropValue(ByVal prop1 As [Property])
        Dim value As String = String.Empty
        If prop1.Value IsNot Nothing Then
            value = prop1.Value.ToString().Trim()
        End If
        Return value
    End Function
    Public Sub CloseDocument(ByRef i As Integer, ByRef activeDocPath As String)
        Dim fInfo As FileInfo = New FileInfo(activeDocPath)
        Dim [readOnly] As Boolean = fInfo.IsReadOnly

        If activeDocPath.ToString().EndsWith(".par") Then

            If [readOnly] = False Then
                objApp.DisplayAlerts = False
                ObjPartdoc.Save()
            End If

            ObjPartdoc.Close()
        ElseIf activeDocPath.ToString().EndsWith(".psm") Then

            If [readOnly] = False Then
                objApp.DisplayAlerts = False
                ObjSheetMetalDoc.Save()
            End If

            ObjSheetMetalDoc.Close()
        ElseIf activeDocPath.ToString().EndsWith(".asm") Then

            If [readOnly] = False Then
                objApp.DisplayAlerts = False
                ObjAssemblyDoc.Save()
            End If
            If (Not activeDocPath = AssemblyDocPath) Then
                ObjAssemblyDoc.Close()
            End If

        End If
    End Sub
End Class