
Public Class MtcReviewBL
    Implements IMtcReviewBL

    Dim objApplication As SolidEdgeFramework.Application = Nothing
    Dim objAssemblyDocument As SolidEdgeAssembly.AssemblyDocument = Nothing

    Sub New(ByVal objApplication As SolidEdgeFramework.Application, ByVal objAssemblyDocument As SolidEdgeAssembly.AssemblyDocument)

        Me.objApplication = objApplication
        Me.objAssemblyDocument = objAssemblyDocument

    End Sub

    Public Function GetAssemblyPartList() As DataTable

        Dim dtAssemblyPartList As DataTable = New DataTable()

        Dim objDrawingView As SolidEdgeDraft.DrawingView = GetDrawingView()

        Dim draftDoc As SolidEdgeDraft.DraftDocument = objApplication.ActiveDocument

        Dim objPartsList As SolidEdgeDraft.PartsList = GetPartList(draftDoc, objDrawingView)

        dtAssemblyPartList = GetPartListDT(objPartsList)

        Return dtAssemblyPartList

    End Function

    Public Function ReadMTCExcel(mtcExcelPath As String) As DataSet Implements IMtcReviewBL.ReadMTCExcel
        Throw New NotImplementedException()
    End Function

    Public Function ReadPropSeedFile(propseedFilePath As String) As PropSeedFile Implements IMtcReviewBL.ReadPropSeedFile
        Throw New NotImplementedException()
    End Function

    Private Function GetDrawingView() As SolidEdgeDraft.DrawingView
        Dim objDrawingView As SolidEdgeDraft.DrawingView = Nothing


        Dim objDocuments As SolidEdgeFramework.Documents = objApplication.Documents
        Dim objDraft As SolidEdgeDraft.DraftDocument = objDocuments.Add("SolidEdge.DraftDocument")
        Dim objSheet As SolidEdgeDraft.Sheet = objDraft.ActiveSheet
        Dim objModelLinks As SolidEdgeDraft.ModelLinks = objDraft.ModelLinks

        Dim assemblyPath As String = objAssemblyDocument.FullName
        Dim filename As String = System.IO.Path.Combine(SolidEdgeCommunity.SolidEdgeUtils.GetTrainingFolderPath(), assemblyPath)

        Dim objModelLink As SolidEdgeDraft.ModelLink = objModelLinks.Add(filename)
        Dim objDrawingViews As SolidEdgeDraft.DrawingViews = objSheet.DrawingViews

        objDrawingView = objDrawingViews.AddAssemblyView([From]:=objModelLink, Orientation:=SolidEdgeDraft.ViewOrientationConstants.igDimetricTopBackLeftView, Scale:=1.0, x:=0.4, y:=0.4, ViewType:=SolidEdgeDraft.AssemblyDrawingViewTypeConstants.seAssemblyDesignedView)
        objDrawingView.Caption = "My New Drawing View"
        objDrawingView.DisplayCaption = False


        Return objDrawingView
    End Function

    Private Function GetPartList(ByVal objDoc As SolidEdgeDraft.DraftDocument, ByVal objDrawingView As SolidEdgeDraft.DrawingView) As SolidEdgeDraft.PartsList

        Dim objPartsLists As SolidEdgeDraft.PartsLists = objDoc.PartsLists

        ' Dim objPartsList As SolidEdgeDraft.PartsList = objPartsLists.Add(objDrawingView, "BEC", 1, 1)

        Dim objPartsList As SolidEdgeDraft.PartsList = objPartsLists.Item(1)

        Return objPartsList

    End Function

    Private Function GetPartListDT(ByVal objPartsList As SolidEdgeDraft.PartsList) As DataTable

        Dim tableCell As SolidEdgeDraft.TableCell = Nothing
        Dim dtPartList As System.Data.DataTable = New Data.DataTable()
        Dim myDataColumn1 As DataColumn = New DataColumn()
        myDataColumn1 = New DataColumn()
        myDataColumn1.ColumnName = "Select"
        myDataColumn1.DefaultValue = "0"
        myDataColumn1.DataType = System.Type.GetType("System.Boolean")
        dtPartList.Columns.Add(myDataColumn1)


        Dim cols As SolidEdgeDraft.TableColumns = objPartsList.Columns
        Dim rows As SolidEdgeDraft.TableRows = objPartsList.Rows
        For Each tableColumn In cols.OfType(Of SolidEdgeDraft.TableColumn)()
            Dim dtcolums As DataColumn = New DataColumn()
            dtcolums.ColumnName = tableColumn.Header
            dtPartList.Columns.Add(dtcolums)
            Debug.Print(tableColumn.Header)

        Next tableColumn

        Dim dtrows As DataRow = dtPartList.NewRow()
        dtPartList.Rows.Add(dtrows)

        Return dtPartList

    End Function

    Private Function IMtcReviewBL_GetAssemblyPartList() As DataTable Implements IMtcReviewBL.GetAssemblyPartList
        Throw New NotImplementedException()
    End Function

End Class
