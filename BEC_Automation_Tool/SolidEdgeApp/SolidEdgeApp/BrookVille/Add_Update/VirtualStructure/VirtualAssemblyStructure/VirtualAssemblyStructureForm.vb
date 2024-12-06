
Public Class VirtualAssemblyStructureForm

    Private Sub btnBrowseExcel_Click(sender As Object, e As EventArgs) Handles btnBrowseExcel.Click
        Try
            Using dialog As New OpenFileDialog
                dialog.Filter = "Excel files (*.xlsx)|*.xlsx"
                If dialog.ShowDialog() <> DialogResult.OK Then Return
                txtExcelPath.Text = dialog.FileName
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error in browse excel file {vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}", "Browse Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try
    End Sub

    Public Sub GetAlloccur(ByVal assemblyDoc As SolidEdgeAssembly.AssemblyDocument)

        Dim docs As SolidEdgeFramework.Documents = application.Documents
        Dim asmpath As String = assemblyDoc.FullName
        assemblyDoc = DirectCast(docs.Open(asmpath), SolidEdgeAssembly.AssemblyDocument)
        Dim occs As SolidEdgeAssembly.Occurrences = assemblyDoc.Occurrences

        For Each occ As SolidEdgeAssembly.Occurrence In occs
            'MsgBox(occ.Name)

            Dim obj As SolidEdgeFramework.ObjectType = occ.Type
            Dim occurname As String = IO.Path.GetFileNameWithoutExtension(occ.OccurrenceFileName)

            If occurname = userasmname Then
                'occ.Select(False)
                '  Dim asm2 As SolidEdgeAssembly.AssemblyDocument = DirectCast(occ.Parent, SolidEdgeAssembly.AssemblyDocument)
                ' Dim docs As SolidEdgeFramework.Documents = application.Documents
                ' Dim asm As SolidEdgeAssembly.AssemblyDocument = DirectCast(docs.Open(asmpath), SolidEdgeAssembly.AssemblyDocument)

                Tofalseoccuranceproperties(occ)
                assemblyDoc.UpdateDocument()
            End If

            If obj = SolidEdgeFramework.ObjectType.igSubAssembly Then
                Try
                    Dim assembldocument As SolidEdgeAssembly.AssemblyDocument = occ.OccurrenceDocument
                    GetAlloccur(assembldocument)
                Catch ex As Exception
                End Try
            End If

        Next
        assemblyDoc.Close(True)

    End Sub

    Dim userasmname As String = Nothing

    Private Sub btnCreateVirtaulAssembly_Click(sender As Object, e As EventArgs) Handles btnCreateVirtaulAssembly.Click

        If Not IO.File.Exists(txtExcelPath.Text) Then
            MessageBox.Show($"Please select Hedge excel", "Select Excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        If Not IO.Directory.Exists(txtDirectoryPath.Text) Then
            MessageBox.Show($"Please select destination folder for assembly creation", "Select Folder", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        If Not IO.File.Exists(txtfilepath.Text) Then
            'MessageBox.Show($"Please select  assembly  ", "Select Assembly", MessageBoxButtons.OK, MessageBoxIcon.Information)
            'Exit Sub
        End If
        If Not txtfilepath.Text.EndsWith(".asm") Then
            'MessageBox.Show($"Please select  assembly ", "Select Assembly", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        userasmname = IO.Path.GetFileNameWithoutExtension(txtfilepath.Text)

        Dim dicMainAssemblyDetails As Dictionary(Of String, Dictionary(Of String, List(Of String))) =
            ExcelUtilVirtualAssemblyCreation.ReadVirtualAssemblyCreationExcel(txtExcelPath.Text)

        waitStartSave()

        CreateAssemblyStructure(dicMainAssemblyDetails)

        For Each kvp As KeyValuePair(Of String, Dictionary(Of String, List(Of String))) In dicMainAssemblyDetails

            Dim asmname As String = kvp.Key
            Dim asmpath As String = txtDirectoryPath.Text + "\" + asmname + ".asm"
            Dim docs As SolidEdgeFramework.Documents = application.Documents
            Dim asm As SolidEdgeAssembly.AssemblyDocument = DirectCast(docs.Open(asmpath), SolidEdgeAssembly.AssemblyDocument)
            Dim occurances As SolidEdgeAssembly.Occurrences = asm.Occurrences

            For Each occur As SolidEdgeAssembly.Occurrence In occurances

                Dim occurname As String = IO.Path.GetFileNameWithoutExtension(occur.OccurrenceFileName)

                If occurname = userasmname Then

                    Tofalseoccuranceproperties(occur)

                End If

                If occur.Type = SolidEdgeFramework.ObjectType.igSubAssembly Then
                    Try
                        Dim assembldocument As SolidEdgeAssembly.AssemblyDocument = occur.OccurrenceDocument
                        GetAlloccur(assembldocument)
                    Catch ex As Exception
                    End Try
                End If
            Next

            asm.Save()
            asm.Close(True)
        Next

        SolidEdgeCommunity.OleMessageFilter.Unregister()
        application.Quit()

        If IO.Directory.Exists(txtDirectoryPath.Text) Then
            Process.Start(txtDirectoryPath.Text)
        End If

        WaitEndSave()

    End Sub

    Private Sub CreateAssemblyStructure(ByVal dicMainAssemblyDetails As Dictionary(Of String, Dictionary(Of String, List(Of String))))

        'Dim application As SolidEdgeFramework.Application = Nothing
        'SolidEdgeCommunity.OleMessageFilter.Register()
        'application = SolidEdgeCommunity.SolidEdgeUtils.Connect(True, False)

        Dim userAssemblyAdded As Boolean = False
        Dim userpath As String = txtfilepath.Text

        'Dim docs As SolidEdgeFramework.Documents = application.Documents
        'Dim userasm As SolidEdgeAssembly.AssemblyDocument = DirectCast(docs.Open(userpath), SolidEdgeAssembly.AssemblyDocument)
        'Dim occur As Object = userasm

        Try

            For Each kvp As KeyValuePair(Of String, Dictionary(Of String, List(Of String))) In dicMainAssemblyDetails

                Dim assemlbyName As String = kvp.Key
                assemlbyName = assemlbyName.Replace("/", "")

                Dim documents As SolidEdgeFramework.Documents = Nothing
                Dim assemblyDocument As SolidEdgeAssembly.AssemblyDocument = Nothing

                documents = application.Documents

                'Add Main assembly here.
                '=============
                assemblyDocument = DirectCast(documents.Add("SolidEdge.AssemblyDocument"), SolidEdgeAssembly.AssemblyDocument)
                assemblyDocument.Name = assemlbyName
                '=============

                Dim dicSubAssembly As Dictionary(Of String, List(Of String)) = kvp.Value
                Dim suboccurances As SolidEdgeAssembly.Occurrences

                For Each kvp2 As KeyValuePair(Of String, List(Of String)) In dicSubAssembly

                    Dim subAssemblyName As String = kvp2.Key
                    subAssemblyName = subAssemblyName.Replace("/", "")
                    If subAssemblyName.Contains("#") Then
                        Continue For
                    End If

                    Dim childAssemblyList As List(Of String) = kvp2.Value
                    Dim subassemblydocument As SolidEdgeAssembly.AssemblyDocument = Nothing
                    Dim subassemblypath As String = txtDirectoryPath.Text + "\" + subAssemblyName + ".asm"

                    If Not IO.File.Exists(subassemblypath) Then

                        subassemblydocument = DirectCast(documents.Add("SolidEdge.AssemblyDocument"), SolidEdgeAssembly.AssemblyDocument)

                        suboccurances = assemblyDocument.Occurrences
                    Else
                        'subassemblydocument = application.Documents.Open(subassemblypath)

                        subassemblydocument = DirectCast(documents.Open(subassemblypath), SolidEdgeAssembly.AssemblyDocument)

                        suboccurances = assemblyDocument.Occurrences

                        Dim subNewlyAddedoccurance As SolidEdgeAssembly.Occurrence = suboccurances.AddByFilename(subassemblypath)

                        subassemblydocument.Close(True)

                        Continue For

                    End If

                    For Each childassembly As String In childAssemblyList

                        childassembly = childassembly.Replace("/", "")
                        If childassembly.Contains("#") Then
                            Continue For
                        End If

                        Dim childpath As String = txtDirectoryPath.Text + "\" + childassembly + ".asm"

                        'Dim blankWorkingAssemblyPath As String = $"{txtDirectoryPath.Text}\{subAssemblyName}_Working.asm"

                        If Not IO.File.Exists(childpath) Then

                            Dim childassemblydocument As SolidEdgeAssembly.AssemblyDocument = DirectCast(documents.Add("SolidEdge.AssemblyDocument"), SolidEdgeAssembly.AssemblyDocument)
                            childassemblydocument.Name = childassembly
                            Dim childocurrences As SolidEdgeAssembly.Occurrences = childassemblydocument.Occurrences

                            If IO.File.Exists(userpath) Then
                                Dim childocurrence As SolidEdgeAssembly.Occurrence = childocurrences.AddByFilename(userpath)
                            End If

                            childassemblydocument.SaveAs(childpath)
                            childassemblydocument.Save()

                            ' Dim a As Integer = childocurrence.OccurrenceID

                            ' Dim occur As SolidEdgeAssembly.Occurrence = childocurrences.GetOccurrence(a)

                            ' Dim occur As SolidEdgeAssembly.Occurrence = childocurrences.GetOccurrence(a)

                            ' Tofalseoccuranceproperties(childocurrence)

                            childassemblydocument.Close(True)

                        End If

                        Dim subassemblyoccurance As SolidEdgeAssembly.Occurrences = subassemblydocument.Occurrences
                        Dim occuranceNewlyAdded As SolidEdgeAssembly.Occurrence = subassemblyoccurance.AddByFilename(childpath)

                        If userAssemblyAdded = False Then

                            If IO.File.Exists(userpath) Then
                                Dim asmaoccur As SolidEdgeAssembly.Occurrence = subassemblyoccurance.AddByFilename(userpath)

                                'Dim blnakWorkingAssemblyOccur As SolidEdgeAssembly.Occurrence = subassemblyoccurance.AddByFilename(blankWorkingAssemblyPath)

                                userAssemblyAdded = True
                            End If

                        End If

                    Next

                    userAssemblyAdded = False
                    If Not IO.File.Exists(subassemblypath) Then
                        subassemblydocument.SaveAs(subassemblypath)
                    End If
                    subassemblydocument.Close(True)
                    Dim mainoccurance As SolidEdgeAssembly.Occurrences = assemblyDocument.Occurrences
                    Dim mainoccuranceNewlyAdded As SolidEdgeAssembly.Occurrence = mainoccurance.AddByFilename(subassemblypath)

                    'Dim useroccuranceNewlyAdded As SolidEdgeAssembly.Occurrence = mainoccurance.AddByFilename(userpath)
                    'userAssemblyAdded = True

                Next

                Dim assemblypath As String = txtDirectoryPath.Text + "\" + assemlbyName + ".asm"
                Dim userasmoccurances As SolidEdgeAssembly.Occurrences = assemblyDocument.Occurrences

                If IO.File.Exists(userpath) Then

                    Dim userasmoccuranceNewlyAdded As SolidEdgeAssembly.Occurrence = userasmoccurances.AddByFilename(userpath)

                End If

                assemblyDocument.SaveAs(assemblypath)
                assemblyDocument.Close(True)

            Next

            ' MessageBox.Show("Completed")
        Catch ex As Exception
            MessageBox.Show($"Error in create assembly structure {ex.Message} {vbNewLine} {ex.StackTrace}")
        End Try
        ' SolidEdgeCommunity.OleMessageFilter.Unregister()
        ' application.Quit()
    End Sub

    Public Sub Tofalseoccuranceproperties(ByVal occur As SolidEdgeAssembly.Occurrence)

        occur.IncludeInBom = False
        occur.DisplayInDrawings = False
        occur.DisplayInSubAssembly = False
        occur.IncludeInInterference = False
        occur.IncludeInPhysicalProperties = False

    End Sub

    Private Sub btnDirectoryPath_Click(sender As Object, e As EventArgs) Handles btnDirectoryPath.Click
        Dim MyFolderBrowser As New System.Windows.Forms.FolderBrowserDialog
        Dim dlgResult As DialogResult = MyFolderBrowser.ShowDialog()

        If dlgResult = Windows.Forms.DialogResult.OK Then
            txtDirectoryPath.Text = MyFolderBrowser.SelectedPath
        End If
    End Sub

    Private Sub CreateAssembly(ByVal directoryPath As String, ByVal assemblyName As String)

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

    Dim application As SolidEdgeFramework.Application = Nothing

    Private Sub VirtualAssemblyStructureForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = $"{Me.Text} ({GlobalEntity.Version}) "
        SolidEdgeCommunity.OleMessageFilter.Register()
        application = SolidEdgeCommunity.SolidEdgeUtils.Connect(True, False)
        application.DisplayAlerts = False

        txtDirectoryPath.Text = Config.configObj.virtualAssemblyOutputDirec

    End Sub

    Private Sub btnopenfile_Click(sender As Object, e As EventArgs) Handles btnopenfile.Click
        Dim MyfileBrowser As New System.Windows.Forms.OpenFileDialog
        Dim dlgResult As DialogResult = MyfileBrowser.ShowDialog()

        If dlgResult = Windows.Forms.DialogResult.OK Then
            txtfilepath.Text = MyfileBrowser.FileName
        End If
    End Sub

End Class