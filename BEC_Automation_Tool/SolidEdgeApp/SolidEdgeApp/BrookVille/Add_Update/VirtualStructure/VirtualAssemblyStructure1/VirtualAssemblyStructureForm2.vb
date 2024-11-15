
Imports System.Reflection
'Imports CADTeam.SolidEdge.Framework.Interop
Imports SolidEdgeFileProperties
Imports WK.Libraries.BetterFolderBrowserNS

Public Class VirtualAssemblyStructureForm2

    Dim application As SolidEdgeFramework.Application = Nothing
    Dim skip As Boolean = True

    Private Sub BtnBrowseExcel_Click(sender As Object, e As EventArgs) Handles btnBrowseExcel.Click
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

    Private Function AddAssembly(ByVal assemblyName As String) As SolidEdgeAssembly.AssemblyDocument
        Dim documents As SolidEdgeFramework.Documents = application.Documents
        Dim assemblyDocument As SolidEdgeAssembly.AssemblyDocument = DirectCast(documents.Add("SolidEdge.AssemblyDocument"), SolidEdgeAssembly.AssemblyDocument)
        assemblyDocument.Name = assemblyName
        Return assemblyDocument
    End Function

    ''' <summary>
    ''' Add assembly document into the application document collection and save the assembly document
    ''' </summary>
    ''' <param name="assemblyName">Assembly Name</param>
    ''' <param name="outputDirPath">Output directory location</param>
    ''' <returns></returns>
    Private Function AddAssembly2(ByVal assemblyName As String, ByVal outputDirPath As String) As SolidEdgeAssembly.AssemblyDocument

        assemblyName = assemblyName.Replace("/", "")
        Dim documents As SolidEdgeFramework.Documents = application.Documents
        Dim assemblyDocument As SolidEdgeAssembly.AssemblyDocument = Nothing
        If IO.File.Exists(IO.Path.Combine(outputDirPath, $"{assemblyName}.asm")) Then
            assemblyDocument = DirectCast(documents.Open(IO.Path.Combine(outputDirPath, $"{assemblyName}.asm")), SolidEdgeAssembly.AssemblyDocument)
        Else
            assemblyDocument = DirectCast(documents.Add("SolidEdge.AssemblyDocument"), SolidEdgeAssembly.AssemblyDocument)
            assemblyDocument.SaveAs(IO.Path.Combine(outputDirPath, $"{assemblyName}.asm"))
        End If

        Return assemblyDocument


    End Function

    ''' <summary>
    ''' Add the child document as occurence into the assembly document
    ''' </summary>
    ''' <param name="parentAssemblyDoc">Parent assembly document</param>
    ''' <param name="childDocPath">Occurence to add into the parent assembly</param>
    ''' <returns></returns>
    Private Function AddAssemblyOccurences(ByRef parentAssemblyDoc As SolidEdgeAssembly.AssemblyDocument, ByVal childDocPath As String, ByRef dicAssemblyChild As Dictionary(Of String, List(Of String))) As SolidEdgeAssembly.Occurrence

        Dim parentAssemblyOccurences As SolidEdgeAssembly.Occurrences = parentAssemblyDoc.Occurrences

        Dim newlyAddedOcc As SolidEdgeAssembly.Occurrence = parentAssemblyOccurences.AddByFilename(childDocPath)

        Return newlyAddedOcc

    End Function
    Private Function AddAssemblyOccurences2(ByRef parentAssemblyDoc As SolidEdgeAssembly.AssemblyDocument, ByVal childDocPath As String, ByRef dicAssemblyChild As Dictionary(Of String, List(Of String))) As SolidEdgeAssembly.Occurrence
        Dim newlyAddedOcc As SolidEdgeAssembly.Occurrence = Nothing
        Try

            Dim parentAssemblyPath As String = parentAssemblyDoc.FullName

            Dim parentAssemblyOccurenceList As New List(Of String)()

            If dicAssemblyChild.ContainsKey(parentAssemblyPath) Then
                parentAssemblyOccurenceList = dicAssemblyChild(parentAssemblyPath)
            End If

            Dim parentAssemblyOccurences As SolidEdgeAssembly.Occurrences = parentAssemblyDoc.Occurrences



            If Not parentAssemblyOccurenceList.Contains(childDocPath) Then

                newlyAddedOcc = parentAssemblyOccurences.AddByFilename(childDocPath)
                parentAssemblyOccurenceList.Add(childDocPath)
            End If

            If dicAssemblyChild.ContainsKey(parentAssemblyPath) Then
                dicAssemblyChild(parentAssemblyPath) = parentAssemblyOccurenceList
            Else
                dicAssemblyChild.Add(parentAssemblyPath, parentAssemblyOccurenceList)
            End If


        Catch ex As Exception
            CustomLogUtil.Log($"In Assembly Occurance:", ex.Message, ex.StackTrace)
        End Try
        Return newlyAddedOcc
    End Function
    Dim dicAssemblyChild As New Dictionary(Of String, List(Of String))()
    Private Sub CreateVirtualAssemblyNew(ByVal dicMainAssemblyDetails As Dictionary(Of String, Dictionary(Of String, List(Of String))), ByVal topLevelAssemblyNames As String)

        dicAssemblyChild = New Dictionary(Of String, List(Of String))()
        Dim mainAssemblyName As String = topLevelAssemblyNames


        For Each kvp As KeyValuePair(Of String, Dictionary(Of String, List(Of String))) In dicMainAssemblyDetails

            Dim firstLevelAssemblyName As String = kvp.Key
            Dim dicTopLevelAssemblyDetails As Dictionary(Of String, List(Of String)) = kvp.Value
            CreateAssemblyDetails(firstLevelAssemblyName, dicTopLevelAssemblyDetails, mainAssemblyName)
            CustomLogUtil.Heading($"Creating First-Level Assembly :{firstLevelAssemblyName}")

        Next

    End Sub
    'temp11APR2023
    Private Sub CreateVirtualAssemblyNew2(ByVal dicMainAssemblyDetails As Dictionary(Of String, Dictionary(Of String, List(Of String))), ByVal topLevelAssemblyNames As String)

        dicAssemblyChild = New Dictionary(Of String, List(Of String))()
        Dim mainAssemblyName As String = topLevelAssemblyNames


        For Each kvp As KeyValuePair(Of String, Dictionary(Of String, List(Of String))) In dicMainAssemblyDetails

            Dim firstLevelAssemblyName As String = kvp.Key
            Dim dicTopLevelAssemblyDetails As Dictionary(Of String, List(Of String)) = kvp.Value
            CreateAssemblyDetailsNew(firstLevelAssemblyName, dicTopLevelAssemblyDetails, mainAssemblyName)
            CustomLogUtil.Heading($"Creating First-Level Assembly :{firstLevelAssemblyName}")

        Next

    End Sub

    Private Sub CreateAssemblyDetails(ByVal firstLevelAssemblyName As String, ByVal dicTopLevelAssemblyDetails As Dictionary(Of String, List(Of String)), ByVal mainAssemblyDoc As String)
        'Dim dicAssemblyChild As Dictionary(Of String, List(Of String)) = New Dictionary(Of String, List(Of String))()
        Try

            Dim firstLevelAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(firstLevelAssemblyName, txtDirectoryPath.Text)
            Dim firstLevelAssemblyDocPath As String = IO.Path.Combine(txtDirectoryPath.Text, $"{firstLevelAssemblyName}.asm")
            If chkAddUserAssembly.Checked Then
                CustomLogUtil.Log($"Adding first-Level AssemblyOccurences: {firstLevelAssemblyDoc}")
                'AddAssemblyOccurences2(firstLevelAssemblyDoc, txtfilepath.Text, dicAssemblyChild)
                AddAssemblyOccurences2(firstLevelAssemblyDoc, txtfilepath.Text, dicAssemblyChild)
            End If

            For Each kvp As KeyValuePair(Of String, List(Of String)) In dicTopLevelAssemblyDetails
                Dim secondLevelAssemblyName As String = kvp.Key
                CustomLogUtil.Log($"Adding Second-Level Assembly Document: {secondLevelAssemblyName}")
                Dim secondLevelAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(secondLevelAssemblyName, txtDirectoryPath.Text)
                Dim secondLevelAssemblyDocPath As String = IO.Path.Combine(txtDirectoryPath.Text, $"{secondLevelAssemblyName}.asm")

                If chkAddUserAssembly.Checked Then
                    CustomLogUtil.Log($"Adding AssemblyOccurences :{secondLevelAssemblyDoc}")
                    AddAssemblyOccurences2(secondLevelAssemblyDoc, txtfilepath.Text, dicAssemblyChild)
                End If

                Dim secondLevelAssemblyNameWorking As String = $"{kvp.Key}_Working"
                CustomLogUtil.Log($"Adding Second-Level Assembly Name Working :{secondLevelAssemblyNameWorking}")
                Dim secondLevelAssemblyDocWorking As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(secondLevelAssemblyNameWorking, txtDirectoryPath.Text)
                Dim secondLevelAssemblyDocPathWorking As String = IO.Path.Combine(txtDirectoryPath.Text, $"{secondLevelAssemblyNameWorking}.asm")

                For Each childAssembly As String In kvp.Value

                    Dim thirdLevelAssemblyName As String = childAssembly
                    CustomLogUtil.Log($"Adding Third-Level Assembly :{thirdLevelAssemblyName}")
                    Dim thirdLevelAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(thirdLevelAssemblyName, txtDirectoryPath.Text)
                    Dim thirdLevelAssemblyDocPath As String = IO.Path.Combine(txtDirectoryPath.Text, $"{thirdLevelAssemblyName}.asm")
                    CustomLogUtil.Log($"Adding AssemblyOccurences :{thirdLevelAssemblyName}")
                    AddAssemblyOccurences2(secondLevelAssemblyDocWorking, thirdLevelAssemblyDocPath, dicAssemblyChild)

                    thirdLevelAssemblyDoc.Save()
                    thirdLevelAssemblyDoc.Close(True)
                Next

                CustomLogUtil.Log($"Adding AssemblyOccurences :{secondLevelAssemblyDoc}")
                AddAssemblyOccurences2(secondLevelAssemblyDoc, secondLevelAssemblyDocPathWorking, dicAssemblyChild)

                CustomLogUtil.Log($"Save Second-Level Assembly DocWorking :{secondLevelAssemblyDoc}")
                secondLevelAssemblyDocWorking.Save()
                secondLevelAssemblyDocWorking.Close(True)

                CustomLogUtil.Log($"Adding First-Level AssemblyOccurences :{firstLevelAssemblyDoc}")
                AddAssemblyOccurences2(firstLevelAssemblyDoc, secondLevelAssemblyDocPath, dicAssemblyChild)



                CustomLogUtil.Log($"Save Second-Level Assembly :{secondLevelAssemblyDoc}")
                secondLevelAssemblyDoc.Save()
                secondLevelAssemblyDoc.Close(True)



            Next

            CustomLogUtil.Log($"Save First-Level Assembly :{firstLevelAssemblyDoc}")
            firstLevelAssemblyDoc.Save()
            firstLevelAssemblyDoc.Close(True)

        Catch ex As Exception
            If skip = True Then
                MessageBox.Show($"While Creating Assembly Details {vbNewLine}{ex.Message}{ ex.StackTrace}", "Message")
            End If
            CustomLogUtil.Log($"While Creating Assembly Details", ex.Message, ex.StackTrace)
        End Try
    End Sub
    'temp11APR2023
    Private Sub CreateAssemblyDetailsNew(ByVal firstLevelAssemblyName As String, ByVal dicTopLevelAssemblyDetails As Dictionary(Of String, List(Of String)), ByVal mainAssemblyDoc As String)
        'Dim dicAssemblyChild As Dictionary(Of String, List(Of String)) = New Dictionary(Of String, List(Of String))()

        Try

            Dim TopLevelAssemblyDetails As Array = mainAssemblyDoc.Split(New Char() {"/C"})
            Dim TopLevelTitleName As String = TopLevelAssemblyDetails(1)
            mainAssemblyDoc = TopLevelAssemblyDetails(0)
            Dim TopLevelAssemblyDocumentPath = IO.Path.Combine(txtDirectoryPath.Text, $"{mainAssemblyDoc}.asm")

            Dim TopLevellAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(mainAssemblyDoc, txtDirectoryPath.Text)
            AddAssemblyOccurences2(TopLevellAssemblyDoc, txtfilepath.Text, dicAssemblyChild)


            Dim firstLevelAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(firstLevelAssemblyName, txtDirectoryPath.Text)
            Dim firstLevelAssemblyDocPath As String = IO.Path.Combine(txtDirectoryPath.Text, $"{firstLevelAssemblyName}.asm")
            If chkAddUserAssembly.Checked Then
                CustomLogUtil.Log($"Adding first-Level AssemblyOccurences: {firstLevelAssemblyDoc}")
                'AddAssemblyOccurences2(firstLevelAssemblyDoc, txtfilepath.Text, dicAssemblyChild)
                AddAssemblyOccurences2(firstLevelAssemblyDoc, txtfilepath.Text, dicAssemblyChild)
            End If
            Dim FirstLevelTitleName As String = String.Empty
            For Each kvp As KeyValuePair(Of String, List(Of String)) In dicTopLevelAssemblyDetails


                Dim RefLevelAssemblyName As String = mainAssemblyDoc + "- Ref"
                CustomLogUtil.Log($"Adding Second-Level Assembly Document: {RefLevelAssemblyName}")
                Dim RefLevelAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(RefLevelAssemblyName, txtDirectoryPath.Text)
                Dim RefLevelAssemblyDocPath As String = IO.Path.Combine(txtDirectoryPath.Text, $"{RefLevelAssemblyName}.asm")

                If chkAddUserAssembly.Checked Then
                    CustomLogUtil.Log($"Adding AssemblyOccurences :{RefLevelAssemblyDoc}")
                    AddAssemblyOccurences2(RefLevelAssemblyDoc, txtfilepath.Text, dicAssemblyChild)
                End If

                Dim secondLevelAssemblyName As String = kvp.Key
                FirstLevelTitleName = secondLevelAssemblyName
                Dim secondLevelAssemblyDoc As SolidEdgeAssembly.AssemblyDocument
                Dim secondLevelAssemblyDocPath As String
                If (Not secondLevelAssemblyName.Contains("Main Assembly")) Then
                    CustomLogUtil.Log($"Adding Second-Level Assembly Name  :{secondLevelAssemblyName}")
                    secondLevelAssemblyDoc = AddAssembly2(secondLevelAssemblyName, txtDirectoryPath.Text)
                    secondLevelAssemblyDocPath = IO.Path.Combine(txtDirectoryPath.Text, $"{secondLevelAssemblyName}.asm")
                End If


                'for each childassembly as string in kvp.value

                '    dim thirdlevelassemblyname as string = childassembly
                '    customlogutil.log($"adding third-level assembly :{thirdlevelassemblyname}")
                '    dim thirdlevelassemblydoc as solidedgeassembly.assemblydocument = addassembly2(thirdlevelassemblyname, txtdirectorypath.text)
                '    dim thirdlevelassemblydocpath as string = io.path.combine(txtdirectorypath.text, $"{thirdlevelassemblyname}.asm")
                '    customlogutil.log($"adding assemblyoccurences :{thirdlevelassemblyname}")
                '    addassemblyoccurences2(secondlevelassemblydoc, thirdlevelassemblydocpath, dicassemblychild)

                '    thirdlevelassemblydoc.save()
                '    thirdlevelassemblydoc.close(true)
                'next

                If (Not secondLevelAssemblyName.Contains("Main Assembly")) Then
                    CustomLogUtil.Log($"Adding AssemblyOccurences :{RefLevelAssemblyDoc}")
                    AddAssemblyOccurences2(RefLevelAssemblyDoc, secondLevelAssemblyDocPath, dicAssemblyChild)

                    CustomLogUtil.Log($"Save Second-Level Assembly Doc :{secondLevelAssemblyDoc}")
                    secondLevelAssemblyDoc.Save()
                    secondLevelAssemblyDoc.Close(True)

                End If



                CustomLogUtil.Log($"Adding First-Level AssemblyOccurences :{firstLevelAssemblyDoc}")
                AddAssemblyOccurences2(firstLevelAssemblyDoc, RefLevelAssemblyDocPath, dicAssemblyChild)

                CustomLogUtil.Log($"Adding Top-Level AssemblyOccurences :{TopLevellAssemblyDoc}")
                AddAssemblyOccurences2(TopLevellAssemblyDoc, firstLevelAssemblyDocPath, dicAssemblyChild)

                CustomLogUtil.Log($"Save Second-Level Assembly :{RefLevelAssemblyDoc}")
                RefLevelAssemblyDoc.Save()
                RefLevelAssemblyDoc.Close(True)




                Dim SubAssmblyName As String = kvp.Key().ToString()
                If (Not SubAssmblyName.Contains(TopLevelTitleName)) Then
                    Dim SubAssmblyDoc As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(SubAssmblyName, txtDirectoryPath.Text)
                    Dim SubAssmblyDocPath As String = IO.Path.Combine(txtDirectoryPath.Text, $"{SubAssmblyName}.asm")
                    AddAssemblyOccurences2(firstLevelAssemblyDoc, SubAssmblyDocPath, dicAssemblyChild)
                    SubAssmblyDoc.Save()
                    For Each childassembly As String In kvp.Value

                        Dim thirdlevelassemblyname As String = childassembly
                        Dim Path = txtDirectoryPath.Text
                        'CustomLogUtil.Log($"adding third-level assembly :{thirdlevelassemblyname}")
                        'Dim thirdlevelassemblydoc As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(thirdlevelassemblyname, txtDirectoryPath.Text)
                        'Dim thirdlevelassemblydocpath As String = IO.Path.Combine(txtDirectoryPath.Text, $"{thirdlevelassemblyname}.asm")
                        'CustomLogUtil.Log($"adding assemblyoccurences :{thirdlevelassemblyname}")
                        'AddAssemblyOccurences2(secondLevelAssemblyDoc, thirdlevelassemblydocpath, dicAssemblyChild)

                        'thirdlevelassemblydoc.Save()
                        'thirdlevelassemblydoc.Close(True)
                        AddToGroup(Path, thirdlevelassemblyname)
                    Next
                    'SetSummaryInformationPropertyForTopAndFirstLevelAssembly(TopLevelTitleName)
                    SetSummaryInformationPropertyForTopAndFirstLevelAssembly(firstLevelAssemblyName, firstLevelAssemblyDocPath)
                    SubAssmblyDoc.Save()
                    SubAssmblyDoc.Close(True)
                End If






            Next


            CustomLogUtil.Log($"Save First-Level Assembly :{firstLevelAssemblyDoc}")
            firstLevelAssemblyDoc.Save()
            firstLevelAssemblyDoc.Close(True)

            'Add Title in Top Level Assembly
            SetSummaryInformationPropertyForTopAndFirstLevelAssembly(TopLevelTitleName, TopLevelAssemblyDocumentPath)

            'SetSummaryInformationPropertyForTopAndFirstLevelAssembly(TopLevelTitleName)
            CustomLogUtil.Log($"Save Top-Level Assembly :{TopLevellAssemblyDoc}")
            TopLevellAssemblyDoc.Save()
            TopLevellAssemblyDoc.Close(True)

        Catch ex As Exception
            If skip = True Then
                MessageBox.Show($"While Creating Assembly Details {vbNewLine}{ex.Message}{ ex.StackTrace}", "Message")
            End If
            CustomLogUtil.Log($"While Creating Assembly Details", ex.Message, ex.StackTrace)
        End Try
    End Sub
    Public Sub AddToGroup(ByVal Path As String, ByVal GroupName As String)
        Dim application As SolidEdgeFramework.Application = GetObject(, "SolidEdge.Application")
        Dim documents1 As SolidEdgeFramework.Documents = Nothing
        Dim objPartDoc As SolidEdgeAssembly.Occurrence
        Dim NumComponents = "1"
        Dim components(NumComponents) As Object



        documents1 = application.Documents()
        If (Not System.IO.File.Exists(IO.Path.Combine(Path, "Dummy.par"))) Then

            documents1.Add("SolidEdge.PartDocument", Missing.Value)
            application.DisplayAlerts = False
            application.ActiveDocument.SaveAs(IO.Path.Combine(Path, "Dummy.par"))
            documents1.CloseDocument(IO.Path.Combine(Path, "Dummy.par"))
        End If

        Dim document As SolidEdgeAssembly.AssemblyDocument = application.ActiveDocument
        objPartDoc = document.Occurrences.AddByFilename(IO.Path.Combine(Path, "Dummy.par"))

        components(0) = objPartDoc

        Dim seAssemblyGroups As SolidEdgeAssembly.AssemblyGroups = document.AssemblyGroups
        Dim group As SolidEdgeAssembly.AssemblyGroup = Nothing
        group = seAssemblyGroups.Add(1, components)
        'group.AddToGroup(NumComponents, components)
        Dim Gname = GroupName
        group.Name = Gname
        'Add more components as needed'
    End Sub
    'temp11APR2023
    Public Sub SetSummaryInformationPropertyForTopAndFirstLevelAssembly(ByRef LevelTitleName As String, ByRef firstLevelAssemblyDocPath As String)
        Dim application As SolidEdgeFramework.Application = GetObject(, "SolidEdge.Application")
        Dim objDocument As SolidEdgeAssembly.AssemblyDocument = application.ActiveDocument
        If Not objDocument Is Nothing Then
            Try

                LevelTitleName = LevelTitleName.Replace(".asm", "")

                Dim propSets As SolidEdgeFramework.PropertySets = objDocument.Properties

                'Dim SummaryProps As Properties = propSets.Item("SummaryInformation")



                ''For Each prop1 As [Property] In SummaryProps
                'For Each prop1 In SummaryProps

                '    Try


                '        If prop1.Name = "Title" Then
                '            prop1.Value = LevelTitleName
                '            Exit For
                '        End If

                '    Catch ex As Exception
                '    End Try

                'Next


                For Each objProps In propSets
                    If objProps.Name = "SummaryInformation" Then
                        For Each objProp In objProps
                            If objProp.Name = "Title" Then
                                objProp.Value = LevelTitleName
                                Exit For
                            End If
                        Next
                    End If
                Next


                propSets.Save()

            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub CreateVirtualAssembly(ByVal dicMainAssemblyDetails As Dictionary(Of String, Dictionary(Of String, List(Of String))), ByVal topLevelAssemblyNames As String)

        Dim AssemblyNames As String() = topLevelAssemblyNames.Split("/")

        'Dim pumpCar As String() = AssemblyNames(0).Split("-")
        Dim pumpCarAssemblyName As String = AssemblyNames(0)

        'Dim GenCar As String() = AssemblyNames(1).Split("-")
        Dim GenCarAssemblyName As String = AssemblyNames(1)

        CreateGenCar(dicMainAssemblyDetails, GenCarAssemblyName)

        CreatePumpCar(dicMainAssemblyDetails, pumpCarAssemblyName)



        Exit Sub

        Dim mainAssemblyName As String = "900-00028"

        Dim mainAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(mainAssemblyName, txtDirectoryPath.Text)

        Dim dicAssemblyChild As New Dictionary(Of String, List(Of String))()

        For Each kvp1 As KeyValuePair(Of String, Dictionary(Of String, List(Of String))) In dicMainAssemblyDetails

            Dim firstLevelAssemblyName As String = kvp1.Key
            If firstLevelAssemblyName.Contains("#") Then
                Continue For
            End If

            Dim firsLevelAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(firstLevelAssemblyName, txtDirectoryPath.Text)
            Dim firstLevelAssemblyDocPath As String = IO.Path.Combine(txtDirectoryPath.Text, $"{firstLevelAssemblyName}.asm")
            AddAssemblyOccurences(mainAssemblyDoc, firstLevelAssemblyDocPath, dicAssemblyChild)
            firsLevelAssemblyDoc.Save()

            For Each kvp2 As KeyValuePair(Of String, List(Of String)) In kvp1.Value


                Dim secondLevelAssemblyName As String = kvp2.Key

                If secondLevelAssemblyName.Contains("#") Then
                    Continue For
                End If

                Dim secondLevelAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(secondLevelAssemblyName, txtDirectoryPath.Text)
                Dim secondLevelAssemblyDocPath As String = IO.Path.Combine(txtDirectoryPath.Text, $"{secondLevelAssemblyName}.asm")
                AddAssemblyOccurences(firsLevelAssemblyDoc, secondLevelAssemblyDocPath, dicAssemblyChild)
                secondLevelAssemblyDoc.Save()

                For Each thirdLevelAssemblyName As String In kvp2.Value

                    If thirdLevelAssemblyName.Contains("#") Then
                        Continue For
                    End If

                    Dim thirdLevelAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(thirdLevelAssemblyName, txtDirectoryPath.Text)
                    Dim thirdtLevelAssemblyDocPath As String = IO.Path.Combine(txtDirectoryPath.Text, $"{thirdLevelAssemblyName}.asm")
                    Dim thirdLevelAssemblyDocPath As String = IO.Path.Combine(txtDirectoryPath.Text, $"{thirdLevelAssemblyName}.asm")
                    AddAssemblyOccurences2(secondLevelAssemblyDoc, thirdLevelAssemblyDocPath, dicAssemblyChild)

                    thirdLevelAssemblyDoc.Save()
                    thirdLevelAssemblyDoc.Close(True)
                Next

                secondLevelAssemblyDoc.Save()
                secondLevelAssemblyDoc.Close(True)

            Next
            firsLevelAssemblyDoc.Save()
            firsLevelAssemblyDoc.Close(True)
        Next

        mainAssemblyDoc.Save()
        mainAssemblyDoc.Close(True)
    End Sub

    Private Sub CreatePumpCar(ByVal dicMainAssemblyDetails As Dictionary(Of String, Dictionary(Of String, List(Of String))), ByVal mainAssemblyName As String)
        'Dim mainAssemblyName As String = "900-00028"
        Try
            Dim mainAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(mainAssemblyName, txtDirectoryPath.Text)

            Dim dicAssemblyChild As New Dictionary(Of String, List(Of String))()

            For Each kvp1 As KeyValuePair(Of String, Dictionary(Of String, List(Of String))) In dicMainAssemblyDetails


                If Not kvp1.Key.ToString().ToUpper().Trim() = "PUMP CAR" Then
                    Continue For
                End If

                Dim firstLevelAssemblyName As String = kvp1.Key
                If firstLevelAssemblyName.Contains("#") Then
                    Continue For
                End If

                Dim firsLevelAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(firstLevelAssemblyName, txtDirectoryPath.Text)
                Dim firstLevelAssemblyDocPath As String = IO.Path.Combine(txtDirectoryPath.Text, $"{firstLevelAssemblyName}.asm")
                AddAssemblyOccurences(mainAssemblyDoc, firstLevelAssemblyDocPath, dicAssemblyChild)

                'Add user assembly path
                If chkAddUserAssembly.Checked Then
                    AddAssemblyOccurences(mainAssemblyDoc, txtfilepath.Text, dicAssemblyChild)
                End If

                firsLevelAssemblyDoc.Save()

                For Each kvp2 As KeyValuePair(Of String, List(Of String)) In kvp1.Value


                    Dim secondLevelAssemblyName As String = kvp2.Key

                    If secondLevelAssemblyName.Contains("#") Then
                        Continue For
                    End If

                    Dim secondLevelAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(secondLevelAssemblyName, txtDirectoryPath.Text)
                    Dim secondLevelAssemblyDocPath As String = IO.Path.Combine(txtDirectoryPath.Text, $"{secondLevelAssemblyName}.asm")
                    AddAssemblyOccurences(firsLevelAssemblyDoc, secondLevelAssemblyDocPath, dicAssemblyChild)


                    'Add user assembly path
                    If chkAddUserAssembly.Checked Then
                        AddAssemblyOccurences(firsLevelAssemblyDoc, txtfilepath.Text, dicAssemblyChild)
                    End If

                    secondLevelAssemblyDoc.Save()

                    For Each thirdLevelAssemblyName As String In kvp2.Value

                        If thirdLevelAssemblyName.Contains("#") Then
                            Continue For
                        End If

                        Dim thirdLevelAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(thirdLevelAssemblyName, txtDirectoryPath.Text)
                        Dim thirdtLevelAssemblyDocPath As String = IO.Path.Combine(txtDirectoryPath.Text, $"{thirdLevelAssemblyName}.asm")
                        Dim thirdLevelAssemblyDocPath As String = IO.Path.Combine(txtDirectoryPath.Text, $"{thirdLevelAssemblyName}.asm")

                        AddAssemblyOccurences2(secondLevelAssemblyDoc, thirdLevelAssemblyDocPath, dicAssemblyChild)


                        'Add user assembly path
                        If chkAddUserAssembly.Checked Then
                            AddAssemblyOccurences(secondLevelAssemblyDoc, txtfilepath.Text, dicAssemblyChild)
                        End If

                        thirdLevelAssemblyDoc.Save()
                        thirdLevelAssemblyDoc.Close(True)
                    Next

                    secondLevelAssemblyDoc.Save()
                    secondLevelAssemblyDoc.Close(True)

                Next
                firsLevelAssemblyDoc.Save()
                firsLevelAssemblyDoc.Close(True)
            Next

            mainAssemblyDoc.Save()
            mainAssemblyDoc.Close(True)


        Catch ex As Exception
            MessageBox.Show($"Error While Create Pump Car : {vbNewLine}{ex.Message}{ ex.StackTrace}", "Message")
            CustomLogUtil.Log("While Create Pump Car ", ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub CreateGenCar(ByVal dicMainAssemblyDetails As Dictionary(Of String, Dictionary(Of String, List(Of String))), ByVal mainAssemblyName As String)
        Try
            Dim mainAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(mainAssemblyName, txtDirectoryPath.Text)

            Dim mainAssemblyNameWorking As String = $"{mainAssemblyName}_Working"

            Dim mainAssemblyDocWorking As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(mainAssemblyNameWorking, txtDirectoryPath.Text)

            mainAssemblyDocWorking.Save()
            mainAssemblyDocWorking.Close(True)

            Dim mainAssemblyDocWorkingPath As String = IO.Path.Combine(txtDirectoryPath.Text, $"{mainAssemblyNameWorking}.asm")

            AddAssemblyOccurences(mainAssemblyDoc, mainAssemblyDocWorkingPath, Nothing)

            Dim dicAssemblyChild As New Dictionary(Of String, List(Of String))()

            For Each kvp1 As KeyValuePair(Of String, Dictionary(Of String, List(Of String))) In dicMainAssemblyDetails

                If Not kvp1.Key.ToString().ToUpper().Trim() = "GEN CAR" Then
                    Continue For
                End If

                Dim firstLevelAssemblyName As String = kvp1.Key
                Dim firstLevelAssemblyNameWorking As String = $"{ kvp1.Key}_Working"
                If firstLevelAssemblyName.Contains("#") Then
                    Continue For
                End If

                Dim firsLevelAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(firstLevelAssemblyName, txtDirectoryPath.Text)
                Dim firstLevelAssemblyDocPath As String = IO.Path.Combine(txtDirectoryPath.Text, $"{firstLevelAssemblyName}.asm")
                AddAssemblyOccurences(mainAssemblyDoc, firstLevelAssemblyDocPath, dicAssemblyChild)

                'Dim firsLevelAssemblyDoc_Working As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(firstLevelAssemblyNameWorking, txtDirectoryPath.Text)
                'firsLevelAssemblyDoc_Working.Save()
                'firsLevelAssemblyDoc_Working.Close(True)

                'Dim firstLevelAssemblyDocPathWorking As String = IO.Path.Combine(txtDirectoryPath.Text, $"{firstLevelAssemblyNameWorking}.asm")
                'AddAssemblyOccurences(firsLevelAssemblyDoc, firstLevelAssemblyDocPathWorking, dicAssemblyChild)



                'Add user assembly path
                If chkAddUserAssembly.Checked Then
                    AddAssemblyOccurences(mainAssemblyDoc, txtfilepath.Text, dicAssemblyChild)
                End If


                firsLevelAssemblyDoc.Save()

                For Each kvp2 As KeyValuePair(Of String, List(Of String)) In kvp1.Value


                    Dim secondLevelAssemblyName As String = kvp2.Key
                    Dim secondLevelAssemblyNameWorking As String = $"{ kvp2.Key}_Working"

                    If secondLevelAssemblyName.Contains("#") Then
                        Continue For
                    End If

                    Dim secondLevelAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(secondLevelAssemblyName, txtDirectoryPath.Text)
                    Dim secondLevelAssemblyDocPath As String = IO.Path.Combine(txtDirectoryPath.Text, $"{secondLevelAssemblyName}.asm")
                    AddAssemblyOccurences(firsLevelAssemblyDoc, secondLevelAssemblyDocPath, dicAssemblyChild)

                    Dim secondLevelAssemblyDocWorking As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(secondLevelAssemblyNameWorking, txtDirectoryPath.Text)
                    secondLevelAssemblyDocWorking.Save()
                    secondLevelAssemblyDocWorking.Close(True)
                    Dim secondLevelAssemblyDocPathWorking As String = IO.Path.Combine(txtDirectoryPath.Text, $"{secondLevelAssemblyNameWorking}.asm")
                    AddAssemblyOccurences(secondLevelAssemblyDoc, secondLevelAssemblyDocPathWorking, dicAssemblyChild)

                    'Add user assembly path
                    If chkAddUserAssembly.Checked Then
                        AddAssemblyOccurences(firsLevelAssemblyDoc, txtfilepath.Text, dicAssemblyChild)
                    End If

                    secondLevelAssemblyDoc.Save()

                    For Each thirdLevelAssemblyName As String In kvp2.Value

                        Dim thirdLevelAssemblyNameWorking As String = $"{ thirdLevelAssemblyName}_Working"

                        If thirdLevelAssemblyName.Contains("#") Then
                            Continue For
                        End If

                        Dim thirdLevelAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(thirdLevelAssemblyName, txtDirectoryPath.Text)
                        Dim thirdtLevelAssemblyDocPath As String = IO.Path.Combine(txtDirectoryPath.Text, $"{thirdLevelAssemblyName}.asm")
                        Dim thirdLevelAssemblyDocPath As String = IO.Path.Combine(txtDirectoryPath.Text, $"{thirdLevelAssemblyName}.asm")
                        AddAssemblyOccurences2(secondLevelAssemblyDoc, thirdLevelAssemblyDocPath, dicAssemblyChild)

                        'Dim thirdLevelAssemblyDocWorking As SolidEdgeAssembly.AssemblyDocument = AddAssembly2(thirdLevelAssemblyNameWorking, txtDirectoryPath.Text)
                        'thirdLevelAssemblyDocWorking.Save()
                        'thirdLevelAssemblyDocWorking.Close(True)

                        'Dim thirdLevelAssemblyDocPathWorking As String = IO.Path.Combine(txtDirectoryPath.Text, $"{thirdLevelAssemblyNameWorking}.asm")
                        'AddAssemblyOccurences2(thirdLevelAssemblyDoc, thirdLevelAssemblyDocPathWorking, dicAssemblyChild)


                        'Add user assembly path
                        If chkAddUserAssembly.Checked Then
                            AddAssemblyOccurences(thirdLevelAssemblyDoc, txtfilepath.Text, dicAssemblyChild)
                        End If


                        thirdLevelAssemblyDoc.Save()
                        thirdLevelAssemblyDoc.Close(True)



                    Next

                    secondLevelAssemblyDoc.Save()
                    secondLevelAssemblyDoc.Close(True)

                Next
                firsLevelAssemblyDoc.Save()
                firsLevelAssemblyDoc.Close(True)
            Next

            mainAssemblyDoc.Save()
            mainAssemblyDoc.Close(True)

        Catch ex As Exception
            MessageBox.Show($"While Create Gen Car :{ex.Message}{ex.StackTrace}", "Message")
            CustomLogUtil.Log("While Create Gen Car", ex.Message, ex.StackTrace)
        End Try
    End Sub
    Private Sub BtnCreateVirtaulAssembly_Click(sender As Object, e As EventArgs) Handles btnCreateVirtaulAssembly.Click
        Dim MainAsmPath = ""
        If Not IO.File.Exists(txtExcelPath.Text) Then
            MessageBox.Show($"Please select virtual assembly structure excel", "Select Excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        If Not IO.Directory.Exists(txtDirectoryPath.Text) Then
            MessageBox.Show($"Please select destination folder for assembly creation", "Select Folder", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        If chkAddUserAssembly.Checked Then
            If Not IO.File.Exists(txtfilepath.Text) Then
                MessageBox.Show($"Please select user assembly", "Select Assembly", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        End If
        WaitStartSave()
        Try
            CustomLogUtil.Heading($"Reading Virtual Assembly Excel :{txtExcelPath.Text}")
            Dim topLevelAssemblyNames As String = String.Empty
            'Dim dicMainAssemblyDetails As Dictionary(Of String, Dictionary(Of String, List(Of String))) =
            'ExcelUtilVirtualAssemblyCreation1.ReadVirtualAssemblyCreationExcel(txtExcelPath.Text, topLevelAssemblyNames)
            'topLevelAssemblyNames = topLevelAssemblyNames.ToUpper().Replace(" ", "").Replace("PUMPCAR-", "").Replace("GENCAR-", "")

            '------------------------------------------------------------------------------------------------------------------------
            'Dim dicMainAssemblyDetails1 As Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, List(Of String)))) =
            'ExcelUtilVirtualAssemblyCreation1.ReadVirtualAssemblyCreationExcelNew(txtExcelPath.Text, topLevelAssemblyNames)

            Dim dicMainAssemblyDetails As Dictionary(Of String, Dictionary(Of String, List(Of String))) =
            ExcelUtilVirtualAssemblyCreation1.ReadVirtualAssemblyCreationExcelNew(txtExcelPath.Text, topLevelAssemblyNames)

            '------------------------------------------------------------------------------------------------------------------------
            CustomLogUtil.Heading($"Creating Virtual Assembly")
            'CreateVirtualAssembly(dicMainAssemblyDetails, topLevelAssemblyNames)
            'CreateVirtualAssemblyNew(dicMainAssemblyDetails, topLevelAssemblyNames)
            '------------------------------------------------------------------------------------------------------------------------
            'temp11APR2023
            CreateVirtualAssemblyNew2(dicMainAssemblyDetails, topLevelAssemblyNames)
            '------------------------------------------------------------------------------------------------------------------------
            WaitEndSave()
            MsgBox("Process Done...")
            'SolidEdgeCommunity.OleMessageFilter.Unregister()
            application.Quit()

            Dim MainAssemblyDetails As Array = topLevelAssemblyNames.Split(New Char() {"/C"})

            Dim MainAsmName = MainAssemblyDetails(0)
            MainAsmName += ".asm"
            MainAsmPath = System.IO.Path.Combine(txtDirectoryPath.Text, MainAsmName)

            CustomLogUtil.Heading("Virtual Structure Form Process completed.....")

            If IO.Directory.Exists(txtDirectoryPath.Text) Then
                Process.Start(txtDirectoryPath.Text)
            End If
        Catch ex As Exception
            SolidEdgeCommunity.OleMessageFilter.Unregister()
            MessageBox.Show($"Error While Creating Virtual Assembly {ex.Message}{ex.StackTrace}", "Message")
            CustomLogUtil.Log("While Creating Virtual Assembly", ex.Message, ex.StackTrace)
        Finally
            'WaitEndSave()
        End Try
        If (System.IO.File.Exists(MainAsmPath)) Then
            Process.Start(MainAsmPath)
        End If
        'waitStartSave()
        'CreateAssemblyStructure(dicMainAssemblyDetails)
        'WaitEndSave()

    End Sub



    Private Sub CreateAssemblyStructure(ByVal dicMainAssemblyDetails As Dictionary(Of String, Dictionary(Of String, List(Of String))))
        Dim application As SolidEdgeFramework.Application = Nothing
        SolidEdgeCommunity.OleMessageFilter.Register()
        application = SolidEdgeCommunity.SolidEdgeUtils.Connect(True, False)
        Try

            For Each kvp As KeyValuePair(Of String, Dictionary(Of String, List(Of String))) In dicMainAssemblyDetails

                Dim assemlbyName As String = kvp.Key
                assemlbyName = assemlbyName.Replace("/", "")

                Dim documents As SolidEdgeFramework.Documents = Nothing
                Dim assemblyDocument As SolidEdgeAssembly.AssemblyDocument = Nothing

                documents = application.Documents
                assemblyDocument = DirectCast(documents.Add("SolidEdge.AssemblyDocument"), SolidEdgeAssembly.AssemblyDocument)
                assemblyDocument.Name = assemlbyName

                Dim dicSubAssembly As Dictionary(Of String, List(Of String)) = kvp.Value

                For Each kvp2 As KeyValuePair(Of String, List(Of String)) In dicSubAssembly

                    Dim subAssemblyName As String = kvp2.Key
                    subAssemblyName = subAssemblyName.Replace("/", "")
                    Dim childAssemblyList As List(Of String) = kvp2.Value
                    Dim subassemblydocument As SolidEdgeAssembly.AssemblyDocument = Nothing
                    Dim subassemblypath As String = txtDirectoryPath.Text + "\" + subAssemblyName + ".asm"
                    If Not IO.File.Exists(subassemblypath) Then
                        subassemblydocument = DirectCast(documents.Add("SolidEdge.AssemblyDocument"), SolidEdgeAssembly.AssemblyDocument)
                    Else
                        'subassemblydocument = application.Documents.Open(subassemblypath)

                        subassemblydocument = DirectCast(documents.Open(subassemblypath), SolidEdgeAssembly.AssemblyDocument)
                    End If

                    For Each childassembly As String In childAssemblyList
                        childassembly = childassembly.Replace("/", "")
                        Dim childpath As String = txtDirectoryPath.Text + "\" + childassembly + ".asm"

                        If Not IO.File.Exists(childpath) Then
                            Dim childassemblydocument As SolidEdgeAssembly.AssemblyDocument = DirectCast(documents.Add("SolidEdge.AssemblyDocument"), SolidEdgeAssembly.AssemblyDocument)
                            childassemblydocument.Name = childassembly
                            childassemblydocument.SaveAs(childpath)
                            childassemblydocument.Close(True)
                        End If

                        Dim subassemblyoccurance As SolidEdgeAssembly.Occurrences = subassemblydocument.Occurrences
                        Dim occuranceNewlyAdded As SolidEdgeAssembly.Occurrence = subassemblyoccurance.AddByFilename(childpath)
                    Next

                    If Not IO.File.Exists(subassemblypath) Then
                        subassemblydocument.SaveAs(subassemblypath)
                    End If
                    subassemblydocument.Close(True)
                    Dim mainoccurance As SolidEdgeAssembly.Occurrences = assemblyDocument.Occurrences
                    Dim mainoccuranceNewlyAdded As SolidEdgeAssembly.Occurrence = mainoccurance.AddByFilename(subassemblypath)

                Next

                Dim assemblypath As String = txtDirectoryPath.Text + "\" + assemlbyName + ".asm"
                assemblyDocument.SaveAs(assemblypath)
                assemblyDocument.Close(True)

            Next

            '  MessageBox.Show("Completed")

        Catch ex As Exception
            MessageBox.Show($"Error in create assembly structure {ex.Message} {vbNewLine} {ex.StackTrace}", "Message")
            CustomLogUtil.Log("in create assembly structure", ex.Message, ex.StackTrace)
        End Try
        SolidEdgeCommunity.OleMessageFilter.Unregister()
        application.Quit()
    End Sub

    Private Sub BtnDirectoryPath_Click(sender As Object, e As EventArgs) Handles btnDirectoryPath.Click
        'Dim MyFolderBrowser As New System.Windows.Forms.FolderBrowserDialog
        'Dim dlgResult As DialogResult = MyFolderBrowser.ShowDialog()
        txtDirectoryPath.Text = OutputDirPath()
        'If dlgResult = Windows.Forms.DialogResult.OK Then
        '    txtDirectoryPath.Text = MyFolderBrowser.SelectedPath
        'End If
    End Sub
    Private Function OutputDirPath() As String
        Dim folderpath As String = ""
        Try

            folderpath = BrowseFolderAdvanced()
            If Not folderpath = String.Empty Then

            End If
        Catch ex1 As Exception
            Try
                Dim MyFolderBrowser As New Ookii.Dialogs.VistaFolderBrowserDialog
                Dim dlgResult As DialogResult = MyFolderBrowser.ShowDialog
                If dlgResult = DialogResult.OK Then
                    folderpath = MyFolderBrowser.SelectedPath
                End If
            Catch ex As Exception
                If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
                    folderpath = FolderBrowserDialog1.SelectedPath

                End If
            End Try

        End Try
        Return folderpath
    End Function
    Public Shared Function BrowseFolderAdvanced() As String

        Dim folderpath As String = ""
        Try
            Dim BetterFolderBrowser As New BetterFolderBrowser With {
                .Title = "Select folders",
                .RootFolder = "C:\\",
                .Multiselect = False
            }
            If BetterFolderBrowser.ShowDialog() = DialogResult.OK Then
                folderpath = BetterFolderBrowser.SelectedFolder
            End If
        Catch ex As Exception
            MessageBox.Show($"Advanced Browse folder{vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}", "Message")
        End Try


        Return folderpath
    End Function

    Dim waitFormObj As Wait
    Public Sub WaitStartSave()
        '==Processing==
        Dim waitThread As System.Threading.Thread
        waitThread = New System.Threading.Thread(AddressOf LaunchWaitSave)
        waitThread.Start()
        Threading.Thread.Sleep(1000)
        waitFormObj.SetWaitMessage("In progress..")

        waitFormObj.SetProgressInformationVisibility(True)
        waitFormObj.SetProgressInformationMessage("")

        waitFormObj.SetProgressCountVisibility(True)
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

    Private Sub VirtualAssemblyStructureForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CustomLogUtil.Heading("Virtual Structure Form Open.....")
        SolidEdgeCommunity.OleMessageFilter.Register()
        application = SolidEdgeCommunity.SolidEdgeUtils.Connect(True, False)
        application.DisplayAlerts = False
        Me.Text += "1.0.69"

        txtDirectoryPath.Text = Config.configObj.virtualAssemblyOutputDirec

        '68 'Add the tag # to skip the assembly
        '69 ' Create the 68 and 69 assembly and isolate the gen and pump car
        'Main assembly name is dynamically changed through excel but need to kept standard
    End Sub

    Private Sub BtnBrowseUserAssembly_Click(sender As Object, e As EventArgs) Handles btnopenfile.Click
        Try
            Using dialog As New OpenFileDialog
                dialog.Filter = "Assembly files (*.asm)|*.asm"
                If dialog.ShowDialog() <> DialogResult.OK Then Return
                txtfilepath.Text = dialog.FileName
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error in browse assembly file {vbNewLine}{ex.Message}{vbNewLine}{ex.StackTrace}", "Browse Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try
    End Sub

    Private Sub ChkAddUserAssembly_CheckedChanged(sender As Object, e As EventArgs) Handles chkAddUserAssembly.CheckedChanged
        Try
            If Not chkAddUserAssembly.Checked Then
                txtfilepath.Enabled = False
                btnopenfile.Enabled = False
                lbl_ReferenceModel.Enabled = False
            Else
                txtfilepath.Enabled = True
                btnopenfile.Enabled = True
                lbl_ReferenceModel.Enabled = True
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub VirtualAssemblyStructureForm2_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'KillSolidEdgeProcess.Kill()
        If application.Visible = False Then
            application.Quit()
        End If
    End Sub
End Class