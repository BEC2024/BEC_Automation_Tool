Imports System.Xml
Imports System.Environment
Imports System.Reflection
Imports System.Collections.Generic

Public Class Config
    Inherits ConfigBase

    Public Shared configObj As Config

    Public m2MFile As String = String.Empty
    Public propseedFile As String = String.Empty
    Public authorFile As String = String.Empty

    Public virtualAssemblyOutputDirec As String = String.Empty
    Public becMaterialExcelPath As String = String.Empty
    Public interferenceExcludeMaterialExcelPath As String = String.Empty

    Public baselineDirectoryPath As String = String.Empty
    Public mtcMtrReportsExportDirLocation As String = String.Empty

    Public rawMaterialEstimationReportDirPath As String = String.Empty
    Public rawMaterialBomExcelPath As String = String.Empty

    Public solidEdgePartTemplateDirectory As String = String.Empty

    Public RoutingSequenceOutputDirectory As String = String.Empty

    Public MTCExcelPath As String = String.Empty

    Public MTRExcelPath As String = String.Empty

    Public RoutingSequenceExcelPath As String = String.Empty

    Public EmployeeExcelPath As String = String.Empty

    Public AutoSaveAuthor As Boolean = False
    Public LogOutputDirectory As String = String.Empty
    Public ConfigTxtFile As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\" + "BECAutomation" + "\ConfigFile.txt"
    Public Shared configFilePath1 As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\" + "BECAutomation" + "\ConfigProperties.xml"


    Sub New(ByVal configFilePath As String)


        Me.configFilePath = configFilePath
        'Dim congigDir = ConfigTxtFile
        'congigDir = congigDir.Replace("\ConfigFile.txt", "")
        'If Not System.IO.Directory.Exists(congigDir) Then
        '    System.IO.Directory.CreateDirectory(congigDir)
        'End If
        'If Not System.IO.File.Exists(ConfigTxtFile) Then

        '    Dim objWriter As New System.IO.StreamWriter(ConfigTxtFile)

        '    objWriter.Write(configFilePath1)
        '    objWriter.Close()
        'End If

        'configFilePath = My.Computer.FileSystem.ReadAllText(ConfigTxtFile)



        'Dim dlgR = MessageBox.Show("Confuguration Path :" + configFilePath, "Do you want to change config path ?", MessageBoxButtons.YesNo)

        'Dim FolderBrowserDialog1 As New FolderBrowserDialog()
        'If dlgR = DialogResult.Yes Then
        '    If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
        '        configFilePath = FolderBrowserDialog1.SelectedPath
        '        configFilePath = System.IO.Path.Combine(configFilePath, "ConfigProperties.xml")
        '        System.IO.File.Delete(ConfigTxtFile)
        '        Dim objWriter As New System.IO.StreamWriter(ConfigTxtFile)

        '        objWriter.Write(configFilePath)
        '        objWriter.Close()
        '    Else
        '        MessageBox.Show("Configuration file unsaved", "No Changes")
        '    End If
        'End If
        'If System.IO.File.Exists(configFilePath) Then
        '    System.IO.File.Delete(configFilePath)
        'End If
        'System.IO.File.Copy(configFilePath1, configFilePath)
        'Me.configFilePath = configFilePath

        Try
            readConfig()
        Catch ex As Exception
            MessageBox.Show("Error in configuration read", "Error")
            CustomLogUtil.Log("Error in congiguration Read", ex.Message, ex.StackTrace)
        End Try
    End Sub
    Protected Overrides Sub ChangeConfigTxt2()
        Dim congigDir = ConfigTxtFile
        congigDir = congigDir.Replace("\ConfigFile.txt", "")
        If Not System.IO.Directory.Exists(congigDir) Then
            System.IO.Directory.CreateDirectory(congigDir)
        End If
        If Not System.IO.File.Exists(ConfigTxtFile) Then

            Dim fd As OpenFileDialog = New OpenFileDialog()
            fd.Title = "Do you want to set Config File ?"
            fd.InitialDirectory = "C:\"
            fd.Filter = "Config Files|*.xml*"
            fd.FilterIndex = 2
            fd.RestoreDirectory = True
            If fd.ShowDialog() = DialogResult.OK Then
                configFilePath1 = fd.FileName
            Else
                saveConfig2()
            End If

            Dim objWriter As New System.IO.StreamWriter(ConfigTxtFile)

            objWriter.Write(configFilePath1)
            objWriter.Close()
        End If
        Me.configFilePath = configFilePath1
    End Sub
    Protected Overrides Sub ChangeConfig2()
        Dim congigDir = ConfigTxtFile
        congigDir = congigDir.Replace("\ConfigFile.txt", "")
        If Not System.IO.Directory.Exists(congigDir) Then
            System.IO.Directory.CreateDirectory(congigDir)
            Dim fd As OpenFileDialog = New OpenFileDialog()
            fd.Title = "Do you want to set Config File"
            fd.InitialDirectory = "C:\"
            fd.Filter = "Config Files|*.xml*"
            fd.FilterIndex = 2
            fd.RestoreDirectory = True
            If fd.ShowDialog() = DialogResult.OK Then
                configFilePath1 = fd.FileName
            Else
                saveConfig2()
            End If

            Dim objWriter As New System.IO.StreamWriter(ConfigTxtFile)

            objWriter.Write(configFilePath1)
            objWriter.Close()
            configFilePath = My.Computer.FileSystem.ReadAllText(ConfigTxtFile)
        Else
            Dim result As DialogResult = MessageBox.Show("Do You Want to Change Config File ?" + vbNewLine + vbNewLine + "Configuration Path :" + configFilePath, "Message", MessageBoxButtons.YesNoCancel)
            If result = DialogResult.Cancel Or result = DialogResult.No Then

            ElseIf result = DialogResult.Yes Then
                Dim fd As OpenFileDialog = New OpenFileDialog()
                fd.Title = "Select Config File"
                fd.InitialDirectory = "C:\"
                fd.Filter = "Config Files|*.xml*"
                fd.FilterIndex = 2
                fd.RestoreDirectory = True
                If fd.ShowDialog() = DialogResult.OK Then
                    configFilePath1 = fd.FileName
                Else
                    Exit Sub
                End If

                Dim objWriter As New System.IO.StreamWriter(ConfigTxtFile)

                objWriter.Write(configFilePath1)
                objWriter.Close()
            End If
        End If
        'If Not System.IO.File.Exists(ConfigTxtFile) Then


        'End If

        configFilePath = My.Computer.FileSystem.ReadAllText(ConfigTxtFile)


        If Not System.IO.File.Exists(ConfigTxtFile) Then
            saveConfig2()
        End If


        'Dim dlgR = MessageBox.Show("Configuration Path :" + configFilePath, "Do you want to change config path ?", MessageBoxButtons.YesNo)

        'Dim FolderBrowserDialog1 As New FolderBrowserDialog()
        'If dlgR = DialogResult.Yes Then
        '    If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
        '        configFilePath = FolderBrowserDialog1.SelectedPath
        '        configFilePath = System.IO.Path.Combine(configFilePath, "ConfigProperties.xml")
        '        System.IO.File.Delete(ConfigTxtFile)
        '        Dim objWriter As New System.IO.StreamWriter(ConfigTxtFile)

        '        objWriter.Write(configFilePath)
        '        objWriter.Close()

        '        If System.IO.File.Exists(configFilePath) Then
        '            System.IO.File.Delete(configFilePath)
        '        End If

        '        System.IO.File.Copy(configFilePath1, configFilePath)
        '    Else
        '        MessageBox.Show("Configuration file unsaved", "No Changes")
        '    End If
        'End If

        Me.configFilePath = configFilePath

    End Sub

    Protected Overrides Sub saveConfig2()

        Dim doc As New Xml.XmlDocument
        Dim rootElement As Xml.XmlElement = doc.CreateElement("ROOT")

        'clientName
        Dim m2MFileElement As Xml.XmlElement = doc.CreateElement("m2MFile")
        rootElement.AppendChild(m2MFileElement)
        m2MFileElement.InnerText = m2MFile.ToString

        'vGrooveBendRadius
        Dim propseedFileElement As Xml.XmlElement = doc.CreateElement("propseedFile")
        rootElement.AppendChild(propseedFileElement)
        propseedFileElement.InnerText = propseedFile.ToString

        'vGrooveKFactor
        Dim authorFileElement As Xml.XmlElement = doc.CreateElement("authorFile")
        rootElement.AppendChild(authorFileElement)
        authorFileElement.InnerText = authorFile.ToString

        'virtualAssemblyOutputDirec
        Dim virtualAssemblyOutputDirecElement As Xml.XmlElement = doc.CreateElement("virtualAssemblyOutputDirec")
        rootElement.AppendChild(virtualAssemblyOutputDirecElement)
        virtualAssemblyOutputDirecElement.InnerText = virtualAssemblyOutputDirec.ToString

        'becMaterialExcelPath
        Dim becMaterialExcelPathElement As Xml.XmlElement = doc.CreateElement("becMaterialExcelPath")
        rootElement.AppendChild(becMaterialExcelPathElement)
        becMaterialExcelPathElement.InnerText = becMaterialExcelPath.ToString

        'interferenceExcludeMaterialExcelPath
        Dim interferenceExcludeMaterialExcelPathElement As Xml.XmlElement = doc.CreateElement("interferenceExcludeMaterialExcelPath")
        rootElement.AppendChild(interferenceExcludeMaterialExcelPathElement)
        interferenceExcludeMaterialExcelPathElement.InnerText = interferenceExcludeMaterialExcelPath.ToString


        'baselineDirectoryPath
        Dim baselineDirectoryPathElement As Xml.XmlElement = doc.CreateElement("baselineDirectoryPath")
        rootElement.AppendChild(baselineDirectoryPathElement)
        baselineDirectoryPathElement.InnerText = baselineDirectoryPath.ToString

        'mtcMtrReportsExportDirLocation
        Dim mtcMtrReportsExportDirLocationElement As Xml.XmlElement = doc.CreateElement("mtcMtrReportsExportDirLocation")
        rootElement.AppendChild(mtcMtrReportsExportDirLocationElement)
        mtcMtrReportsExportDirLocationElement.InnerText = mtcMtrReportsExportDirLocation.ToString


        'rawMaterialEstimationReportDirPath
        Dim rawMaterialEstimationReportDirPathElement As Xml.XmlElement = doc.CreateElement("rawMaterialEstimationReportDirPath")
        rootElement.AppendChild(rawMaterialEstimationReportDirPathElement)
        rawMaterialEstimationReportDirPathElement.InnerText = rawMaterialEstimationReportDirPath.ToString

        'rawMaterialBomExcelPath
        Dim rawMaterialBomExcelPathElement As Xml.XmlElement = doc.CreateElement("rawMaterialBomExcelPath")
        rootElement.AppendChild(rawMaterialBomExcelPathElement)
        rawMaterialBomExcelPathElement.InnerText = rawMaterialBomExcelPath.ToString

        'solidEdgePartTemplateDirectory
        Dim solidEdgePartTemplateDirectoryElement As Xml.XmlElement = doc.CreateElement("solidEdgePartTemplateDirectory")
        rootElement.AppendChild(solidEdgePartTemplateDirectoryElement)
        solidEdgePartTemplateDirectoryElement.InnerText = solidEdgePartTemplateDirectory.ToString



        'RoutingSequenceOutputDirectory
        Dim RoutingSequenceOutputDirectoryElement As Xml.XmlElement = doc.CreateElement("RoutingSequenceOutputDirectory")
        rootElement.AppendChild(RoutingSequenceOutputDirectoryElement)
        RoutingSequenceOutputDirectoryElement.InnerText = RoutingSequenceOutputDirectory.ToString
        doc.AppendChild(rootElement)


        'MTCExcelPath
        Dim MTCExcelPathElement As Xml.XmlElement = doc.CreateElement("MTCExcelPath")
        rootElement.AppendChild(MTCExcelPathElement)
        MTCExcelPathElement.InnerText = MTCExcelPath.ToString
        doc.AppendChild(rootElement)

        'MTRExcelPath
        Dim MTRExcelPathElement As Xml.XmlElement = doc.CreateElement("MTRExcelPath")
        rootElement.AppendChild(MTRExcelPathElement)
        MTRExcelPathElement.InnerText = MTRExcelPath.ToString
        doc.AppendChild(rootElement)

        'RoutingSequenceExcelPath
        Dim RoutingSequenceExcelPathElement As Xml.XmlElement = doc.CreateElement("RoutingSequenceExcelPath")
        rootElement.AppendChild(RoutingSequenceExcelPathElement)
        RoutingSequenceExcelPathElement.InnerText = RoutingSequenceExcelPath.ToString
        doc.AppendChild(rootElement)

        'EmployeeExcelPath
        Dim EmployeeExcelPathElement As Xml.XmlElement = doc.CreateElement("EmployeeExcelPath")
        rootElement.AppendChild(EmployeeExcelPathElement)
        EmployeeExcelPathElement.InnerText = EmployeeExcelPath.ToString
        doc.AppendChild(rootElement)

        'AutoSaveAuthor
        Dim AutoSaveAuthorElement As Xml.XmlElement = doc.CreateElement("AutoSaveAuthor")
        rootElement.AppendChild(AutoSaveAuthorElement)
        AutoSaveAuthorElement.InnerText = AutoSaveAuthor.ToString
        doc.AppendChild(rootElement)

        'LogOutputDirectory
        Dim LogOutputDirectoryElement As Xml.XmlElement = doc.CreateElement("LogOutputDirectory")
        rootElement.AppendChild(LogOutputDirectoryElement)
        LogOutputDirectoryElement.InnerText = LogOutputDirectory.ToString
        doc.AppendChild(rootElement)

        doc.Save(configFilePath)
    End Sub

    'temp29Jan2021
    Protected Overrides Sub readConfig2()
        If (System.IO.File.Exists(ConfigTxtFile)) Then
            configFilePath = My.Computer.FileSystem.ReadAllText(ConfigTxtFile)
        Else
            ChangeConfigTxt2()
            configFilePath = My.Computer.FileSystem.ReadAllText(ConfigTxtFile)
        End If


        Dim doc As New Xml.XmlDocument
        doc.Load(configFilePath)

        'm2MFile
        Dim m2MFileElement As XmlNodeList = doc.GetElementsByTagName("m2MFile")
        If Not m2MFileElement.Count = 0 Then
            If Not m2MFileElement.ItemOf(0).InnerText = Nothing Then
                m2MFile = m2MFileElement.ItemOf(0).InnerText
            End If
        End If



        'propseedFile
        Dim propseedFileElement As XmlNodeList = doc.GetElementsByTagName("propseedFile")
        If Not propseedFileElement.Count = 0 Then
            If Not propseedFileElement.ItemOf(0).InnerText = Nothing Then
                propseedFile = propseedFileElement.ItemOf(0).InnerText
            End If
        End If


        'authorFile
        Dim authorFileElement As XmlNodeList = doc.GetElementsByTagName("authorFile")
        If Not authorFileElement.Count = 0 Then
            If Not authorFileElement.ItemOf(0).InnerText = Nothing Then
                authorFile = authorFileElement.ItemOf(0).InnerText
            End If
        End If

        'virtualAssemblyOutputDirec
        Dim virtualAssemblyOutputDirecElement As XmlNodeList = doc.GetElementsByTagName("virtualAssemblyOutputDirec")
        If Not virtualAssemblyOutputDirecElement.Count = 0 Then
            If Not virtualAssemblyOutputDirecElement.ItemOf(0).InnerText = Nothing Then
                virtualAssemblyOutputDirec = virtualAssemblyOutputDirecElement.ItemOf(0).InnerText
            End If
        End If

        'becMaterialExcelPath
        Dim becMaterialExcelPathElement As XmlNodeList = doc.GetElementsByTagName("becMaterialExcelPath")
        If Not becMaterialExcelPathElement.Count = 0 Then
            If Not becMaterialExcelPathElement.ItemOf(0).InnerText = Nothing Then
                becMaterialExcelPath = becMaterialExcelPathElement.ItemOf(0).InnerText
            End If
        End If


        'interferenceExcludeMaterialExcelPath
        Dim interferenceExcludeMaterialExcelPathElement As XmlNodeList = doc.GetElementsByTagName("interferenceExcludeMaterialExcelPath")
        If Not interferenceExcludeMaterialExcelPathElement.Count = 0 Then
            If Not interferenceExcludeMaterialExcelPathElement.ItemOf(0).InnerText = Nothing Then
                interferenceExcludeMaterialExcelPath = interferenceExcludeMaterialExcelPathElement.ItemOf(0).InnerText
            End If
        End If


        'baselineDirectoryPath
        Dim baselineDirectoryPathElement As XmlNodeList = doc.GetElementsByTagName("baselineDirectoryPath")
        If Not baselineDirectoryPathElement.Count = 0 Then
            If Not baselineDirectoryPathElement.ItemOf(0).InnerText = Nothing Then
                baselineDirectoryPath = baselineDirectoryPathElement.ItemOf(0).InnerText
            End If
        End If

        'mtcMtrReportsExportDirLocation
        Dim mtcMtrReportsExportDirLocationElement As XmlNodeList = doc.GetElementsByTagName("mtcMtrReportsExportDirLocation")
        If Not mtcMtrReportsExportDirLocationElement.Count = 0 Then
            If Not mtcMtrReportsExportDirLocationElement.ItemOf(0).InnerText = Nothing Then
                mtcMtrReportsExportDirLocation = mtcMtrReportsExportDirLocationElement.ItemOf(0).InnerText
            End If
        End If


        'rawMaterialEstimationReportDirPath

        Dim rawMaterialEstimationReportDirPathElement As XmlNodeList = doc.GetElementsByTagName("rawMaterialEstimationReportDirPath")
        If Not rawMaterialEstimationReportDirPathElement.Count = 0 Then
            If Not rawMaterialEstimationReportDirPathElement.ItemOf(0).InnerText = Nothing Then
                rawMaterialEstimationReportDirPath = rawMaterialEstimationReportDirPathElement.ItemOf(0).InnerText
            End If
        End If

        'rawMaterialBomExcelPath

        Dim rawMaterialBomExcelPathElement As XmlNodeList = doc.GetElementsByTagName("rawMaterialBomExcelPath")
        If Not rawMaterialBomExcelPathElement.Count = 0 Then
            If Not rawMaterialBomExcelPathElement.ItemOf(0).InnerText = Nothing Then
                rawMaterialBomExcelPath = rawMaterialBomExcelPathElement.ItemOf(0).InnerText
            End If
        End If



        'solidEdgePartTemplateDirectory
        Dim solidEdgePartTemplateDirectoryElement As XmlNodeList = doc.GetElementsByTagName("solidEdgePartTemplateDirectory")
        If Not solidEdgePartTemplateDirectoryElement.Count = 0 Then
            If Not solidEdgePartTemplateDirectoryElement.ItemOf(0).InnerText = Nothing Then
                solidEdgePartTemplateDirectory = solidEdgePartTemplateDirectoryElement.ItemOf(0).InnerText
            End If
        End If

        'RoutingSequenceOutputDirectory
        Dim RoutingSequenceOutputDirectoryElement As XmlNodeList = doc.GetElementsByTagName("RoutingSequenceOutputDirectory")
        If Not RoutingSequenceOutputDirectoryElement.Count = 0 Then
            If Not RoutingSequenceOutputDirectoryElement.ItemOf(0).InnerText = Nothing Then
                RoutingSequenceOutputDirectory = RoutingSequenceOutputDirectoryElement.ItemOf(0).InnerText
            End If
        End If

        'MTCExcelPath
        Dim MTCExcelPathElement As XmlNodeList = doc.GetElementsByTagName("MTCExcelPath")
        If Not MTCExcelPathElement.Count = 0 Then
            If Not MTCExcelPathElement.ItemOf(0).InnerText = Nothing Then
                MTCExcelPath = MTCExcelPathElement.ItemOf(0).InnerText
            End If
        End If


        'MTCExcelPath
        Dim MTRExcelPathElement As XmlNodeList = doc.GetElementsByTagName("MTRExcelPath")
        If Not MTRExcelPathElement.Count = 0 Then
            If Not MTRExcelPathElement.ItemOf(0).InnerText = Nothing Then
                MTRExcelPath = MTRExcelPathElement.ItemOf(0).InnerText
            End If
        End If

        'RoutingSequenceExcelPath
        Dim RoutingSequenceExcelPathElement As XmlNodeList = doc.GetElementsByTagName("RoutingSequenceExcelPath")
        If Not RoutingSequenceExcelPathElement.Count = 0 Then
            If Not RoutingSequenceExcelPathElement.ItemOf(0).InnerText = Nothing Then
                RoutingSequenceExcelPath = RoutingSequenceExcelPathElement.ItemOf(0).InnerText
            End If
        End If

        'EmployeeExcelPath
        Dim EmployeeExcelPathElement As XmlNodeList = doc.GetElementsByTagName("EmployeeExcelPath")
        If Not EmployeeExcelPathElement.Count = 0 Then
            If Not EmployeeExcelPathElement.ItemOf(0).InnerText = Nothing Then
                EmployeeExcelPath = EmployeeExcelPathElement.ItemOf(0).InnerText
            End If
        End If

        'AutoSaveAuthor
        Dim AutoSaveAuthorElement As XmlNodeList = doc.GetElementsByTagName("AutoSaveAuthor")
        If Not AutoSaveAuthorElement.Count = 0 Then
            If Not AutoSaveAuthorElement.ItemOf(0).InnerText = Nothing Then
                AutoSaveAuthor = AutoSaveAuthorElement.ItemOf(0).InnerText
            End If
        End If

        'LogOutputDirectory
        Dim LogOutputDirectoryElement As XmlNodeList = doc.GetElementsByTagName("LogOutputDirectory")
        If Not LogOutputDirectoryElement.Count = 0 Then
            If Not LogOutputDirectoryElement.ItemOf(0).InnerText = Nothing Then
                LogOutputDirectory = LogOutputDirectoryElement.ItemOf(0).InnerText
            End If
        End If
    End Sub

    Public Sub updatePartNameToConfigFile(ByVal lstPartNames1 As List(Of String))

        Dim configFilename As String = configFilePath1
        Dim doc As New Xml.XmlDocument
        doc.Load(configFilename)

        Dim partNamesElements As XmlNode = doc.GetElementsByTagName("PartNames").ItemOf(0)
        partNamesElements.RemoveAll()

        For Each partName As String In lstPartNames1
            If partName = "" Then
                Continue For
            End If
            Dim partNameElement As Xml.XmlElement = doc.CreateElement("PartName")
            partNameElement.InnerText = partName
            partNamesElements.AppendChild(partNameElement)
        Next
        doc.Save(configFilename)

    End Sub
End Class
