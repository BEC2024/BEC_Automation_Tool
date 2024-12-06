Option Infer On

Imports System.IO
Imports System.Runtime.InteropServices

Public Class SolidEdgeUtil
    'Dim objApp As SolidEdgeFramework.Application = Nothing
    'Dim objDocument As SolidEdgeFramework.SolidEdgeDocument = Nothing
    'Dim objSheetMetalDocument As SolidEdgePart.SheetMetalDocument = Nothing
    'Dim partDocument As SolidEdgePart.PartDocument = Nothing

    'Dim objAssemblyDocument As SolidEdgeAssembly.AssemblyDocument
    'Dim dtAssemblyData As DataTable = Nothing
    'Dim objMatTable As SolidEdgeFramework.MatTable = Nothing

    Dim rstObj As New RountingSequenceClass
    Dim oApp As SolidEdgeFramework.Application = Nothing
    Dim oDocument As SolidEdgeFramework.SolidEdgeDocument = Nothing
    Dim oView As SolidEdgeFramework.View = Nothing
    Public Sub SetSolidEdgeInstance()
        Try
            'SolidEdgeCommunity.OleMessageFilter.Register()
            'oApp = SolidEdgeCommunity.SolidEdgeUtils.Connect(True, False)


            oApp = CreateObject("SolidEdge.Application")

            oApp.Visible = False

            oApp.DisplayAlerts = False
            'oApp = DirectCast(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
        Catch ex As Exception
            MessageBox.Show("Please open solidedge", "Message")
        End Try

    End Sub

    Public Sub OpenDocument(rstObj As RountingSequenceClass)
        SetSolidEdgeInstance()
        Try
            oApp.Documents.Open(rstObj.FilePath)


            oDocument = oApp.ActiveDocument

            oView = oApp.ActiveWindow.view
            oView.Fit()
            Dim dWidth = Screen.PrimaryScreen.WorkingArea.Width
            Dim dHeight = Screen.PrimaryScreen.WorkingArea.Height
            Dim mypath As New System.IO.DirectoryInfo(System.IO.Path.Combine(Environment.CurrentDirectory, ".."))
            If Not System.IO.Directory.Exists(mypath.FullName + "\Thumbnail") Then
                System.IO.Directory.CreateDirectory(mypath.FullName + "\Thumbnail")
            End If


            Dim sJpgFile As String = Nothing
            sJpgFile = mypath.FullName + "\Thumbnail\" + rstObj.PartName + ".jpg"
            'Dim AltViewStyle As Object = "Default"
            Dim AltViewStyle As Object = Nothing
            Dim Resolution As Object = 1
            Dim ColorDepth As Object = 24
            Dim ImageQuality = SolidEdgeFramework.SeImageQualityType.seImageQualityHigh
            Dim Invert As Boolean = False
            'oDocument.SaveAs(sJpgFile)
            'oView.SaveAsImage(sJpgFile, dWidth, dHeight, Resolution, ColorDepth, ImageQuality, Invert)
            If Not System.IO.File.Exists(sJpgFile) Then
                oView.SaveAsImage(sJpgFile, dWidth, dHeight, AltViewStyle, Resolution, ColorDepth, ImageQuality, Invert)
            End If

            rstObj.image = Image.FromFile(sJpgFile)
            oDocument.Save()
            oApp.Documents.Close()

            oApp.Quit()
            SolidEdgeCommunity.OleMessageFilter.Unregister()
            'releaseObject()
            'killProcess()
        Catch ex As Exception
            MessageBox.Show(ex.Message + ex.StackTrace, "Message")
        End Try

    End Sub
    'Public Function SetSolidEdgeInstance()
    '    Try
    '        ' SolidEdgeCommunity.OleMessageFilter.Register()
    '        'objApp = SolidEdgeCommunity.SolidEdgeUtils.Connect(True, False)
    '        objApp = DirectCast(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
    '    Catch ex As Exception
    '        MsgBox($"Error in fetching the Solid-Edge instance {ex.Message} {vbNewLine} {ex.StackTrace}")
    '    End Try
    'End Function
    'Public Function OpenDocument(rstObj As RountingSequenceClass)
    '    Try
    '        SetSolidEdgeInstance()
    '        objApp.Documents.Open(rstObj.FilePath)

    '        objDocument = objApp.ActiveDocument
    '        Dim Filename() As String = rstObj.FilePath.Split("\"c)
    '        Dim FinalName() As String = Filename(Filename.Count - 1).Split("."c)
    '        rstObj.image = Filename(0) + "\" + Filename(1) + "\" + FinalName(0) + ".pdf"
    '        objDocument.SaveAs(rstObj.image,, ".pdf")

    '        objApp.Documents.Close()
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message + ex.StackTrace, "Message")
    '    End Try

    'End Function







    'Public Function OpenDocument(rstObj As RountingSequenceClass)
    '    Dim application As SolidEdgeFramework.Application = Nothing
    '    Dim partDocument As SolidEdgePart.PartDocument = Nothing

    '    Try
    '        ' See "Handling 'Application is Busy' and 'Call was Rejected By Callee' errors" topic.
    '        SolidEdgeCommunity.OleMessageFilter.Register()

    '        ' Attempt to connect to a running instance of Solid Edge.
    '        application = DirectCast(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
    '        partDocument = CType(application.ActiveDocument, SolidEdgePart.PartDocument)

    '        If partDocument IsNot Nothing Then
    '            Dim NewName = rstObj.FilePath
    '            partDocument.SaveAs(NewName)
    '        End If
    '    Catch ex As System.Exception
    '        Console.WriteLine(ex)
    '    Finally
    '        SolidEdgeCommunity.OleMessageFilter.Unregister()
    '    End Try
    'End Function


    Public Sub OpenSEPart(rstObj As RountingSequenceClass)
        oApp = CreateObject("SolidEdge.Application")
        oApp.Visible = True
        oApp.Documents.Open(rstObj.FilePath)
        oDocument = oApp.ActiveDocument
    End Sub

#Region "KillProcess"
    Private Sub ReleaseObject()
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject("SolidEdge.Application")

        Catch ex As Exception

        Finally
            GC.Collect()
        End Try
    End Sub

    Public Sub KillProcess()
        Dim _proceses As Process()
        _proceses = Process.GetProcessesByName("SolidEdge.Application")
        For Each proces As Process In _proceses
            proces.Kill()
        Next
    End Sub

#End Region

End Class
