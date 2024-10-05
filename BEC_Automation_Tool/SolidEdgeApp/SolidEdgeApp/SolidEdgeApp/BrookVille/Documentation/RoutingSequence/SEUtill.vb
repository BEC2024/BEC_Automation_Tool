Imports System.Runtime.InteropServices

Public Class SEUtill
    Dim objApp As SolidEdgeFramework.Application = Nothing
    Dim objDocument As SolidEdgeFramework.SolidEdgeDocument = Nothing
    Dim objSheetMetalDocument As SolidEdgePart.SheetMetalDocument = Nothing
    Dim view As SolidEdgeFramework.View = Nothing
    Dim partDocument As SolidEdgePart.PartDocument = Nothing

    Dim objAssemblyDocument As SolidEdgeAssembly.AssemblyDocument
    Dim dtAssemblyData As DataTable = Nothing
    Dim objMatTable As SolidEdgeFramework.MatTable = Nothing
    Public Sub SetSolidEdgeInstance()
        Try

            SolidEdgeCommunity.OleMessageFilter.Register()
            objApp = SolidEdgeCommunity.SolidEdgeUtils.Connect(True, False)
            objApp.Visible = False
            objApp.DisplayAlerts = False
            objApp = DirectCast(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
        Catch ex As Exception
            MessageBox.Show("Please open solidedge", "Message")
        End Try

    End Sub
    Public Sub OpenDocument(rstObj As RountingSequenceClass)
        SetSolidEdgeInstance()
        Try
            objApp.Documents.Open(rstObj.FilePath)


            objDocument = objApp.ActiveDocument

            view = objApp.ActiveWindow.view
            view.Fit()
            Dim dWidth = Screen.PrimaryScreen.WorkingArea.Width
            Dim dHeight = Screen.PrimaryScreen.WorkingArea.Height
            'Dim mypath As New System.IO.DirectoryInfo(System.IO.Path.Combine(Environment.CurrentDirectory, ".."))
            Dim mypath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\" + "BECAutomation"
            'If Not System.IO.Directory.Exists(mypath.FullName + "\Thumbnail") Then
            '    System.IO.Directory.CreateDirectory(mypath.FullName + "\Thumbnail")
            'End If
            If Not System.IO.Directory.Exists(mypath + "\Thumbnail") Then
                System.IO.Directory.CreateDirectory(mypath + "\Thumbnail")
            End If

            Dim sJpgFile As String = Nothing
            'sJpgFile = mypath.FullName + "\Thumbnail\" + rstObj.PartName + ".jpg"
            sJpgFile = mypath + "\Thumbnail\" + rstObj.PartName + ".jpg"
            'Dim AltViewStyle As Object = "Default"
            Dim AltViewStyle As Object = Nothing
            Dim Resolution As Object = 1
            Dim ColorDepth As Object = 24
            Dim ImageQuality = SolidEdgeFramework.SeImageQualityType.seImageQualityHigh
            Dim Invert As Boolean = False
            'oDocument.SaveAs(sJpgFile)
            'oView.SaveAsImage(sJpgFile, dWidth, dHeight, Resolution, ColorDepth, ImageQuality, Invert)
            If Not System.IO.File.Exists(sJpgFile) Then
                view.SaveAsImage(sJpgFile, dWidth, dHeight, AltViewStyle, Resolution, ColorDepth, ImageQuality, Invert)
            End If

            rstObj.image = Image.FromFile(sJpgFile)
            'objDocument.Save()
            objApp.Documents.Close()

            objApp.Quit()
            SolidEdgeCommunity.OleMessageFilter.Unregister()
            'releaseObject()
            'killProcess()
        Catch ex As Exception
            MessageBox.Show(ex.Message + ex.StackTrace, "Message")
        End Try

    End Sub
    Public Sub OpenSEPart(rstObj As RountingSequenceClass)
        objApp = CreateObject("SolidEdge.Application")
        objApp.Visible = True
        objApp.Documents.Open(rstObj.FilePath)
        objDocument = objApp.ActiveDocument
    End Sub

End Class
