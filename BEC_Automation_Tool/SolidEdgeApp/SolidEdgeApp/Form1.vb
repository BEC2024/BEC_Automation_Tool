Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports SolidEdgeFramework
Imports SolidEdgePart
Imports SeThumbnailLib
Imports RevisionManager
Public Class Form1

    Dim objApp As SolidEdgeFramework.Application = Nothing
    Dim objDocument As SolidEdgeFramework.SolidEdgeDocument = Nothing

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        DisplayThumbnail("C:\Users\vimalb\Downloads\Test Assembly\Test Assembly\1-2NCB82.par")


        'Dim fileInfo As FileInfo = New FileInfo("C:\Users\vimalb\Downloads\Test Assembly\Test Assembly\1-2NCB82.par")
        'Dim imag As Bitmap = Thumbnails.ExtractThumbNail(fileInfo)
        'Debug.Print("aa")
        'SolidEdgeCommunity.OleMessageFilter.Register()
        'objApp = SolidEdgeCommunity.SolidEdgeUtils.Connect(True, False)
        'objApp.Visible = False

        'Dim document As SolidEdgeFramework.SolidEdgeDocument = Nothing
        'Dim path As String = "C:\Users\vimalb\Downloads\Test Assembly\Test Assembly\1-2NCB82.par"
        'document = objApp.Documents.Open(path)
        'Dim bmp As Bitmap = document.get
        'Dim SEThumb As New SeThumbnailExtractor

        'Dim hImageSE As Integer

        'Dim imagePic As Image


        'SEThumb.GetThumbnail("C:\Users\vimalb\Downloads\Test Assembly\Test Assembly\1-2NCB82.par", hImageSE)

        'imagePic = Image.FromHbitmap(hImageSE)


        ''and now you can save the Image, example png

        'imagePic.Save("abc" & ".png", Imaging.ImageFormat.Png)

        'imagePic = Nothing

        'If Not SEThumb Is Nothing Then

        '    System.Runtime.InteropServices.Marshal.ReleaseComObject(SEThumb)

        '    SEThumb = Nothing

        'End If

    End Sub

    Private Sub DisplayThumbnail(ByVal filepath As String)
        Try
            Dim objApplication As RevisionManager.Application = Nothing
            objApplication = CreateObject("RevisionManager.Application")
            Dim doc As RevisionManager.Document = objApplication.Open(filepath)

            ' Dim img As Image = Compatibility.VB6.IPictureDispToImage(doc.Thumbnail)
            'PictureBox1.Image = img

        Catch ex As Exception
        End Try
    End Sub

    Private Sub SetSolidEdgeInstance()
        Try

            objApp = Marshal.GetActiveObject("SolidEdge.Application")

        Catch ex As Exception
            MsgBox($"Error in fetching the Solid-Edge instance {ex.Message} {vbNewLine} {ex.StackTrace}")

        End Try

    End Sub


End Class