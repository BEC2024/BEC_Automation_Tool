Imports System.IO

Public Class RST_Preview
    Dim resSave As New RST_BL
    Dim rstObj As New RountingSequenceClass
    Private Sub btnOpenSEPart_Click(sender As Object, e As EventArgs) Handles btnOpenSEPart.Click
        btnOpenSEPart.Enabled = False
        rstObj.FilePath = Label1.Text
        resSave.OpenSEPart(rstObj)
    End Sub
    Public Sub OpenSEPart(rstObj As RountingSequenceClass)
        resSave.OpenSEPart(rstObj)
        PictureBox1.Image = rstObj.image
    End Sub

    Private Sub RST_Preview_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Public Sub DeleteDir()
        PictureBox1.Image.Dispose()
        PictureBox1.Image = Nothing
        'Dim mypath As New System.IO.DirectoryInfo(System.IO.Path.Combine(Environment.CurrentDirectory, ".."))
        'Dim files As String() = Directory.GetFiles(mypath.FullName + "\Thumbnail")
        'Dim dirs As String() = Directory.GetDirectories(mypath.FullName + "\Thumbnail")
        'For Each file As String In files

        '    System.IO.File.SetAttributes(file, FileAttributes.Normal)
        '    System.IO.File.Delete(file)
        'Next
        'Directory.Delete(mypath.FullName + "\Thumbnail")
    End Sub

    Private Sub RST_Preview_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        DeleteDir()
    End Sub

    Private Sub btnOpenSEPart_Click_1(sender As Object, e As EventArgs)

    End Sub

End Class