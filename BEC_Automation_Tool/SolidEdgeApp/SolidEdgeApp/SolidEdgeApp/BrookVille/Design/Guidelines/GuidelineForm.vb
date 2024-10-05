Public Class GuidelineForm
    Public Enum GuideLineCat
        Flange
        Hem
        Louver
        Bend
    End Enum

    Dim imageNo As Integer = 0
    Dim currentCatName As String = String.Empty
    Dim totalImageCnt As Integer = 0
    Dim currentImageCnt As Integer = 0


    Private Sub SetButtonColor()

        If currentCatName = GuideLineCat.Flange.ToString() Then

            btnFlange.ForeColor = Color.Gray
            btnHem.ForeColor = Color.Black
            btnLouver.ForeColor = Color.Black
            btnBend.ForeColor = Color.Black

        ElseIf currentCatName = GuideLineCat.Hem.ToString() Then

            btnFlange.ForeColor = Color.Black
            btnHem.ForeColor = Color.Gray
            btnLouver.ForeColor = Color.Black
            btnBend.ForeColor = Color.Black

        ElseIf currentCatName = GuideLineCat.Louver.ToString() Then

            btnFlange.ForeColor = Color.Black
            btnHem.ForeColor = Color.Black
            btnLouver.ForeColor = Color.Gray
            btnBend.ForeColor = Color.Black

        ElseIf currentCatName = GuideLineCat.Bend.ToString() Then
            btnFlange.ForeColor = Color.Black
            btnHem.ForeColor = Color.Black
            btnLouver.ForeColor = Color.Black
            btnBend.ForeColor = Color.Gray
        End If

    End Sub

    Private Sub btnFlange_Click(sender As Object, e As EventArgs) Handles btnFlange.Click

        imageNo = 1

        currentCatName = GuideLineCat.Flange.ToString()

        SetImage()

        InitializeImageCnt()

        SetLabel()

        SetButtonColor()
    End Sub

    Private Sub btnHem_Click(sender As Object, e As EventArgs) Handles btnHem.Click

        imageNo = 1

        currentCatName = GuideLineCat.Hem.ToString()

        SetImage()

        InitializeImageCnt()

        SetLabel()

        SetButtonColor()
    End Sub


    Private Sub btnLouver_Click(sender As Object, e As EventArgs) Handles btnLouver.Click

        imageNo = 1

        currentCatName = GuideLineCat.Louver.ToString()

        SetImage()

        InitializeImageCnt()

        SetLabel()

        SetButtonColor()
    End Sub

    Private Sub btnBend_Click(sender As Object, e As EventArgs) Handles btnBend.Click

        imageNo = 1

        currentCatName = GuideLineCat.Bend.ToString()

        SetImage()

        InitializeImageCnt()

        SetLabel()

        SetButtonColor()
    End Sub

    Public Function GetImagePath(ByVal dirPath As String, ByVal catName As String, ByVal imageNo As Integer)

        Dim imagePath As String = IO.Path.Combine(dirPath, $"{catName}{imageNo.ToString()}.png")

        Return imagePath

    End Function

    Public Sub SetLabel()

        lblProgress.Text = $"{ currentImageCnt.ToString()}/{totalImageCnt.ToString()}"

    End Sub

    Public Sub InitializeImageCnt()

        Dim dirLocation As String = GetDirLocation(currentCatName)

        totalImageCnt = IO.Directory.GetFiles(dirLocation).Length

        currentImageCnt = 1

    End Sub

    Public Sub SetImage()


        Dim dirLocation As String = GetDirLocation(currentCatName)

        Dim imagePath As String = GetImagePath(dirLocation, currentCatName, imageNo)

        If IO.File.Exists(imagePath) Then
            PictureBox1.ImageLocation = (imagePath)
            PictureBox1.Load()

        Else
            imageNo = 1
            imagePath = GetImagePath(dirLocation, currentCatName, imageNo)
            PictureBox1.ImageLocation = (imagePath)
            PictureBox1.Load()
        End If


    End Sub

    Public Function GetDirLocation(ByVal catName As String) As String

        Dim path As String = System.Reflection.Assembly.GetExecutingAssembly().Location

        Dim dirPath As String = IO.Path.GetDirectoryName(path)

        Dim imagePath As String = IO.Path.Combine(dirPath, "Images")

        Dim guideLineDirPath As String = IO.Path.Combine(imagePath, catName)

        Return guideLineDirPath

    End Function

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click

        If Not currentImageCnt < totalImageCnt Then
            Exit Sub
        End If

        imageNo = imageNo + 1
        currentImageCnt = currentImageCnt + 1

        SetImage()

        SetLabel()

    End Sub

    Private Sub btnPrev_Click(sender As Object, e As EventArgs) Handles btnPrev.Click

        If Not currentImageCnt > 1 Then
            Exit Sub
        End If

        imageNo = imageNo - 1

        currentImageCnt = currentImageCnt - 1

        SetImage()

        SetLabel()
    End Sub

End Class