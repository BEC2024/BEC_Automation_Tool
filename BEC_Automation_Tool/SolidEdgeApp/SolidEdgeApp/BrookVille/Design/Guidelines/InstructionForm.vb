Public Class InstructionForm

    Dim imageName As String = String.Empty
    Sub New(ByVal imageName As String)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.imageName = imageName
    End Sub
    Private Sub InstructionForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim path As String = System.Reflection.Assembly.GetExecutingAssembly().Location

        Dim dirPath As String = IO.Path.GetDirectoryName(path)

        Dim imagePath As String = IO.Path.Combine(dirPath, "Images")

        imagePath = IO.Path.Combine(imagePath, imageName)


        If IO.File.Exists(imagePath) Then

            PictureBox1.ImageLocation = (imagePath)
            PictureBox1.Load()
        Else
            Me.Close()
        End If

    End Sub
End Class