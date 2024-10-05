Imports System.ComponentModel
Imports WK.Libraries.BetterFolderBrowserNS

Public Class ConfigurationForm
    Public success As Boolean = False


    Private Sub ConfigurationForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Not txtLogOutputDirectory.Text = "" Then
            CustomLogUtil.Heading("ConfigurationForm  Open.....")
        End If


        Me.Text = $"{Me.Text} ({GlobalEntity.Version}) "

        InitializeConfigData()

        Me.TableLayoutPanel8.BringToFront()

    End Sub

    Public Sub InitializeConfigData()
        Try

            Dim configPath1 As String
            Config.configObj = New Config(configPath1)

            Dim ConfigtxtFile As String = Config.configObj.ConfigTxtFile
            If (System.IO.File.Exists(ConfigtxtFile)) Then
                configPath1 = My.Computer.FileSystem.ReadAllText(ConfigtxtFile)
            Else
                Config.configObj.ChangeConfigTxt()
                configPath1 = Config.configFilePath1
            End If




            'Dim dlgR As DialogResult

            'dlgR = MessageBox.Show("Confuguration Path :" + configPath1, "Do you want to change config path ?", MessageBoxButtons.YesNo)

            '' then test it:
            'If dlgR = DialogResult.Yes Then
            '    If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            '        configPath1 = FolderBrowserDialog1.SelectedPath
            '    Else
            '        MessageBox.Show("Configuration file unsaved", "No Changes")
            '    End If

            '    configPath1 = System.IO.Path.Combine(configPath1, "ConfigProperties.xml")

            '    If (Not System.IO.File.Exists(configPath1)) Then
            '        System.IO.File.Copy(Config.configFilePath1, configPath1)

            '    End If
            'End If
            'Dim DirPath As String = IO.Path.GetDirectoryName(configPath1)
            'If Not IO.Directory.Exists(DirPath) Then
            '    IO.Directory.CreateDirectory(DirPath)
            'End If



            Config.configObj.readConfig()


            TextBoxM2mFile.Text = Config.configObj.m2MFile

            TextBoxPropseedFile.Text = Config.configObj.propseedFile

            TextBoxAuthorFile.Text = Config.configObj.authorFile

            txtVirtualAssemblyOutputDirPath.Text = Config.configObj.virtualAssemblyOutputDirec

            txtBecMaterialExcelPath.Text = Config.configObj.becMaterialExcelPath

            txtInterferenceExcludeMaterialExcelPath.Text = Config.configObj.interferenceExcludeMaterialExcelPath

            txtBaseLineDirectoryPath.Text = Config.configObj.baselineDirectoryPath

            txtMtcMtrReportExportDirLocation.Text = Config.configObj.mtcMtrReportsExportDirLocation

            txtRawMaterialEstimationReportDirPath.Text = Config.configObj.rawMaterialEstimationReportDirPath


            txtRawMaterialBomExcelPath.Text = Config.configObj.rawMaterialBomExcelPath

            TxtSolidEdgePartsTemplateDirectory.Text = Config.configObj.solidEdgePartTemplateDirectory

            txtRoutingSequenceOutputDirectory.Text = Config.configObj.RoutingSequenceOutputDirectory

            txtMTCExcelPath.Text = Config.configObj.MTCExcelPath
            txtMTRExcelPath.Text = Config.configObj.MTRExcelPath

            txtRoutingSequenceExcelPath.Text = Config.configObj.RoutingSequenceExcelPath

            txtEmployeeExcelPath.Text = Config.configObj.EmployeeExcelPath

            txtLogOutputDirectory.Text = Config.configObj.LogOutputDirectory
            ChkAutoSaveAuthor.Checked = Config.configObj.AutoSaveAuthor

            If txtLogOutputDirectory.Text = "" Then
                MessageBox.Show("Please Fill all Configurations", "Message")
            End If
        Catch ex As Exception
            MessageBox.Show("Error while set config data on form.", "Set config data", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CustomLogUtil.Log($"while set config data on form", ex.Message, ex.StackTrace)
        End Try
    End Sub



    Private Sub btnRawMaterialEstimationReportDirPath_Click(sender As Object, e As EventArgs) Handles btnRawMaterialEstimationReportDirPath.Click
        BrowseDirPath(txtRawMaterialEstimationReportDirPath)
    End Sub

    Private Sub BtnBrowseExportDir_Click(sender As Object, e As EventArgs) Handles BtnBrowseMtcMtrExportDir.Click
        BrowseDirPath(txtMtcMtrReportExportDirLocation)
        ' MonarchLog.Log($"Add MTC-MTR Output Directory : {txtMtcMtrReportExportDirLocation.Text}")
    End Sub

    Private Sub btnBrowseM2MFile_Click(sender As Object, e As EventArgs) Handles btnBrowseM2MFile.Click
        Try
            Using dialog As New OpenFileDialog
                'dialog.Filter = "Excel files (*.xlsx)|*.xlsx"
                If dialog.ShowDialog() <> DialogResult.OK Then Return
                TextBoxM2mFile.Text = dialog.FileName
            End Using
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnBrowseVirtualAssemblyDirectoryPath_Click(sender As Object, e As EventArgs) Handles btnBrowseVirtualAssemblyDirectoryPath.Click
        BrowseDirPath(txtVirtualAssemblyOutputDirPath)
    End Sub

    Private Sub BrowseDirPath(ByRef txt As System.Windows.Forms.TextBox)
        Try
            Dim folderpath As String = ""
            folderpath = browseFolderAdvanced()
            If Not folderpath = String.Empty Then
                txt.Text = folderpath
            End If
        Catch ex1 As Exception
            Try
                Dim MyFolderBrowser As Ookii.Dialogs.VistaFolderBrowserDialog = New Ookii.Dialogs.VistaFolderBrowserDialog
                Dim dlgResult As DialogResult = MyFolderBrowser.ShowDialog
                If dlgResult = DialogResult.OK Then
                    txt.Text = MyFolderBrowser.SelectedPath
                End If
            Catch ex As Exception
                If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
                    Dim path As String = FolderBrowserDialog1.SelectedPath
                    txt.Text = path
                End If
            End Try

        End Try
    End Sub

    Private Sub btnBrowseBECMaterialExcel_Click(sender As Object, e As EventArgs) Handles btnBrowseBECMaterialExcel.Click
        Try
            Using dialog As New OpenFileDialog
                '4th Sep 2024
                dialog.Filter = "Excel files (*.xltx)|*.xltx" '"Excel files (*.xlsx)|*.xlsx"
                If dialog.ShowDialog() <> DialogResult.OK Then Return
                txtBecMaterialExcelPath.Text = dialog.FileName
            End Using
        Catch ex As Exception

        End Try
    End Sub

    Public Shared Function browseFolderAdvanced() As String

        Dim folderpath As String = ""
        Try
            Dim BetterFolderBrowser As New BetterFolderBrowser()

            BetterFolderBrowser.Title = "Select folders"

            BetterFolderBrowser.RootFolder = "C:\\"

            BetterFolderBrowser.Multiselect = False
            If BetterFolderBrowser.ShowDialog() = DialogResult.OK Then
                folderpath = BetterFolderBrowser.SelectedFolder
            End If
        Catch ex As Exception
            MessageBox.Show($"Advanced Browse folder", "Error")
            CustomLogUtil.Log("Advance Browse folder", ex.Message, ex.StackTrace)
        End Try


        Return folderpath
    End Function

    Private Sub BtnBrowseBaselineDirPath_Click(sender As Object, e As EventArgs) Handles BtnBrowseBaselineDirPath.Click
        BrowseDirPath(txtBaseLineDirectoryPath)
    End Sub

    Private Sub btnBrowseInterferenceExcludeMaterialExcelPath_Click(sender As Object, e As EventArgs) Handles btnBrowseInterferenceExcludeMaterialExcelPath.Click
        Try
            Using dialog As New OpenFileDialog
                '4th Sep 2024
                dialog.Filter = "Excel files (*.xltx)|*.xltx" '"Excel files (*.xlsx)|*.xlsx"
                If dialog.ShowDialog() <> DialogResult.OK Then Return
                txtInterferenceExcludeMaterialExcelPath.Text = dialog.FileName

            End Using
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnBrowsePropseedFile_Click(sender As Object, e As EventArgs) Handles btnBrowsePropseedFile.Click
        Try
            Using dialog As New OpenFileDialog
                'dialog.Filter = "Excel files (*.xlsx)|*.xlsx"
                If dialog.ShowDialog() <> DialogResult.OK Then Return
                TextBoxPropseedFile.Text = dialog.FileName
            End Using
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnBrowseAuthorFile_Click(sender As Object, e As EventArgs) Handles btnBrowseAuthorFile.Click
        Try
            Using dialog As New OpenFileDialog
                'dialog.Filter = "Excel files (*.xlsx)|*.xlsx"
                If dialog.ShowDialog() <> DialogResult.OK Then Return
                TextBoxAuthorFile.Text = dialog.FileName
                'Log.Info($"Add Author File Path : {TextBoxAuthorFile.Text}")
            End Using
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnBrowseRawMaterialBomExcelPath_Click(sender As Object, e As EventArgs) Handles btnBrowseRawMaterialBomExcelPath.Click
        Try
            Using dialog As New OpenFileDialog
                dialog.Filter = "Excel files (*.xlsx)|*.xlsx"
                If dialog.ShowDialog() <> DialogResult.OK Then Return
                txtRawMaterialBomExcelPath.Text = dialog.FileName
                'Log.Info($"Add Raw Material BOM Excel Path: {txtRawMaterialBomExcelPath.Text}")
            End Using
        Catch ex As Exception

        End Try
    End Sub

    Private Sub BtnBrowseSolidEdgePartTemplateDirectory_Click(sender As Object, e As EventArgs) Handles BtnBrowseSolidEdgePartTemplateDirectory.Click
        BrowseDirPath(TxtSolidEdgePartsTemplateDirectory)
    End Sub

    Private Sub btnBrowseRoutingSequenceOutputDirectory_Click(sender As Object, e As EventArgs) Handles btnBrowseRoutingSequenceOutputDirectory.Click
        BrowseDirPath(txtRoutingSequenceOutputDirectory)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        BrowseDirPath(txtLogOutputDirectory)
        Config.configObj.LogOutputDirectory = txtLogOutputDirectory.Text
        Config.configObj.saveConfig()
        Try
            CustomLogUtil.Log($"Add Log Output Directory : {txtLogOutputDirectory.Text}")
        Catch ex As Exception
            MessageBox.Show($"Log Output Directory Could not Save ", "Error")
            CustomLogUtil.Log("Log Output Directory Could not Save ", ex.Message, ex.StackTrace)
            Application.Exit()
        End Try

    End Sub

    Private Sub btnBrowseMTCExcelPath_Click(sender As Object, e As EventArgs) Handles btnBrowseMTCExcelPath.Click
        Try
            Using dialog As New OpenFileDialog
                '4th Sep 2024
                dialog.Filter = "Excel files (*.xltx)|*.xltx" '"Excel files (*.xlsx)|*.xlsx"
                If dialog.ShowDialog() <> DialogResult.OK Then Return
                txtMTCExcelPath.Text = dialog.FileName
                'Log.Info($"Add Raw Material BOM Excel Path: {txtRawMaterialBomExcelPath.Text}")
            End Using
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnBrowseMTRExcelPath_Click(sender As Object, e As EventArgs) Handles btnBrowseMTRExcelPath.Click
        Try
            Using dialog As New OpenFileDialog
                dialog.Filter = "Excel files (*.xlsx)|*.xlsx"
                If dialog.ShowDialog() <> DialogResult.OK Then Return
                txtMTRExcelPath.Text = dialog.FileName
                'Log.Info($"Add Raw Material BOM Excel Path: {txtRawMaterialBomExcelPath.Text}")
            End Using
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnBrowseRoutingSequenceExcelPath_Click(sender As Object, e As EventArgs) Handles btnBrowseRoutingSequenceExcelPath.Click
        Try
            Using dialog As New OpenFileDialog
                '4th Sep 2024
                dialog.Filter = "Excel files (*.xltx)|*.xltx" '"Excel files (*.xlsx)|*.xlsx"
                If dialog.ShowDialog() <> DialogResult.OK Then Return
                txtRoutingSequenceExcelPath.Text = dialog.FileName
                'Log.Info($"Add Raw Material BOM Excel Path: {txtRawMaterialBomExcelPath.Text}")
            End Using
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TableLayoutPanel1_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel1.Paint

    End Sub



    Private Sub btn_login_Click(sender As Object, e As EventArgs) Handles btn_login.Click
        If txt_password.Text = "Bec@1234" Then
            success = True

            Me.TableLayoutPanel1.Enabled = True
            Me.TableLayoutPanel4.Visible = False

        Else
            MsgBox("Please Enter Valid Password")
            success = False
            Me.TableLayoutPanel1.Enabled = False
            Me.TableLayoutPanel4.Visible = True
        End If
    End Sub

    Private Sub btnEmployeeExcelPath_Click(sender As Object, e As EventArgs) Handles btnEmployeeExcelPath.Click
        Try
            Using dialog As New OpenFileDialog
                '4th Sep 2024
                dialog.Filter = "Excel files (*.xltx)|*.xltx" '"Excel files (*.xlsx)|*.xlsx"
                If dialog.ShowDialog() <> DialogResult.OK Then Return
                txtEmployeeExcelPath.Text = dialog.FileName
                'Log.Info($"Add Raw Material BOM Excel Path: {txtRawMaterialBomExcelPath.Text}")
            End Using
        Catch ex As Exception

        End Try
    End Sub
    Public Sub SaveBtnData()
        If txtLogOutputDirectory.Text = "" Then
            MessageBox.Show("Please Fill Log Output Directory..", "Message")
        Else
            Config.configObj.m2MFile = TextBoxM2mFile.Text
            Config.configObj.propseedFile = TextBoxPropseedFile.Text
            Config.configObj.authorFile = TextBoxAuthorFile.Text

            Config.configObj.virtualAssemblyOutputDirec = txtVirtualAssemblyOutputDirPath.Text
            Config.configObj.becMaterialExcelPath = txtBecMaterialExcelPath.Text
            Config.configObj.interferenceExcludeMaterialExcelPath = txtInterferenceExcludeMaterialExcelPath.Text
            Config.configObj.baselineDirectoryPath = txtBaseLineDirectoryPath.Text
            Config.configObj.mtcMtrReportsExportDirLocation = txtMtcMtrReportExportDirLocation.Text
            Config.configObj.rawMaterialEstimationReportDirPath = txtRawMaterialEstimationReportDirPath.Text

            Config.configObj.rawMaterialBomExcelPath = txtRawMaterialBomExcelPath.Text
            Config.configObj.solidEdgePartTemplateDirectory = TxtSolidEdgePartsTemplateDirectory.Text

            Config.configObj.RoutingSequenceOutputDirectory = txtRoutingSequenceOutputDirectory.Text

            Config.configObj.MTCExcelPath = txtMTCExcelPath.Text

            Config.configObj.MTRExcelPath = txtMTRExcelPath.Text

            Config.configObj.RoutingSequenceExcelPath = txtRoutingSequenceExcelPath.Text

            Config.configObj.LogOutputDirectory = txtLogOutputDirectory.Text

            Config.configObj.EmployeeExcelPath = txtEmployeeExcelPath.Text
            Config.configObj.AutoSaveAuthor = ChkAutoSaveAuthor.Checked
            Config.configObj.saveConfig()
            Me.Close()

            CustomLogUtil.Heading("Configuration Data save.....")

        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        SaveBtnData()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
#Region "Change Config 2 Buttons "
    '1. on config table layout
    Private Sub BtnChangeConfig1_Click(sender As Object, e As EventArgs) Handles BtnChangeConfig1.Click
        Config.configObj.changeConfig()
        InitializeConfigData()
    End Sub


    '2. right side of login table layout
    Private Sub BtnConfigChange2_Click(sender As Object, e As EventArgs) Handles BtnConfigChange2.Click
        Config.configObj.changeConfig()

        InitializeConfigData()

        Me.TableLayoutPanel8.BringToFront()
    End Sub

#End Region
End Class