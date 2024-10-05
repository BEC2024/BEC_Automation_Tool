Imports System.IO
Imports System.Runtime.InteropServices
Imports ExcelDataReader
Imports FontAwesome.Sharp
Imports SolidEdgeCommunity
Imports SolidEdgeFileProperties
Namespace MonarchTemplates

End Namespace
Public Class MainForm
    Dim count As Integer = 0
    ''https://www.youtube.com/watch?v=5AsJJl7Bhvc

#Region "Private Fields"
    Private currentBtn As IconButton
    Private ReadOnly leftBorderBtn As Panel


    ''This style rule concerns specifying the readonly (C#) or ReadOnly (Visual Basic) modifier for private fields that are initialized (either inline or inside of a constructor) but never reassigned.
    Private currentChildForm As Form
    Private Const CS_DropShadow As Integer = &H2000 '&H20000

    Public Structure RGBColors
        Public Shared color1 As Color = Color.FromArgb(172, 126, 241)
        Public Shared color2 As Color = Color.FromArgb(249, 118, 176)
        Public Shared color3 As Color = Color.FromArgb(253, 138, 114)
        Public Shared color4 As Color = Color.FromArgb(95, 77, 221)
        Public Shared color5 As Color = Color.FromArgb(249, 88, 155)
        Public Shared color6 As Color = Color.FromArgb(24, 161, 251)

        ''' <summary>
        ''' Panel backcolor
        ''' </summary>
        Public Shared color7 As Color = Color.FromArgb(40, 42, 145) '(43, 87, 154)

        ''' <summary>
        ''' panel highlight color
        ''' </summary>
        Public Shared color8 As Color = Color.FromArgb(43, 87, 154)

    End Structure

#End Region

#Region "Constructors"

    Public Sub New()

        ' This call is required by the designer.'
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.'


        leftBorderBtn = New Panel With {
            .Size = New Size(7, 60)
        }
        SideMenuPanel.Controls.Add(leftBorderBtn)


        'leftBorderBtn1 = New Panel With {
        '    .Size = New Size(7, 60)
        '}
        'panelAddUpdateSubMenu.Controls.Add(leftBorderBtn1)

        'Form'
        Me.Text = String.Empty
        Me.ControlBox = False
        Me.DoubleBuffered = True
        'Me.MaximizedBounds = Screen.PrimaryScreen.WorkingArea
        HideSubMenu()

        'AddHandler panelTitlBar.DoubleClick += AddressOf doublbCli

        'For Each c As Control In panelTitlBar.Controls
        '    If TypeOf Not c Is Button Then c.DoubleClick += AddressOf doublbCli
        'Next

    End Sub
    'Private Sub doublbCli(ByVal sender As Object, ByVal e As EventArgs)
    '    MessageBox.Show("info")
    'End Sub

#End Region

#Region "Private Methods"

    Private Sub ActivateButton(senderBtn As Object, customColor As Color)
        If senderBtn IsNot Nothing Then
            DisableButton()
            'Button'

            currentBtn = CType(senderBtn, IconButton)
            currentBtn.BackColor = Color.FromArgb(37, 36, 81)
            currentBtn.ForeColor = customColor
            currentBtn.IconColor = customColor
            'currentBtn.TextAlign = ContentAlignment.MiddleCenter
            'currentBtn.ImageAlign = ContentAlignment.MiddleRight
            'currentBtn.TextImageRelation = TextImageRelation.TextBeforeImage

            BtnValidation()

            'Left Border'
            leftBorderBtn.Size = New Size(7, currentBtn.Height)
            leftBorderBtn.BackColor = customColor
            leftBorderBtn.Location = New Point(currentBtn.Location.X, currentBtn.Location.Y)
            leftBorderBtn.Visible = True
            leftBorderBtn.BringToFront()

            'leftBorderBtn1.Size = New Size(7, currentBtn.Height)
            'leftBorderBtn1.BackColor = customColor
            'leftBorderBtn1.Location = New Point(currentBtn.Location.X, currentBtn.Location.Y)
            'leftBorderBtn1.Visible = True
            'leftBorderBtn1.BringToFront()

            'current Form icon'
            lblFormTitle.Text = currentBtn.Text
            'IconCurrentForm.IconChar = currentBtn.IconChar
            'IconCurrentForm.IconColor = customColor
            FormIconValidation()

            'reset on Main Menu button click
            MainFormResetValidation()


        End If
    End Sub
    Private Sub MainFormResetValidation()
        If currentChildForm IsNot Nothing And currentBtn.Text = "Add/ Update" Or currentChildForm IsNot Nothing And currentBtn.Text = "Design" Or currentChildForm IsNot Nothing And currentBtn.Text = "QC Report" Or currentChildForm IsNot Nothing And currentBtn.Text = "Documentation" Then
            currentChildForm.Close()
        End If
    End Sub
    Private Sub FormIconValidation()

        If panelAddUpdateSubMenu.Visible = True Then

            IconCurrentForm.IconChar = IconChar.Add
        End If

        If panelDesignSubMenu.Visible = True Then
            IconCurrentForm.IconChar = IconChar.DraftingCompass

        End If

        If panelQCSubMenu.Visible = True Then
            IconCurrentForm.IconChar = IconChar.CheckDouble

        End If


        If currentBtn.Text = "Configuration" Then

            IconCurrentForm.IconChar = IconChar.Wrench
        End If

    End Sub
    Private Sub BtnValidation()
        If currentBtn.Name.Contains("AddUpdate") And panelAddUpdateSubMenu.Visible = False Or panelAddUpdateSubMenu.Visible = True And currentBtn.Name = "BtnAddUpdate" Then
            SideMenuPanel.Controls.Add(leftBorderBtn)

        ElseIf panelAddUpdateSubMenu.Visible = True Then
            panelAddUpdateSubMenu.Controls.Add(leftBorderBtn)


        ElseIf currentBtn.Name.Contains("Design") And panelDesignSubMenu.Visible = False Or panelDesignSubMenu.Visible = True And currentBtn.Name = "BtnDesign" Then
            SideMenuPanel.Controls.Add(leftBorderBtn)

        ElseIf panelDesignSubMenu.Visible = True Then
            panelDesignSubMenu.Controls.Add(leftBorderBtn)


        ElseIf currentBtn.Name.Contains("QC") And panelQCSubMenu.Visible = False Or panelQCSubMenu.Visible = True And currentBtn.Name = "BtnQC" Then
            SideMenuPanel.Controls.Add(leftBorderBtn)

        ElseIf panelQCSubMenu.Visible = True Then
            panelQCSubMenu.Controls.Add(leftBorderBtn)


            'ElseIf currentBtn.Name.Contains("Documentation") And panelDocumentationSubMenu.Visible = False Or panelDocumentationSubMenu.Visible = True And currentBtn.Name = "btnDocumentation" Then
            '    SideMenuPanel.Controls.Add(leftBorderBtn)

            'ElseIf panelDocumentationSubMenu.Visible = True Then
            '    panelDocumentationSubMenu.Controls.Add(leftBorderBtn)
        End If
        BtnHelp.Visible = False
        LblHelp.Visible = False
    End Sub

    Private Sub DisableButton()
        If currentBtn IsNot Nothing Then
            currentBtn.BackColor = Color.FromArgb(31, 30, 68)
            currentBtn.ForeColor = Color.Gainsboro
            currentBtn.IconColor = Color.Gainsboro
            currentBtn.TextAlign = ContentAlignment.MiddleLeft
            currentBtn.ImageAlign = ContentAlignment.MiddleLeft
            currentBtn.TextImageRelation = TextImageRelation.ImageBeforeText
        End If
    End Sub

    Private Sub OpenOnlyForm()
        'Open only form'
        If currentChildForm IsNot Nothing Then
            currentChildForm.Close()
        End If
    End Sub

    Private Sub OpenChildForm(childForm As Form)

        OpenOnlyForm()
        currentChildForm = childForm
        childForm.TopLevel = False
        childForm.FormBorderStyle = FormBorderStyle.None
        childForm.Dock = DockStyle.Fill
        panelDesktop.Controls.Add(childForm)
        panelDesktop.Tag = childForm

        childForm.BringToFront()
        childForm.Show()

        'this

        If count = 1 Then
            lblFormTitle.Text = "Configuration"
            count += 1
        Else
            lblFormTitle.Text = currentBtn.Text
        End If

        BtnHelp.Visible = True
        LblHelp.Visible = True


    End Sub

    Private Sub Reset()
        DisableButton()
        leftBorderBtn.Visible = False
        IconCurrentForm.IconChar = IconChar.Home
        IconCurrentForm.IconColor = RGBColors.color6
        lblFormTitle.Text = "Home"

        BtnHelp.Visible = False
        LblHelp.Visible = False

        OpenOnlyForm()


    End Sub

    Private Sub ShowSubMenu(ByRef subMenu As Panel)
        If subMenu.Visible = False Then
            HideSubMenu()
            subMenu.Visible = True

        Else
            subMenu.Visible = False
        End If
    End Sub

    Private Sub HideSubMenu()


        panelQCSubMenu.Visible = False
        panelDesignSubMenu.Visible = False
        panelAddUpdateSubMenu.Visible = False
    End Sub

#End Region
#Region "version"
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        lblVersion.Text = String.Empty
        lblVersion.Text = lblVersion.Text + $" {GlobalEntity.Version}"

        BtnVersion.Text = lblVersion.Text

        Config.configObj = New Config(Config.configFilePath1)

        If lblFormTitle.Text = "Home" And count = 0 Then
            count = 1
        End If

        OpenChildForm(New ConfigurationForm)

        '19th Sep 2024
        If (Config.configObj.AutoSaveAuthor = True) Then
            Dim path = Config.configObj.EmployeeExcelPath
            If (CheckSEInstance() = False) Then
                'MessageBox.Show("SolidEdge is not running!", "Warning", MessageBoxButtons.OK)
                'Me.Close()
            Else
                _applicationEvents = CType(objApp.ApplicationEvents, SolidEdgeFramework.ISEApplicationEvents_Event)

                AddHandler _applicationEvents.BeforeDocumentSave, AddressOf _applicationEvents_BeforeDocumentSave

                Dim dtUsers As DataTable = ReadAuthorExcelData(dtUsers)
                dt = dtUsers
                If msg IsNot String.Empty And msg IsNot Nothing And Not msg = "" Then
                    ''NotifyIcon1.Icon = New System.Drawing.Icon()
                    'NotifyIcon1.Text = "My applicaiton"
                    'NotifyIcon1.Visible = True
                    'NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info

                    'NotifyIcon1.BalloonTipText = msg

                    'NotifyIcon1.BalloonTipTitle = "Message"


                    '----------------------------------------------------------------------
                    'NotifyIcon1.BalloonTipText = msg
                    'NotifyIcon1.BalloonTipTitle = "Error"
                    'NotifyIcon1.Icon = SystemIcons.Error


                    'NotifyIcon1.ShowBalloonTip(1000)


                    MessageBox.Show(msg, "Message")
                    msg = String.Empty

                End If
            End If

        End If

    End Sub
#End Region
#Region "SolidEdge Validations"
    Public Sub CloseChildForm()
        Dim mainObj As New MainClass
        Dim a As New AssemblyAutomationForm
        Dim pa As New PartAutomationForm
        Dim cp As New CopyPartForm
        Dim op As New OccurencePropertiesUpdateForm
        a.Closefn(mainObj)
        pa.Closefn(mainObj)
        cp.Closefn(mainObj)
        op.Closefn(mainObj)
        If mainObj.SolidEdgeinstance = "Close" Then
            Reset()
            HideSubMenu()
        End If
    End Sub
#End Region


#Region "Add/ Update"

    Public Overridable Sub BtnAddUpdate_Click(sender As Object, e As EventArgs) Handles BtnAddUpdate.Click


        ShowSubMenu(panelAddUpdateSubMenu)

        ActivateButton(sender, RGBColors.color6)


    End Sub

    Private Sub BtnAddUpdateVirtualStructure_Click(sender As Object, e As EventArgs) Handles BtnAddUpdateVirtualStructure.Click

        ActivateButton(sender, RGBColors.color6)

        'OpenChildForm(New VirtualAssemblyStructureForm2)
        OpenChildForm(New VirtualAssemblyStructureForm2)
    End Sub

    Private Sub BtnAddUpdateNewPartCreation_Click(sender As Object, e As EventArgs) Handles BtnAddUpdateNewPartCreation.Click
        ActivateButton(sender, RGBColors.color6)

        OpenChildForm(New CreateNewPartForm)
    End Sub

    Private Sub BtnAddUpdatePartSheetMetalUpdate_Click(sender As Object, e As EventArgs) Handles BtnAddUpdatePartSheetMetalUpdate.Click

        ActivateButton(sender, RGBColors.color6)

        OpenChildForm(New PartAutomationForm)
        'CloseChildForm()
    End Sub

    Private Sub BtnAddUpdateAssemblyValidation_Click(sender As Object, e As EventArgs) Handles BtnAddUpdateAssemblyValidation.Click
        ActivateButton(sender, RGBColors.color6)

        OpenChildForm(New AssemblyAutomationForm)
        'CloseChildForm()
    End Sub
    Private Sub BtnAddUpdateAuthorUpdation_Click(sender As Object, e As EventArgs)
        ActivateButton(sender, RGBColors.color6)

        OpenChildForm(New Author_Updation)
    End Sub

#End Region

#Region "Design"

    Public Overridable Sub BtnDesign_Click(sender As Object, e As EventArgs) Handles BtnDesign.Click

        ShowSubMenu(panelDesignSubMenu)
        ActivateButton(sender, RGBColors.color6)
        'OpenChildForm(New VirtualStructureForm)
    End Sub


    Private Sub BtnDesignGuideline_Click(sender As Object, e As EventArgs)
        'ShowSubMenu(panelDesignSubMenu)
        'ActivateButton(sender, RGBColors.color6)

        'OpenChildForm(New GuidelineForm)
    End Sub

    Private Sub BtnDesignCopyTransfer_Click(sender As Object, e As EventArgs) Handles BtnDesignCopyTransfer.Click
        'ShowSubMenu(panelDesignSubMenu)
        ActivateButton(sender, RGBColors.color6)

        OpenChildForm(New CopyPartForm)
        'CloseChildForm()
    End Sub

    Private Sub BtnDesignOccurenceProperties_Click(sender As Object, e As EventArgs) Handles BtnDesignOccurenceProperties.Click
        'ShowSubMenu(panelDesignSubMenu)
        ActivateButton(sender, RGBColors.color6)

        OpenChildForm(New OccurencePropertiesUpdateForm)
        'CloseChildForm()
    End Sub

#End Region

#Region "QC"

    Public Overridable Sub BtnQC_Click(sender As Object, e As EventArgs) Handles BtnQC.Click

        ShowSubMenu(panelQCSubMenu)
        ActivateButton(sender, RGBColors.color6)
    End Sub

    Private Sub BtnQCInterference_Click(sender As Object, e As EventArgs) Handles btnQCInterference.Click
        'ShowSubMenu(panelQCSubMenu)
        ActivateButton(sender, RGBColors.color6)

        OpenChildForm(New CheckInterferenceForm2)
    End Sub

    Private Sub BtnQCMTC_Click(sender As Object, e As EventArgs) Handles btnQCMTC.Click
        'ShowSubMenu(panelQCSubMenu)
        ActivateButton(sender, RGBColors.color6)

        OpenChildForm(New MTC_MTR_ReviewForm2)
    End Sub
    Private Sub BtnQCKPI_Click(sender As Object, e As EventArgs) Handles BtnQCKPI.Click
        ActivateButton(sender, RGBColors.color6)

        OpenChildForm(New KPI_ReportForm)
    End Sub
    Private Sub BtnQCRawMaterialEstimation_Click(sender As Object, e As EventArgs) Handles BtnQCRawMaterialEstimation.Click
        ActivateButton(sender, RGBColors.color6)

        OpenChildForm(New AssemblyBomForm)
    End Sub

    ''2nd Sep 2024
    ''Routing Sequence Tool:  remove this tool entirely
    'Private Sub BtnQCRoutingSequence_Click(sender As Object, e As EventArgs) Handles BtnQCRoutingSequence.Click
    '    ActivateButton(sender, RGBColors.color6)

    '    OpenChildForm(New RST_Design1)
    '    'ActiveForm.Size = New Size(1500, 800)
    'End Sub

#End Region

#Region "BtnHelp_Click"
    Private Sub BtnHelp_Click(sender As Object, e As EventArgs) Handles BtnHelp.Click

        Dim applicationExePath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)
        Dim helpDocDirPath As String = IO.Path.Combine(applicationExePath, "Help")

        btnQCMTC.Text = "MTC"
        BtnAddUpdatePartSheetMetalUpdate.Text = "Part Sheet-Metal Update"

        Dim name As ArrayList
        name = NameOfForms()

        'virtual structure
        For i = 0 To name.Count - 1
            If currentBtn.Text = name(i) Then

                Dim helpDocName As String = $"{ name.Item(i)} Help.pdf"
                Dim helpDocPath As String = IO.Path.Combine(helpDocDirPath, helpDocName)
                helpDocPath = helpDocPath.Replace("file:\", "")
                If IO.File.Exists(helpDocPath) Then
                    Process.Start(helpDocPath)
                    Exit For
                End If
            End If
        Next

        BtnAddUpdatePartSheetMetalUpdate.Text = "Part/ Sheet-Metal Update"
    End Sub
    Public Function NameOfForms()
        Dim FormName As New ArrayList From {
            "Virtual Structure",
            "New Part Creation",
            "Part Sheet-Metal Update",
            "Assembly Validation",
            "Copy && Transfer Part",
            "Occurence Properties",
            "Interference",
            "MTC",
            "KPI",
            "Raw Material Estimation",
            "Routing Sequence",
            "Configuration"
        }
        Return FormName
    End Function

#End Region


#Region "Configuration"

    Private Sub BtnConfiguration_Click(sender As Object, e As EventArgs) Handles BtnConfiguration.Click
        ActivateButton(sender, RGBColors.color6)
        OpenChildForm(New ConfigurationForm)
    End Sub
#End Region
#Region "LOGO"

    Public Overridable Sub BtnHome_Click(sender As Object, e As EventArgs) Handles BtnHome.Click
        HideSubMenu()
        Reset()
    End Sub

    Private Sub BtnHome_MouseEnter(sender As Object, e As EventArgs) Handles BtnHome.MouseEnter
        BtnHome.Cursor = System.Windows.Forms.Cursors.Hand
    End Sub

    Private Sub BtnHome_MouseLeave(sender As Object, e As EventArgs) Handles BtnHome.MouseLeave
        ' BtnHome.Cursor = System.Windows.Forms.Cursors.Hand
    End Sub

#End Region

#Region "Custom title bar selection and we can move the form"

    'Drag Form'
    <DllImport("user32.DLL", EntryPoint:="ReleaseCapture")>
    Private Shared Sub ReleaseCapture()
    End Sub

    <DllImport("user32.DLL", EntryPoint:="SendMessage")>
    Private Shared Sub SendMessage(ByVal hWnd As System.IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer)
    End Sub

    Private Sub PanelTitlBar_MouseDown(sender As Object, e As MouseEventArgs) Handles panelTitlBar.MouseDown
        ReleaseCapture()
        SendMessage(Me.Handle, &H112&, &HF012&, 0)
    End Sub

#End Region

#Region "Titlebar Exit Button Events"

    'Minimize
    '========================================

    Private Sub BtnMinimize_Click(sender As Object, e As EventArgs) Handles BtnMinimize.Click
        WindowState = FormWindowState.Minimized
    End Sub

    Private Sub BtnMinimize_MouseEnter(sender As Object, e As EventArgs) Handles BtnMinimize.MouseEnter
        BtnMinimize.BackColor = Color.FromArgb(42, 41, 93)
        BtnMinimize.IconColor = Color.White
    End Sub

    Private Sub BtnMinimize_MouseLeave(sender As Object, e As EventArgs) Handles BtnMinimize.MouseLeave
        BtnMinimize.BackColor = Color.FromArgb(31, 30, 68)
        BtnMinimize.IconColor = Color.Gainsboro
    End Sub

    'Maximize
    '========================================
    Private Sub BtnMaximize_Click(sender As Object, e As EventArgs) Handles BtnMaximize.Click
        If WindowState = FormWindowState.Normal Then
            WindowState = FormWindowState.Maximized
        Else
            WindowState = FormWindowState.Normal
        End If
    End Sub

    Private Sub BtnMaximize_MouseEnter(sender As Object, e As EventArgs) Handles BtnMaximize.MouseEnter
        BtnMaximize.BackColor = Color.FromArgb(42, 41, 93)
        BtnMaximize.IconColor = Color.White
    End Sub

    Private Sub BtnMaximize_MouseLeave(sender As Object, e As EventArgs) Handles BtnMaximize.MouseLeave
        BtnMaximize.BackColor = Color.FromArgb(31, 30, 68)
        BtnMaximize.IconColor = Color.Gainsboro
    End Sub

    'Exit
    '========================================

    Private Sub BtnExit_Click(sender As Object, e As EventArgs) Handles BtnExit.Click
        Application.Exit()
    End Sub

    Public Sub BtnExit_MouseEnter(sender As Object, e As EventArgs) Handles BtnExit.MouseEnter

        BtnExit.BackColor = Color.Red
        BtnExit.IconColor = Color.FromArgb(31, 30, 68)

    End Sub

    Private Sub BtnExit_MouseLeave(sender As Object, e As EventArgs) Handles BtnExit.MouseLeave
        BtnExit.BackColor = Color.FromArgb(31, 30, 68)
        BtnExit.IconColor = Color.Red
    End Sub

#End Region

#Region "Footer Buttons Events"

    Private Sub BtnClose_Click(sender As Object, e As EventArgs)
        HideSubMenu()
        Reset()
        OpenOnlyForm()
    End Sub

#End Region

#Region "Overrides Methods"

    'Note: We have commented following code as it cause the flickering in child form
    'To stop control flickering, we have added bouble buffer =true in the custom control that we have created.

    'Instead we added this code in child form (it make the form process slow) DO NOT USE IN CHILD FORM

    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim handleParam As CreateParams = MyBase.CreateParams
            handleParam.ExStyle = handleParam.ExStyle Or &H2000000
            handleParam.ClassStyle = CS_DropShadow
            Return handleParam
        End Get
    End Property

#Region "SolidEdge AutoSave Feature"
    Dim objApp As SolidEdgeFramework.Application
    Dim _applicationEvents As SolidEdgeFramework.ISEApplicationEvents_Event
    Dim Objdocuemnts As SolidEdgeFramework.Documents = Nothing
    Dim ObjdraftDoc As SolidEdgeDraft.DraftDocument = Nothing
    Dim ObjAssemblyDoc As SolidEdgeAssembly.AssemblyDocument = Nothing
    Dim ObjPartdoc As SolidEdgePart.PartDocument = Nothing
    Dim ObjSheetMetalDoc As SolidEdgePart.SheetMetalDocument = Nothing
    Dim Objsheet As SolidEdgeDraft.Sheet = Nothing
    Dim ObjdrwViews As SolidEdgeDraft.DrawingViews = Nothing
    Dim ObjdrwView As SolidEdgeDraft.DrawingView = Nothing
    Dim ObjpartLists As SolidEdgeDraft.PartsLists = Nothing
    Dim ObjpartList As SolidEdgeDraft.PartsList = Nothing
    Dim objModelLink As SolidEdgeDraft.ModelLink = Nothing
    Dim objModelLinks As SolidEdgeDraft.ModelLinks = Nothing
    Dim ObjTableCell As SolidEdgeDraft.TableCell = Nothing
    Dim ObjCols As SolidEdgeDraft.TableColumns = Nothing
    Dim ObjRows As SolidEdgeDraft.TableRows = Nothing
    Dim objDoc As SolidEdgeAssembly.AssemblyDocument
    Dim objParts As SolidEdgeAssembly.Occurrences

    Dim PartdocPath As String = String.Empty
    Dim SheetMetalDocPath As String = String.Empty
    Dim AssemblyDocPath As String = String.Empty
    Dim DocumentPath1 As String = Nothing
    Dim activeDocPath As String
    Dim MainAssemblyPath As String
    Dim dt As DataTable = New DataTable()
    Dim msg As String = String.Empty


    '19th Sep 2024
    'Private Sub MainForm_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
    '    If (Config.configObj.AutoSaveAuthor = True) Then
    '        Dim path = Config.configObj.EmployeeExcelPath
    '        If (CheckSEInstance() = False) Then

    '        Else
    '            _applicationEvents = CType(objApp.ApplicationEvents, SolidEdgeFramework.ISEApplicationEvents_Event)

    '            AddHandler _applicationEvents.BeforeDocumentSave, AddressOf _applicationEvents_BeforeDocumentSave

    '            Dim dtUsers As DataTable = ReadAuthorExcelData(dtUsers)
    '            dt = dtUsers
    '            If msg IsNot String.Empty And msg IsNot Nothing And Not msg = "" Then
    '                ''NotifyIcon1.Icon = New System.Drawing.Icon()
    '                'NotifyIcon1.Text = "My applicaiton"
    '                'NotifyIcon1.Visible = True
    '                'NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info

    '                'NotifyIcon1.BalloonTipText = msg

    '                'NotifyIcon1.BalloonTipTitle = "Message"


    '                '----------------------------------------------------------------------
    '                'NotifyIcon1.BalloonTipText = msg
    '                'NotifyIcon1.BalloonTipTitle = "Error"
    '                'NotifyIcon1.Icon = SystemIcons.Error


    '                'NotifyIcon1.ShowBalloonTip(1000)


    '                MessageBox.Show(msg, "Message")
    '                msg = String.Empty

    '            End If
    '        End If

    '    End If

    'End Sub
    Public Function CheckSEInstance() As Boolean

        OleMessageFilter.Register()
        Try
            objApp = Marshal.GetActiveObject("SolidEdge.Application")
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function ReadAuthorExcelData(ByVal dt As DataTable) As DataTable

        Dim Path = Config.configObj.EmployeeExcelPath

        If Path = "" Then
            MessageBox.Show("Employee Excel Path not Exist in Config, Please set Employee Excel Path in Config.")
            Exit Function
        End If

        If (System.IO.File.Exists(Path)) Then
            Dim ds As New DataSet("Data")

            Using stream As FileStream = File.Open(Path, FileMode.Open, FileAccess.Read)
                Using reader As ExcelDataReader.IExcelDataReader = ExcelReaderFactory.CreateOpenXmlReader(stream)
                    Dim conf = New ExcelDataSetConfiguration With
                        {
                            .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration With
                            {
                                .UseHeaderRow = True
                            }
                        }
                    ds = reader.AsDataSet(conf)


                End Using

            End Using
            dt = ds.Tables(0)

        Else
            MessageBox.Show("Employee Excel File Missing at Location.")  'Employee Excel Path Missing or not Exist in Config
            Exit Function
        End If

        Return dt
    End Function

    Private Sub _applicationEvents_BeforeDocumentSave(ByVal theDocument As Object)
        Try
            msg = UpdateSummaryInfoAuthor(dt)


        Catch ex As Exception

        End Try
    End Sub

    Public Function UpdateSummaryInfoAuthor(ByRef dtUsers As DataTable) As String
        'Dim objDocument As SolidEdgeFramework.SolidEdgeDocument = objApp.ActiveDocument
        'Dim objParts As SolidEdgeAssembly.Occurrences

        'Dim propSets As SolidEdgeFramework.PropertySets = objApp.ActiveDocument.Properties

        'Dim custProps As SolidEdgeFramework.Properties = propSets.Item("SummaryInformation")

        ' Getting the parts objects of the AssemblyDocument object.
        'objParts = objDocument.Occurrences
        'Dim a As Integer

        'For a = 0 To dtUsers.Rows.Count - 1

        Try
            Dim objApp As SolidEdgeFramework.Application = System.Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application")
            Dim objPropSets As SolidEdgeFramework.PropertySets = Nothing
            Dim objProp As SolidEdgeFramework.Property = Nothing
            Dim objProps As SolidEdgeFramework.Properties = Nothing
            Dim objDocument As SolidEdgeFramework.SolidEdgeDocument = Nothing
            objDocument = objApp.ActiveDocument

            objPropSets = objDocument.Properties

            Dim finished As Integer = 0
            For Each objProps In objPropSets
                If objProps.Name = "SummaryInformation" Then
                    For Each objProp In objProps
                        If objProp.Name = "Author" Then
                            Dim dv As New DataView(dtUsers)
                            Dim WindowsAuthorName As String = "Windows Usernames"
                            Dim FullNameTitle As String = "FULL NAME"
                            Dim BECStdFormateTitle As String = "BEC Standard Format"
                            Dim value1 As String = objProp.Value.ToString()
                            'Dim filter As String = $"Convert([{WindowsAuthorName}], 'System.String') = '{value1}'"
                            Dim filter As String = $"Convert([{WindowsAuthorName}], 'System.String') = '{value1}' OR Convert([{FullNameTitle}], 'System.String') = '{value1}'OR Convert([{BECStdFormateTitle}], 'System.String') = '{value1}'"
                            dv.RowFilter = filter '$"{materialUsedCol}='{cmbMaterialUsed2_Mw.Text}'"
                            If dv.Count = 0 Then
                                Return "Author Name Not Exist"

                            Else
                                For Each drv As DataRowView In dv
                                    'mtcMtrModelObj.authorList.Remove(OldValue)
                                    Dim Author = If(drv("BEC Standard Format") = Nothing, "", drv("BEC Standard Format"))

                                    'If (Author.ToString.Trim().Contains(" ")) Then
                                    objProp.Value = Author.ToString.ToUpper()
                                        objProps.Save()
                                        Exit Function
                                    'Else
                                    '    Return "Please Follow BEC Standard Formate to Save Author Field"
                                    'End If

                                    'objDocument.Save()
                                    'MessageBox.Show("AUTHOR NAME UPDATED : " + objProp.Value, "Message")


                                Next
                            End If


                        End If
                    Next
                End If
            Next

            'objDocument.Close()

            'Next
        Catch ex As Exception

        End Try

    End Function
    Private Function GetPropValue(ByVal prop1 As [Property])
        Dim value As String = String.Empty
        If prop1.Value IsNot Nothing Then
            value = prop1.Value.ToString().Trim()
        End If
        Return value
    End Function



#End Region




#End Region

End Class