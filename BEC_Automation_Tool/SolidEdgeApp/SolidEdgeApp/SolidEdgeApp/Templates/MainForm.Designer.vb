<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.panelTitlBar = New System.Windows.Forms.Panel()
        Me.LblHelp = New System.Windows.Forms.Label()
        Me.BtnHelp = New FontAwesome.Sharp.IconButton()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.BtnMinimize = New FontAwesome.Sharp.IconButton()
        Me.BtnMaximize = New FontAwesome.Sharp.IconButton()
        Me.BtnExit = New FontAwesome.Sharp.IconButton()
        Me.lblFormTitle = New System.Windows.Forms.Label()
        Me.IconCurrentForm = New FontAwesome.Sharp.IconPictureBox()
        Me.panelDesktop = New System.Windows.Forms.Panel()
        Me.pnlFooterShadow = New System.Windows.Forms.Panel()
        Me.SideMenuPanel = New MoCustomControls.MoPanel()
        Me.BtnVersion = New FontAwesome.Sharp.IconButton()
        Me.BtnConfiguration = New FontAwesome.Sharp.IconButton()
        Me.panelQCSubMenu = New MoCustomControls.MoPanel()
        Me.BtnQCRawMaterialEstimation = New FontAwesome.Sharp.IconButton()
        Me.BtnQCKPI = New FontAwesome.Sharp.IconButton()
        Me.btnQCMTC = New FontAwesome.Sharp.IconButton()
        Me.btnQCInterference = New FontAwesome.Sharp.IconButton()
        Me.BtnQC = New FontAwesome.Sharp.IconButton()
        Me.panelDesignSubMenu = New MoCustomControls.MoPanel()
        Me.BtnDesignOccurenceProperties = New FontAwesome.Sharp.IconButton()
        Me.BtnDesignCopyTransfer = New FontAwesome.Sharp.IconButton()
        Me.BtnDesign = New FontAwesome.Sharp.IconButton()
        Me.panelAddUpdateSubMenu = New MoCustomControls.MoPanel()
        Me.BtnAddUpdateAssemblyValidation = New FontAwesome.Sharp.IconButton()
        Me.BtnAddUpdatePartSheetMetalUpdate = New FontAwesome.Sharp.IconButton()
        Me.BtnAddUpdateNewPartCreation = New FontAwesome.Sharp.IconButton()
        Me.BtnAddUpdateVirtualStructure = New FontAwesome.Sharp.IconButton()
        Me.BtnAddUpdate = New FontAwesome.Sharp.IconButton()
        Me.logoPanel = New System.Windows.Forms.Panel()
        Me.BtnHome = New System.Windows.Forms.PictureBox()
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.BeC_Automation_Installer1 = New SolidEdgeApp.BEC_Automation_Installer()
        Me.panelTitlBar.SuspendLayout()
        Me.FlowLayoutPanel1.SuspendLayout()
        CType(Me.IconCurrentForm, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SideMenuPanel.SuspendLayout()
        Me.panelQCSubMenu.SuspendLayout()
        Me.panelDesignSubMenu.SuspendLayout()
        Me.panelAddUpdateSubMenu.SuspendLayout()
        Me.logoPanel.SuspendLayout()
        CType(Me.BtnHome, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'panelTitlBar
        '
        Me.panelTitlBar.BackColor = System.Drawing.Color.FromArgb(CType(CType(31, Byte), Integer), CType(CType(30, Byte), Integer), CType(CType(68, Byte), Integer))
        Me.panelTitlBar.Controls.Add(Me.LblHelp)
        Me.panelTitlBar.Controls.Add(Me.BtnHelp)
        Me.panelTitlBar.Controls.Add(Me.FlowLayoutPanel1)
        Me.panelTitlBar.Controls.Add(Me.BtnMinimize)
        Me.panelTitlBar.Controls.Add(Me.BtnMaximize)
        Me.panelTitlBar.Controls.Add(Me.BtnExit)
        Me.panelTitlBar.Controls.Add(Me.lblFormTitle)
        Me.panelTitlBar.Controls.Add(Me.IconCurrentForm)
        Me.panelTitlBar.Dock = System.Windows.Forms.DockStyle.Top
        Me.panelTitlBar.Location = New System.Drawing.Point(230, 0)
        Me.panelTitlBar.Name = "panelTitlBar"
        Me.panelTitlBar.Size = New System.Drawing.Size(1054, 75)
        Me.panelTitlBar.TabIndex = 11
        '
        'LblHelp
        '
        Me.LblHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblHelp.AutoSize = True
        Me.LblHelp.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblHelp.ForeColor = System.Drawing.Color.Gainsboro
        Me.LblHelp.Location = New System.Drawing.Point(1015, 59)
        Me.LblHelp.Name = "LblHelp"
        Me.LblHelp.Size = New System.Drawing.Size(31, 13)
        Me.LblHelp.TabIndex = 8
        Me.LblHelp.Text = "Help"
        Me.LblHelp.Visible = False
        '
        'BtnHelp
        '
        Me.BtnHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnHelp.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnHelp.ForeColor = System.Drawing.Color.FromArgb(CType(CType(31, Byte), Integer), CType(CType(30, Byte), Integer), CType(CType(68, Byte), Integer))
        Me.BtnHelp.IconChar = FontAwesome.Sharp.IconChar.CircleQuestion
        Me.BtnHelp.IconColor = System.Drawing.Color.Gainsboro
        Me.BtnHelp.IconFont = FontAwesome.Sharp.IconFont.[Auto]
        Me.BtnHelp.IconSize = 25
        Me.BtnHelp.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnHelp.Location = New System.Drawing.Point(1014, 31)
        Me.BtnHelp.Name = "BtnHelp"
        Me.BtnHelp.Size = New System.Drawing.Size(32, 27)
        Me.BtnHelp.TabIndex = 7
        Me.BtnHelp.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnHelp.UseVisualStyleBackColor = True
        Me.BtnHelp.Visible = False
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FlowLayoutPanel1.Controls.Add(Me.lblVersion)
        Me.FlowLayoutPanel1.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft
        Me.FlowLayoutPanel1.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(823, 47)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(200, 22)
        Me.FlowLayoutPanel1.TabIndex = 6
        '
        'lblVersion
        '
        Me.lblVersion.AutoSize = True
        Me.lblVersion.Font = New System.Drawing.Font("Segoe UI Light", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVersion.ForeColor = System.Drawing.Color.Gainsboro
        Me.lblVersion.Location = New System.Drawing.Point(173, 0)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(24, 20)
        Me.lblVersion.TabIndex = 5
        Me.lblVersion.Text = "V-"
        Me.lblVersion.Visible = False
        '
        'BtnMinimize
        '
        Me.BtnMinimize.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnMinimize.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnMinimize.IconChar = FontAwesome.Sharp.IconChar.WindowMinimize
        Me.BtnMinimize.IconColor = System.Drawing.Color.Gainsboro
        Me.BtnMinimize.IconFont = FontAwesome.Sharp.IconFont.[Auto]
        Me.BtnMinimize.IconSize = 18
        Me.BtnMinimize.Location = New System.Drawing.Point(973, 5)
        Me.BtnMinimize.Name = "BtnMinimize"
        Me.BtnMinimize.Size = New System.Drawing.Size(24, 24)
        Me.BtnMinimize.TabIndex = 4
        Me.BtnMinimize.Text = " "
        Me.BtnMinimize.UseVisualStyleBackColor = True
        '
        'BtnMaximize
        '
        Me.BtnMaximize.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnMaximize.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnMaximize.IconChar = FontAwesome.Sharp.IconChar.WindowMaximize
        Me.BtnMaximize.IconColor = System.Drawing.Color.Gainsboro
        Me.BtnMaximize.IconFont = FontAwesome.Sharp.IconFont.[Auto]
        Me.BtnMaximize.IconSize = 18
        Me.BtnMaximize.Location = New System.Drawing.Point(999, 5)
        Me.BtnMaximize.Name = "BtnMaximize"
        Me.BtnMaximize.Size = New System.Drawing.Size(24, 24)
        Me.BtnMaximize.TabIndex = 3
        Me.BtnMaximize.Text = " "
        Me.BtnMaximize.UseVisualStyleBackColor = True
        '
        'BtnExit
        '
        Me.BtnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnExit.IconChar = FontAwesome.Sharp.IconChar.PowerOff
        Me.BtnExit.IconColor = System.Drawing.Color.Red
        Me.BtnExit.IconFont = FontAwesome.Sharp.IconFont.[Auto]
        Me.BtnExit.IconSize = 18
        Me.BtnExit.Location = New System.Drawing.Point(1025, 5)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(24, 24)
        Me.BtnExit.TabIndex = 2
        Me.BtnExit.Text = " "
        Me.BtnExit.UseVisualStyleBackColor = True
        '
        'lblFormTitle
        '
        Me.lblFormTitle.AutoSize = True
        Me.lblFormTitle.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFormTitle.ForeColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(161, Byte), Integer), CType(CType(251, Byte), Integer))
        Me.lblFormTitle.Location = New System.Drawing.Point(49, 36)
        Me.lblFormTitle.Name = "lblFormTitle"
        Me.lblFormTitle.Size = New System.Drawing.Size(51, 20)
        Me.lblFormTitle.TabIndex = 1
        Me.lblFormTitle.Text = "Home"
        '
        'IconCurrentForm
        '
        Me.IconCurrentForm.BackColor = System.Drawing.Color.FromArgb(CType(CType(31, Byte), Integer), CType(CType(30, Byte), Integer), CType(CType(68, Byte), Integer))
        Me.IconCurrentForm.ForeColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(161, Byte), Integer), CType(CType(251, Byte), Integer))
        Me.IconCurrentForm.IconChar = FontAwesome.Sharp.IconChar.Home
        Me.IconCurrentForm.IconColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(161, Byte), Integer), CType(CType(251, Byte), Integer))
        Me.IconCurrentForm.IconFont = FontAwesome.Sharp.IconFont.[Auto]
        Me.IconCurrentForm.Location = New System.Drawing.Point(11, 24)
        Me.IconCurrentForm.Name = "IconCurrentForm"
        Me.IconCurrentForm.Size = New System.Drawing.Size(32, 32)
        Me.IconCurrentForm.TabIndex = 0
        Me.IconCurrentForm.TabStop = False
        '
        'panelDesktop
        '
        Me.panelDesktop.Dock = System.Windows.Forms.DockStyle.Fill
        Me.panelDesktop.Location = New System.Drawing.Point(230, 75)
        Me.panelDesktop.Name = "panelDesktop"
        Me.panelDesktop.Size = New System.Drawing.Size(1054, 636)
        Me.panelDesktop.TabIndex = 12
        '
        'pnlFooterShadow
        '
        Me.pnlFooterShadow.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(113, Byte), Integer), CType(CType(227, Byte), Integer))
        Me.pnlFooterShadow.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlFooterShadow.Location = New System.Drawing.Point(230, 710)
        Me.pnlFooterShadow.Name = "pnlFooterShadow"
        Me.pnlFooterShadow.Size = New System.Drawing.Size(1054, 1)
        Me.pnlFooterShadow.TabIndex = 29
        '
        'SideMenuPanel
        '
        Me.SideMenuPanel.AutoScroll = True
        Me.SideMenuPanel.BackColor = System.Drawing.Color.FromArgb(CType(CType(31, Byte), Integer), CType(CType(30, Byte), Integer), CType(CType(68, Byte), Integer))
        Me.SideMenuPanel.Controls.Add(Me.BtnVersion)
        Me.SideMenuPanel.Controls.Add(Me.BtnConfiguration)
        Me.SideMenuPanel.Controls.Add(Me.panelQCSubMenu)
        Me.SideMenuPanel.Controls.Add(Me.BtnQC)
        Me.SideMenuPanel.Controls.Add(Me.panelDesignSubMenu)
        Me.SideMenuPanel.Controls.Add(Me.BtnDesign)
        Me.SideMenuPanel.Controls.Add(Me.panelAddUpdateSubMenu)
        Me.SideMenuPanel.Controls.Add(Me.BtnAddUpdate)
        Me.SideMenuPanel.Controls.Add(Me.logoPanel)
        Me.SideMenuPanel.Dock = System.Windows.Forms.DockStyle.Left
        Me.SideMenuPanel.Location = New System.Drawing.Point(0, 0)
        Me.SideMenuPanel.Name = "SideMenuPanel"
        Me.SideMenuPanel.Size = New System.Drawing.Size(230, 711)
        Me.SideMenuPanel.TabIndex = 10
        '
        'BtnVersion
        '
        Me.BtnVersion.BackColor = System.Drawing.Color.Transparent
        Me.BtnVersion.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.BtnVersion.FlatAppearance.BorderSize = 0
        Me.BtnVersion.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.BtnVersion.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.BtnVersion.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnVersion.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnVersion.ForeColor = System.Drawing.Color.Gainsboro
        Me.BtnVersion.IconChar = FontAwesome.Sharp.IconChar.None
        Me.BtnVersion.IconColor = System.Drawing.Color.Gainsboro
        Me.BtnVersion.IconFont = FontAwesome.Sharp.IconFont.[Auto]
        Me.BtnVersion.IconSize = 32
        Me.BtnVersion.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnVersion.Location = New System.Drawing.Point(0, 770)
        Me.BtnVersion.Name = "BtnVersion"
        Me.BtnVersion.Padding = New System.Windows.Forms.Padding(10, 0, 10, 0)
        Me.BtnVersion.Size = New System.Drawing.Size(213, 38)
        Me.BtnVersion.TabIndex = 15
        Me.BtnVersion.Text = "Version"
        Me.BtnVersion.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.BtnVersion.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.BtnVersion.UseVisualStyleBackColor = False
        '
        'BtnConfiguration
        '
        Me.BtnConfiguration.BackColor = System.Drawing.Color.Transparent
        Me.BtnConfiguration.Dock = System.Windows.Forms.DockStyle.Top
        Me.BtnConfiguration.FlatAppearance.BorderSize = 0
        Me.BtnConfiguration.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.BtnConfiguration.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.BtnConfiguration.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnConfiguration.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnConfiguration.ForeColor = System.Drawing.Color.Gainsboro
        Me.BtnConfiguration.IconChar = FontAwesome.Sharp.IconChar.Wrench
        Me.BtnConfiguration.IconColor = System.Drawing.Color.Gainsboro
        Me.BtnConfiguration.IconFont = FontAwesome.Sharp.IconFont.[Auto]
        Me.BtnConfiguration.IconSize = 32
        Me.BtnConfiguration.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnConfiguration.Location = New System.Drawing.Point(0, 710)
        Me.BtnConfiguration.Name = "BtnConfiguration"
        Me.BtnConfiguration.Padding = New System.Windows.Forms.Padding(10, 0, 10, 0)
        Me.BtnConfiguration.Size = New System.Drawing.Size(213, 60)
        Me.BtnConfiguration.TabIndex = 14
        Me.BtnConfiguration.Text = "Configuration"
        Me.BtnConfiguration.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnConfiguration.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.BtnConfiguration.UseVisualStyleBackColor = False
        '
        'panelQCSubMenu
        '
        Me.panelQCSubMenu.Controls.Add(Me.BtnQCRawMaterialEstimation)
        Me.panelQCSubMenu.Controls.Add(Me.BtnQCKPI)
        Me.panelQCSubMenu.Controls.Add(Me.btnQCMTC)
        Me.panelQCSubMenu.Controls.Add(Me.btnQCInterference)
        Me.panelQCSubMenu.Dock = System.Windows.Forms.DockStyle.Top
        Me.panelQCSubMenu.Location = New System.Drawing.Point(0, 525)
        Me.panelQCSubMenu.Name = "panelQCSubMenu"
        Me.panelQCSubMenu.Size = New System.Drawing.Size(213, 185)
        Me.panelQCSubMenu.TabIndex = 11
        '
        'BtnQCRawMaterialEstimation
        '
        Me.BtnQCRawMaterialEstimation.BackColor = System.Drawing.Color.Transparent
        Me.BtnQCRawMaterialEstimation.Dock = System.Windows.Forms.DockStyle.Top
        Me.BtnQCRawMaterialEstimation.FlatAppearance.BorderSize = 0
        Me.BtnQCRawMaterialEstimation.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.BtnQCRawMaterialEstimation.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.BtnQCRawMaterialEstimation.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnQCRawMaterialEstimation.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnQCRawMaterialEstimation.ForeColor = System.Drawing.Color.Gainsboro
        Me.BtnQCRawMaterialEstimation.IconChar = FontAwesome.Sharp.IconChar.None
        Me.BtnQCRawMaterialEstimation.IconColor = System.Drawing.Color.Gainsboro
        Me.BtnQCRawMaterialEstimation.IconFont = FontAwesome.Sharp.IconFont.[Auto]
        Me.BtnQCRawMaterialEstimation.IconSize = 32
        Me.BtnQCRawMaterialEstimation.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnQCRawMaterialEstimation.Location = New System.Drawing.Point(0, 135)
        Me.BtnQCRawMaterialEstimation.Name = "BtnQCRawMaterialEstimation"
        Me.BtnQCRawMaterialEstimation.Padding = New System.Windows.Forms.Padding(30, 0, 10, 0)
        Me.BtnQCRawMaterialEstimation.Size = New System.Drawing.Size(213, 45)
        Me.BtnQCRawMaterialEstimation.TabIndex = 11
        Me.BtnQCRawMaterialEstimation.Text = "Raw Material Estimation"
        Me.BtnQCRawMaterialEstimation.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnQCRawMaterialEstimation.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.BtnQCRawMaterialEstimation.UseVisualStyleBackColor = False
        '
        'BtnQCKPI
        '
        Me.BtnQCKPI.BackColor = System.Drawing.Color.Transparent
        Me.BtnQCKPI.Dock = System.Windows.Forms.DockStyle.Top
        Me.BtnQCKPI.FlatAppearance.BorderSize = 0
        Me.BtnQCKPI.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.BtnQCKPI.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.BtnQCKPI.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnQCKPI.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnQCKPI.ForeColor = System.Drawing.Color.Gainsboro
        Me.BtnQCKPI.IconChar = FontAwesome.Sharp.IconChar.None
        Me.BtnQCKPI.IconColor = System.Drawing.Color.Gainsboro
        Me.BtnQCKPI.IconFont = FontAwesome.Sharp.IconFont.[Auto]
        Me.BtnQCKPI.IconSize = 32
        Me.BtnQCKPI.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnQCKPI.Location = New System.Drawing.Point(0, 90)
        Me.BtnQCKPI.Name = "BtnQCKPI"
        Me.BtnQCKPI.Padding = New System.Windows.Forms.Padding(30, 0, 10, 0)
        Me.BtnQCKPI.Size = New System.Drawing.Size(213, 45)
        Me.BtnQCKPI.TabIndex = 9
        Me.BtnQCKPI.Text = "KPI"
        Me.BtnQCKPI.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnQCKPI.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.BtnQCKPI.UseVisualStyleBackColor = False
        '
        'btnQCMTC
        '
        Me.btnQCMTC.BackColor = System.Drawing.Color.Transparent
        Me.btnQCMTC.Dock = System.Windows.Forms.DockStyle.Top
        Me.btnQCMTC.FlatAppearance.BorderSize = 0
        Me.btnQCMTC.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.btnQCMTC.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.btnQCMTC.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnQCMTC.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnQCMTC.ForeColor = System.Drawing.Color.Gainsboro
        Me.btnQCMTC.IconChar = FontAwesome.Sharp.IconChar.None
        Me.btnQCMTC.IconColor = System.Drawing.Color.Gainsboro
        Me.btnQCMTC.IconFont = FontAwesome.Sharp.IconFont.[Auto]
        Me.btnQCMTC.IconSize = 32
        Me.btnQCMTC.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnQCMTC.Location = New System.Drawing.Point(0, 45)
        Me.btnQCMTC.Name = "btnQCMTC"
        Me.btnQCMTC.Padding = New System.Windows.Forms.Padding(30, 0, 10, 0)
        Me.btnQCMTC.Size = New System.Drawing.Size(213, 45)
        Me.btnQCMTC.TabIndex = 8
        Me.btnQCMTC.Text = "MTC"
        Me.btnQCMTC.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnQCMTC.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnQCMTC.UseVisualStyleBackColor = False
        '
        'btnQCInterference
        '
        Me.btnQCInterference.BackColor = System.Drawing.Color.Transparent
        Me.btnQCInterference.Dock = System.Windows.Forms.DockStyle.Top
        Me.btnQCInterference.FlatAppearance.BorderSize = 0
        Me.btnQCInterference.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.btnQCInterference.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.btnQCInterference.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnQCInterference.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnQCInterference.ForeColor = System.Drawing.Color.Gainsboro
        Me.btnQCInterference.IconChar = FontAwesome.Sharp.IconChar.None
        Me.btnQCInterference.IconColor = System.Drawing.Color.Gainsboro
        Me.btnQCInterference.IconFont = FontAwesome.Sharp.IconFont.[Auto]
        Me.btnQCInterference.IconSize = 32
        Me.btnQCInterference.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnQCInterference.Location = New System.Drawing.Point(0, 0)
        Me.btnQCInterference.Name = "btnQCInterference"
        Me.btnQCInterference.Padding = New System.Windows.Forms.Padding(30, 0, 10, 0)
        Me.btnQCInterference.Size = New System.Drawing.Size(213, 45)
        Me.btnQCInterference.TabIndex = 7
        Me.btnQCInterference.Text = "Interference"
        Me.btnQCInterference.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnQCInterference.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnQCInterference.UseVisualStyleBackColor = False
        '
        'BtnQC
        '
        Me.BtnQC.BackColor = System.Drawing.Color.Transparent
        Me.BtnQC.Dock = System.Windows.Forms.DockStyle.Top
        Me.BtnQC.FlatAppearance.BorderSize = 0
        Me.BtnQC.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.BtnQC.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.BtnQC.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnQC.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnQC.ForeColor = System.Drawing.Color.Gainsboro
        Me.BtnQC.IconChar = FontAwesome.Sharp.IconChar.CheckDouble
        Me.BtnQC.IconColor = System.Drawing.Color.Gainsboro
        Me.BtnQC.IconFont = FontAwesome.Sharp.IconFont.[Auto]
        Me.BtnQC.IconSize = 32
        Me.BtnQC.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnQC.Location = New System.Drawing.Point(0, 465)
        Me.BtnQC.Name = "BtnQC"
        Me.BtnQC.Padding = New System.Windows.Forms.Padding(10, 0, 10, 0)
        Me.BtnQC.Size = New System.Drawing.Size(213, 60)
        Me.BtnQC.TabIndex = 10
        Me.BtnQC.Text = "QC Report"
        Me.BtnQC.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnQC.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.BtnQC.UseVisualStyleBackColor = False
        '
        'panelDesignSubMenu
        '
        Me.panelDesignSubMenu.Controls.Add(Me.BtnDesignOccurenceProperties)
        Me.panelDesignSubMenu.Controls.Add(Me.BtnDesignCopyTransfer)
        Me.panelDesignSubMenu.Dock = System.Windows.Forms.DockStyle.Top
        Me.panelDesignSubMenu.Location = New System.Drawing.Point(0, 375)
        Me.panelDesignSubMenu.Name = "panelDesignSubMenu"
        Me.panelDesignSubMenu.Size = New System.Drawing.Size(213, 90)
        Me.panelDesignSubMenu.TabIndex = 9
        '
        'BtnDesignOccurenceProperties
        '
        Me.BtnDesignOccurenceProperties.BackColor = System.Drawing.Color.Transparent
        Me.BtnDesignOccurenceProperties.Dock = System.Windows.Forms.DockStyle.Top
        Me.BtnDesignOccurenceProperties.FlatAppearance.BorderSize = 0
        Me.BtnDesignOccurenceProperties.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.BtnDesignOccurenceProperties.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.BtnDesignOccurenceProperties.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnDesignOccurenceProperties.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnDesignOccurenceProperties.ForeColor = System.Drawing.Color.Gainsboro
        Me.BtnDesignOccurenceProperties.IconChar = FontAwesome.Sharp.IconChar.None
        Me.BtnDesignOccurenceProperties.IconColor = System.Drawing.Color.Gainsboro
        Me.BtnDesignOccurenceProperties.IconFont = FontAwesome.Sharp.IconFont.[Auto]
        Me.BtnDesignOccurenceProperties.IconSize = 32
        Me.BtnDesignOccurenceProperties.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnDesignOccurenceProperties.Location = New System.Drawing.Point(0, 45)
        Me.BtnDesignOccurenceProperties.Name = "BtnDesignOccurenceProperties"
        Me.BtnDesignOccurenceProperties.Padding = New System.Windows.Forms.Padding(30, 0, 10, 0)
        Me.BtnDesignOccurenceProperties.Size = New System.Drawing.Size(213, 45)
        Me.BtnDesignOccurenceProperties.TabIndex = 10
        Me.BtnDesignOccurenceProperties.Text = "Occurence Properties"
        Me.BtnDesignOccurenceProperties.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnDesignOccurenceProperties.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.BtnDesignOccurenceProperties.UseVisualStyleBackColor = False
        '
        'BtnDesignCopyTransfer
        '
        Me.BtnDesignCopyTransfer.BackColor = System.Drawing.Color.Transparent
        Me.BtnDesignCopyTransfer.Dock = System.Windows.Forms.DockStyle.Top
        Me.BtnDesignCopyTransfer.FlatAppearance.BorderSize = 0
        Me.BtnDesignCopyTransfer.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.BtnDesignCopyTransfer.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.BtnDesignCopyTransfer.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnDesignCopyTransfer.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnDesignCopyTransfer.ForeColor = System.Drawing.Color.Gainsboro
        Me.BtnDesignCopyTransfer.IconChar = FontAwesome.Sharp.IconChar.None
        Me.BtnDesignCopyTransfer.IconColor = System.Drawing.Color.Gainsboro
        Me.BtnDesignCopyTransfer.IconFont = FontAwesome.Sharp.IconFont.[Auto]
        Me.BtnDesignCopyTransfer.IconSize = 32
        Me.BtnDesignCopyTransfer.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnDesignCopyTransfer.Location = New System.Drawing.Point(0, 0)
        Me.BtnDesignCopyTransfer.Name = "BtnDesignCopyTransfer"
        Me.BtnDesignCopyTransfer.Padding = New System.Windows.Forms.Padding(30, 0, 10, 0)
        Me.BtnDesignCopyTransfer.Size = New System.Drawing.Size(213, 45)
        Me.BtnDesignCopyTransfer.TabIndex = 9
        Me.BtnDesignCopyTransfer.Text = "Copy && Transfer Part"
        Me.BtnDesignCopyTransfer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnDesignCopyTransfer.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.BtnDesignCopyTransfer.UseVisualStyleBackColor = False
        '
        'BtnDesign
        '
        Me.BtnDesign.BackColor = System.Drawing.Color.Transparent
        Me.BtnDesign.Dock = System.Windows.Forms.DockStyle.Top
        Me.BtnDesign.FlatAppearance.BorderSize = 0
        Me.BtnDesign.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.BtnDesign.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.BtnDesign.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnDesign.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnDesign.ForeColor = System.Drawing.Color.Gainsboro
        Me.BtnDesign.IconChar = FontAwesome.Sharp.IconChar.DraftingCompass
        Me.BtnDesign.IconColor = System.Drawing.Color.Gainsboro
        Me.BtnDesign.IconFont = FontAwesome.Sharp.IconFont.[Auto]
        Me.BtnDesign.IconSize = 32
        Me.BtnDesign.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnDesign.Location = New System.Drawing.Point(0, 315)
        Me.BtnDesign.Name = "BtnDesign"
        Me.BtnDesign.Padding = New System.Windows.Forms.Padding(10, 0, 10, 0)
        Me.BtnDesign.Size = New System.Drawing.Size(213, 60)
        Me.BtnDesign.TabIndex = 8
        Me.BtnDesign.Text = "Design"
        Me.BtnDesign.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnDesign.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.BtnDesign.UseVisualStyleBackColor = False
        '
        'panelAddUpdateSubMenu
        '
        Me.panelAddUpdateSubMenu.Controls.Add(Me.BtnAddUpdateAssemblyValidation)
        Me.panelAddUpdateSubMenu.Controls.Add(Me.BtnAddUpdatePartSheetMetalUpdate)
        Me.panelAddUpdateSubMenu.Controls.Add(Me.BtnAddUpdateNewPartCreation)
        Me.panelAddUpdateSubMenu.Controls.Add(Me.BtnAddUpdateVirtualStructure)
        Me.panelAddUpdateSubMenu.Dock = System.Windows.Forms.DockStyle.Top
        Me.panelAddUpdateSubMenu.Location = New System.Drawing.Point(0, 135)
        Me.panelAddUpdateSubMenu.Name = "panelAddUpdateSubMenu"
        Me.panelAddUpdateSubMenu.Size = New System.Drawing.Size(213, 180)
        Me.panelAddUpdateSubMenu.TabIndex = 7
        '
        'BtnAddUpdateAssemblyValidation
        '
        Me.BtnAddUpdateAssemblyValidation.BackColor = System.Drawing.Color.Transparent
        Me.BtnAddUpdateAssemblyValidation.Dock = System.Windows.Forms.DockStyle.Top
        Me.BtnAddUpdateAssemblyValidation.FlatAppearance.BorderSize = 0
        Me.BtnAddUpdateAssemblyValidation.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.BtnAddUpdateAssemblyValidation.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.BtnAddUpdateAssemblyValidation.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnAddUpdateAssemblyValidation.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAddUpdateAssemblyValidation.ForeColor = System.Drawing.Color.Gainsboro
        Me.BtnAddUpdateAssemblyValidation.IconChar = FontAwesome.Sharp.IconChar.None
        Me.BtnAddUpdateAssemblyValidation.IconColor = System.Drawing.Color.Gainsboro
        Me.BtnAddUpdateAssemblyValidation.IconFont = FontAwesome.Sharp.IconFont.[Auto]
        Me.BtnAddUpdateAssemblyValidation.IconSize = 32
        Me.BtnAddUpdateAssemblyValidation.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnAddUpdateAssemblyValidation.Location = New System.Drawing.Point(0, 135)
        Me.BtnAddUpdateAssemblyValidation.Name = "BtnAddUpdateAssemblyValidation"
        Me.BtnAddUpdateAssemblyValidation.Padding = New System.Windows.Forms.Padding(30, 0, 10, 0)
        Me.BtnAddUpdateAssemblyValidation.Size = New System.Drawing.Size(213, 45)
        Me.BtnAddUpdateAssemblyValidation.TabIndex = 10
        Me.BtnAddUpdateAssemblyValidation.Text = "Assembly Validation"
        Me.BtnAddUpdateAssemblyValidation.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnAddUpdateAssemblyValidation.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.BtnAddUpdateAssemblyValidation.UseVisualStyleBackColor = False
        '
        'BtnAddUpdatePartSheetMetalUpdate
        '
        Me.BtnAddUpdatePartSheetMetalUpdate.BackColor = System.Drawing.Color.Transparent
        Me.BtnAddUpdatePartSheetMetalUpdate.Dock = System.Windows.Forms.DockStyle.Top
        Me.BtnAddUpdatePartSheetMetalUpdate.FlatAppearance.BorderSize = 0
        Me.BtnAddUpdatePartSheetMetalUpdate.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.BtnAddUpdatePartSheetMetalUpdate.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.BtnAddUpdatePartSheetMetalUpdate.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnAddUpdatePartSheetMetalUpdate.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAddUpdatePartSheetMetalUpdate.ForeColor = System.Drawing.Color.Gainsboro
        Me.BtnAddUpdatePartSheetMetalUpdate.IconChar = FontAwesome.Sharp.IconChar.None
        Me.BtnAddUpdatePartSheetMetalUpdate.IconColor = System.Drawing.Color.Gainsboro
        Me.BtnAddUpdatePartSheetMetalUpdate.IconFont = FontAwesome.Sharp.IconFont.[Auto]
        Me.BtnAddUpdatePartSheetMetalUpdate.IconSize = 32
        Me.BtnAddUpdatePartSheetMetalUpdate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnAddUpdatePartSheetMetalUpdate.Location = New System.Drawing.Point(0, 90)
        Me.BtnAddUpdatePartSheetMetalUpdate.Name = "BtnAddUpdatePartSheetMetalUpdate"
        Me.BtnAddUpdatePartSheetMetalUpdate.Padding = New System.Windows.Forms.Padding(30, 0, 10, 0)
        Me.BtnAddUpdatePartSheetMetalUpdate.Size = New System.Drawing.Size(213, 45)
        Me.BtnAddUpdatePartSheetMetalUpdate.TabIndex = 9
        Me.BtnAddUpdatePartSheetMetalUpdate.Text = "Part/ Sheet-Metal Update"
        Me.BtnAddUpdatePartSheetMetalUpdate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnAddUpdatePartSheetMetalUpdate.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.BtnAddUpdatePartSheetMetalUpdate.UseVisualStyleBackColor = False
        '
        'BtnAddUpdateNewPartCreation
        '
        Me.BtnAddUpdateNewPartCreation.BackColor = System.Drawing.Color.Transparent
        Me.BtnAddUpdateNewPartCreation.Dock = System.Windows.Forms.DockStyle.Top
        Me.BtnAddUpdateNewPartCreation.FlatAppearance.BorderSize = 0
        Me.BtnAddUpdateNewPartCreation.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.BtnAddUpdateNewPartCreation.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.BtnAddUpdateNewPartCreation.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnAddUpdateNewPartCreation.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAddUpdateNewPartCreation.ForeColor = System.Drawing.Color.Gainsboro
        Me.BtnAddUpdateNewPartCreation.IconChar = FontAwesome.Sharp.IconChar.None
        Me.BtnAddUpdateNewPartCreation.IconColor = System.Drawing.Color.Gainsboro
        Me.BtnAddUpdateNewPartCreation.IconFont = FontAwesome.Sharp.IconFont.[Auto]
        Me.BtnAddUpdateNewPartCreation.IconSize = 32
        Me.BtnAddUpdateNewPartCreation.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnAddUpdateNewPartCreation.Location = New System.Drawing.Point(0, 45)
        Me.BtnAddUpdateNewPartCreation.Name = "BtnAddUpdateNewPartCreation"
        Me.BtnAddUpdateNewPartCreation.Padding = New System.Windows.Forms.Padding(30, 0, 10, 0)
        Me.BtnAddUpdateNewPartCreation.Size = New System.Drawing.Size(213, 45)
        Me.BtnAddUpdateNewPartCreation.TabIndex = 8
        Me.BtnAddUpdateNewPartCreation.Text = "New Part Creation"
        Me.BtnAddUpdateNewPartCreation.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnAddUpdateNewPartCreation.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.BtnAddUpdateNewPartCreation.UseVisualStyleBackColor = False
        '
        'BtnAddUpdateVirtualStructure
        '
        Me.BtnAddUpdateVirtualStructure.BackColor = System.Drawing.Color.Transparent
        Me.BtnAddUpdateVirtualStructure.Dock = System.Windows.Forms.DockStyle.Top
        Me.BtnAddUpdateVirtualStructure.FlatAppearance.BorderSize = 0
        Me.BtnAddUpdateVirtualStructure.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.BtnAddUpdateVirtualStructure.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.BtnAddUpdateVirtualStructure.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnAddUpdateVirtualStructure.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAddUpdateVirtualStructure.ForeColor = System.Drawing.Color.Gainsboro
        Me.BtnAddUpdateVirtualStructure.IconChar = FontAwesome.Sharp.IconChar.None
        Me.BtnAddUpdateVirtualStructure.IconColor = System.Drawing.Color.Gainsboro
        Me.BtnAddUpdateVirtualStructure.IconFont = FontAwesome.Sharp.IconFont.[Auto]
        Me.BtnAddUpdateVirtualStructure.IconSize = 32
        Me.BtnAddUpdateVirtualStructure.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnAddUpdateVirtualStructure.Location = New System.Drawing.Point(0, 0)
        Me.BtnAddUpdateVirtualStructure.Name = "BtnAddUpdateVirtualStructure"
        Me.BtnAddUpdateVirtualStructure.Padding = New System.Windows.Forms.Padding(30, 0, 10, 0)
        Me.BtnAddUpdateVirtualStructure.Size = New System.Drawing.Size(213, 45)
        Me.BtnAddUpdateVirtualStructure.TabIndex = 7
        Me.BtnAddUpdateVirtualStructure.Text = "Virtual Structure"
        Me.BtnAddUpdateVirtualStructure.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnAddUpdateVirtualStructure.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.BtnAddUpdateVirtualStructure.UseVisualStyleBackColor = False
        '
        'BtnAddUpdate
        '
        Me.BtnAddUpdate.BackColor = System.Drawing.Color.Transparent
        Me.BtnAddUpdate.Dock = System.Windows.Forms.DockStyle.Top
        Me.BtnAddUpdate.FlatAppearance.BorderSize = 0
        Me.BtnAddUpdate.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.BtnAddUpdate.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.BtnAddUpdate.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnAddUpdate.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAddUpdate.ForeColor = System.Drawing.Color.Gainsboro
        Me.BtnAddUpdate.IconChar = FontAwesome.Sharp.IconChar.Plus
        Me.BtnAddUpdate.IconColor = System.Drawing.Color.Gainsboro
        Me.BtnAddUpdate.IconFont = FontAwesome.Sharp.IconFont.[Auto]
        Me.BtnAddUpdate.IconSize = 32
        Me.BtnAddUpdate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnAddUpdate.Location = New System.Drawing.Point(0, 75)
        Me.BtnAddUpdate.Name = "BtnAddUpdate"
        Me.BtnAddUpdate.Padding = New System.Windows.Forms.Padding(10, 0, 10, 0)
        Me.BtnAddUpdate.Size = New System.Drawing.Size(213, 60)
        Me.BtnAddUpdate.TabIndex = 3
        Me.BtnAddUpdate.Text = "Add/ Update"
        Me.BtnAddUpdate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnAddUpdate.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.BtnAddUpdate.UseVisualStyleBackColor = False
        '
        'logoPanel
        '
        Me.logoPanel.BackColor = System.Drawing.Color.Transparent
        Me.logoPanel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.logoPanel.Controls.Add(Me.BtnHome)
        Me.logoPanel.Dock = System.Windows.Forms.DockStyle.Top
        Me.logoPanel.Location = New System.Drawing.Point(0, 0)
        Me.logoPanel.Name = "logoPanel"
        Me.logoPanel.Size = New System.Drawing.Size(213, 75)
        Me.logoPanel.TabIndex = 2
        '
        'BtnHome
        '
        Me.BtnHome.BackColor = System.Drawing.Color.Transparent
        Me.BtnHome.BackgroundImage = Global.SolidEdgeApp.My.Resources.Resources.bec_logo
        Me.BtnHome.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.BtnHome.Location = New System.Drawing.Point(1, 17)
        Me.BtnHome.Name = "BtnHome"
        Me.BtnHome.Size = New System.Drawing.Size(229, 41)
        Me.BtnHome.TabIndex = 1
        Me.BtnHome.TabStop = False
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.Text = "NotifyIcon1"
        Me.NotifyIcon1.Visible = True
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(1284, 711)
        Me.Controls.Add(Me.pnlFooterShadow)
        Me.Controls.Add(Me.panelDesktop)
        Me.Controls.Add(Me.panelTitlBar)
        Me.Controls.Add(Me.SideMenuPanel)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.KeyPreview = True
        Me.MinimumSize = New System.Drawing.Size(1300, 726)
        Me.Name = "MainForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Child Form Template"
        Me.panelTitlBar.ResumeLayout(False)
        Me.panelTitlBar.PerformLayout()
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.FlowLayoutPanel1.PerformLayout()
        CType(Me.IconCurrentForm, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SideMenuPanel.ResumeLayout(False)
        Me.panelQCSubMenu.ResumeLayout(False)
        Me.panelDesignSubMenu.ResumeLayout(False)
        Me.panelAddUpdateSubMenu.ResumeLayout(False)
        Me.logoPanel.ResumeLayout(False)
        CType(Me.BtnHome, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SideMenuPanel As MoCustomControls.MoPanel
    Friend WithEvents panelQCSubMenu As MoCustomControls.MoPanel
    Friend WithEvents btnQCMTC As FontAwesome.Sharp.IconButton
    Friend WithEvents btnQCInterference As FontAwesome.Sharp.IconButton
    Friend WithEvents BtnQC As FontAwesome.Sharp.IconButton
    Friend WithEvents panelDesignSubMenu As MoCustomControls.MoPanel
    Friend WithEvents BtnDesignOccurenceProperties As FontAwesome.Sharp.IconButton
    Friend WithEvents BtnDesignCopyTransfer As FontAwesome.Sharp.IconButton
    Friend WithEvents BtnDesign As FontAwesome.Sharp.IconButton
    Friend WithEvents panelAddUpdateSubMenu As MoCustomControls.MoPanel
    Friend WithEvents BtnAddUpdateAssemblyValidation As FontAwesome.Sharp.IconButton
    Friend WithEvents BtnAddUpdatePartSheetMetalUpdate As FontAwesome.Sharp.IconButton
    Friend WithEvents BtnAddUpdateNewPartCreation As FontAwesome.Sharp.IconButton
    Friend WithEvents BtnAddUpdate As FontAwesome.Sharp.IconButton
    Friend WithEvents logoPanel As Panel
    Friend WithEvents BtnHome As PictureBox
    Friend WithEvents panelTitlBar As Panel
    Friend WithEvents BtnMinimize As FontAwesome.Sharp.IconButton
    Friend WithEvents BtnMaximize As FontAwesome.Sharp.IconButton
    Friend WithEvents BtnExit As FontAwesome.Sharp.IconButton
    Friend WithEvents lblFormTitle As Label
    Friend WithEvents IconCurrentForm As FontAwesome.Sharp.IconPictureBox
    Friend WithEvents panelDesktop As Panel
    Friend WithEvents pnlFooterShadow As Panel
    Friend WithEvents BtnAddUpdateVirtualStructure As FontAwesome.Sharp.IconButton
    Friend WithEvents BtnConfiguration As FontAwesome.Sharp.IconButton
    Friend WithEvents BtnQCRawMaterialEstimation As FontAwesome.Sharp.IconButton
    Friend WithEvents BtnQCKPI As FontAwesome.Sharp.IconButton
    Friend WithEvents FlowLayoutPanel1 As FlowLayoutPanel
    Friend WithEvents BtnVersion As FontAwesome.Sharp.IconButton
    Friend WithEvents lblVersion As Label
    Friend WithEvents BtnHelp As FontAwesome.Sharp.IconButton
    Friend WithEvents LblHelp As Label
    Friend WithEvents NotifyIcon1 As NotifyIcon
    Friend WithEvents BeC_Automation_Installer1 As BEC_Automation_Installer
End Class
