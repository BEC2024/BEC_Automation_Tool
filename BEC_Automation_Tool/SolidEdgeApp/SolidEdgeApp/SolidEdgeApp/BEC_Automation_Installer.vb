Imports System.ComponentModel
Imports System.Configuration.Install

Public Class BEC_Automation_Installer

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add initialization code after the call to InitializeComponent

    End Sub

    Private Sub BEC_Automation_Installer_AfterUninstall(sender As Object, e As InstallEventArgs) Handles MyBase.AfterUninstall
        Dim path As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\" + "BECAutomation"
        'Public Shared configFilePath1 As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\" + "BECAutomation" + "\ConfigProperties.xml"
        If System.IO.Directory.Exists(path) Then
            System.IO.Directory.Delete(path, True)
        End If
    End Sub


End Class
