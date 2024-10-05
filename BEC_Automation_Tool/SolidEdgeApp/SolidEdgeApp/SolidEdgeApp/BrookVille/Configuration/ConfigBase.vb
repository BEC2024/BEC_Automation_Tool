Imports System.Windows.Forms

Public MustInherit Class ConfigBase

    Public configFilePath As String

    Public Sub saveConfig()

        Dim configDirectory As String = IO.Path.GetDirectoryName(configFilePath)
        Dim dir As New IO.DirectoryInfo(configDirectory)
        If Not dir.Exists Then
            dir.Create()
        End If

        Try
            changeConfig2()
            saveConfig2()
        Catch ex As Exception
            IO.File.Delete(configFilePath)
            changeConfig2()
            saveConfig2()
        End Try

    End Sub

    Public Sub readConfig()

        'If Not IO.File.Exists(configFilePath) Then
        '    Exit Sub
        'End If

        Try
            'If (Not System.IO.File.Exists(configFilePath)) Then
            '    saveConfig2()
            'End If
            readConfig2()
        Catch ex As Exception
            'IO.File.Delete(configFilePath)
            MessageBox.Show("Error reading configuration", "Configuration Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CustomLogUtil.Log("Error reading configuration", ex.Message, ex.StackTrace)
        End Try

    End Sub
    Public Sub ChangeConfigTxt()
        ChangeConfigTxt2()
    End Sub
    Public Sub changeConfig()
        changeConfig2()
    End Sub

    'saves this object to config file
    Protected MustOverride Sub saveConfig2()

    'reads config file and initializes all fields
    Protected MustOverride Sub readConfig2()

    Protected MustOverride Sub changeConfig2()
    Protected MustOverride Sub ChangeConfigTxt2()
End Class
