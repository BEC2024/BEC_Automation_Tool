Imports System.Runtime.InteropServices
Imports System.Text
'Imports SolidEdge.Framework.Interop
Imports SolidEdgePart
Imports SolidEdgeFramework
Imports SolidEdgeAssembly

Public Class CopyPartForm

    Dim objApp As SolidEdgeFramework.Application = Nothing
    Dim objAsm As SolidEdgeAssembly.AssemblyDocument = Nothing
    Dim mainObj As New MainClass
    Public Sub SetSolidEdgeInstance()
        Try
            objApp = Marshal.GetActiveObject("SolidEdge.Application")
            objAsm = objApp.ActiveDocument
        Catch ex As Exception
        End Try

    End Sub
    Public Sub DisableBtn()
        If objApp Is Nothing And objApp Is Nothing Then
            btnUpdateAll.Enabled = False
            Button1.Enabled = False
            btnCopyPartForm.Enabled = False
        Else
            btnUpdateAll.Enabled = True
            Button1.Enabled = True
            btnCopyPartForm.Enabled = True
        End If
    End Sub
    Private Function IsValid() As Boolean

        SetSolidEdgeInstance()
        DisableBtn()
        If objApp Is Nothing And objAsm Is Nothing Then

            Return False
        Else
            Return True

        End If

    End Function

    Private Sub CopyPartForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CustomLogUtil.Heading("Copy && Transfer Part From Open.....")
        If IsValid() Then
            Me.Text = $"{Me.Text} ({GlobalEntity.Version})"
        Else
            MessageBox.Show($"Please Open the Solid-Edge Assembly and Restart the Application", "Message")
            CustomLogUtil.Log($"Please Open the Solid-Edge Assembly and Restart the Application", "", "")
        End If

    End Sub
    Public Sub Closefn(mainObj As MainClass)
        mainObj.SolidEdgeinstance = "Close"
    End Sub
    Private Sub Abc()

        Dim selectSet As SelectSet = objAsm.SelectSet
        Dim partDocument As SolidEdgePart.PartDocument = Nothing

        For Each item As Object In selectSet

            Dim occurances As SolidEdgeAssembly.Occurrences = objAsm.Occurrences
            Dim a As Integer = item.OccurrenceID
            Dim occur As Occurrence = occurances.GetOccurrence(a)


            Dim X As Double = Nothing
            Dim Y As Double = Nothing
            Dim Z As Double = Nothing
            Dim Anglex As Double = Nothing
            Dim angley As Double = Nothing
            Dim anglez As Double = Nothing

            occur.GetTransform(X, Y, Z, Anglex, angley, anglez)
            Dim newoccur As Occurrence = occurances.AddWithTransform(item.Occurrencedocument.fullname, X, Y, Z, Anglex, angley, anglez)



        Next

        MessageBox.Show("Process completed", "Message")
    End Sub
    Private Sub BtnCopyPartForm_Click(sender As Object, e As EventArgs) Handles btnCopyPartForm.Click

        Dim selectSet As SolidEdgeFramework.SelectSet = objAsm.SelectSet
        Dim partDocument As SolidEdgePart.PartDocument = Nothing

        For Each item As Object In selectSet

            Dim occurances As Occurrences = objAsm.Occurrences
            Dim a As Integer = item.OccurrenceID
            Dim occur As Occurrence = occurances.GetOccurrence(a)


            Dim X As Double = Nothing
            Dim Y As Double = Nothing
            Dim Z As Double = Nothing
            Dim Anglex As Double = Nothing
            Dim angley As Double = Nothing
            Dim anglez As Double = Nothing

            occur.GetTransform(X, Y, Z, Anglex, angley, anglez)
            Dim newoccur As Occurrence = occurances.AddWithTransform(item.Occurrencedocument.fullname, X, Y, Z, Anglex, angley, anglez)
            If newoccur.Style IsNot Nothing Then
                newoccur.Style = "Glass"
            End If

        Next

        '40236 > Update all open documents command
        objApp.StartCommand(40236)

        MessageBox.Show("Process completed", "Message")
        CustomLogUtil.Heading("Copy Part Same Location : proccess done......")
    End Sub

    Private Sub RemoveAllSelection()

        objAsm.SelectSet.RemoveAll()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        'Update all open document command
        'objApp.StartCommand(40236)


        ''transfer document command
        objApp.StartCommand(40256)

        Exit Sub

        Dim selectSet As SolidEdgeFramework.SelectSet = objAsm.SelectSet
        Dim partDocument As SolidEdgePart.PartDocument = Nothing

        For Each item As Object In selectSet

            Dim occurances As Occurrences = objAsm.Occurrences
            Dim a As Integer = item.OccurrenceID
            Dim occur As Occurrence = occurances.GetOccurrence(a)

            Dim X As Double = Nothing
            Dim Y As Double = Nothing
            Dim Z As Double = Nothing
            Dim Anglex As Double = Nothing
            Dim angley As Double = Nothing
            Dim anglez As Double = Nothing
            occur.GetTransform(X, Y, Z, Anglex, angley, anglez)


            RemoveAllSelection()

            MessageBox.Show("Please select assembly", "Message")

            Dim selectSet2 As SolidEdgeFramework.SelectSet = objAsm.SelectSet

            For Each item2 As Object In selectSet2

                'Dim occurances As Occurrences = objAsm.Occurrences
                Dim a1 As Integer = item2.OccurrenceID
                Dim occur2 As Occurrence = occurances.GetOccurrence(a1)
                Dim selectedAssemblyDoc As AssemblyDocument = occur2.OccurrenceDocument
                Dim selectedAssemblyOccurences As Occurrences = selectedAssemblyDoc.Occurrences


                Dim newoccur As Occurrence = selectedAssemblyOccurences.AddWithTransform(item.Occurrencedocument.fullname, X, Y, Z, Anglex, angley, anglez)
                newoccur.Visible = False

                'newoccur.Parent = selectedAssemblyDoc

                selectedAssemblyDoc.UpdateAll()

                Debug.Print("aaa")
                Exit For

            Next

            objAsm.UpdateAll()

            'Dim occurences2 As Occurrences = objAsm.Occurrences

            'occur.GetTransform(X, Y, Z, Anglex, angley, anglez)

            'Dim newoccur As Occurrence = occurances.AddWithTransform(item.Occurrencedocument.fullname, X, Y, Z, Anglex, angley, anglez)

        Next

        MessageBox.Show("Process completed", "Message")
        CustomLogUtil.Heading($"Transfer Selected Document : proccess done......")
    End Sub

    Private Sub BtnUpdateAll_Click(sender As Object, e As EventArgs) Handles btnUpdateAll.Click

        'Update all open document command
        objApp.StartCommand(40236)
        CustomLogUtil.Heading("Update All Open Documents : proccess done......")
    End Sub
End Class