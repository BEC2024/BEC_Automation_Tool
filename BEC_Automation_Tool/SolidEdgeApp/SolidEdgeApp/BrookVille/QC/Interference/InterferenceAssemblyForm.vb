Imports System.Runtime.InteropServices
'Imports SolidEdge.Framework.Interop
Imports SolidEdgeFramework
Public Class InterferenceAssemblyForm

    Dim objApp As SolidEdgeFramework.Application = Nothing

    Dim objAssemblyDocument As SolidEdgeAssembly.AssemblyDocument

    Private Sub InterferenceAssemblyForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        objApp = SetSolidEdgeInstance()

        If objApp Is Nothing Then
            MessageBox.Show("Please Open SolidEdge")
            Me.Close()
            Exit Sub
        End If

        objAssemblyDocument = objApp.ActiveDocument

        Me.Text = $"{Me.Text} ({GlobalEntity.Version})"

    End Sub

    Private Function SetSolidEdgeInstance()
        Try
            objApp = Marshal.GetActiveObject("SolidEdge.Application")
        Catch ex As Exception
            '  MessageBox.Show("Please Open SolidEdge")
            '  Me.Close()
            '  MsgBox($"Error in fetching the Solid-Edge instance {ex.Message} {vbNewLine} {ex.StackTrace}")
        End Try
        Return objApp

    End Function

    Private Function GetReferenceAssemblySelectSet() As SelectSet

        Dim objReferenceDocSelectSet As SolidEdgeFramework.SelectSet = objAssemblyDocument.RelationshipsSelectSet
        Return objReferenceDocSelectSet

    End Function

    Private Sub btnCheckInterference_Click(sender As Object, e As EventArgs) Handles btnCheckInterference.Click

        Dim objSelectSet As SolidEdgeFramework.SelectSet = GetReferenceAssemblySelectSet()

        Dim invalidSelectDoc As String = String.Empty

        Dim isValidSelect As Boolean = IsValidSelection(objSelectSet, invalidSelectDoc)

        If Not isValidSelect Then
            MessageBox.Show("Please select all top level assembly only")
        End If


    End Sub

    Private Function IsValidSelection(ByRef objSelectSet As SolidEdgeFramework.SelectSet, ByRef inValidSelectedDoc As String) As Boolean

        Dim isvalid As Boolean = True
        For Each occur As Object In objSelectSet
            Dim occurenceType As ObjectType = occur.Type
            If Not occurenceType = ObjectType.igSubAssembly Then
                isvalid = False
                Exit For
            End If
        Next

        Return isvalid

    End Function

End Class