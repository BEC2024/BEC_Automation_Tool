Imports System.Runtime.InteropServices
Imports System.Text
'Imports SolidEdge.Framework.Interop
Imports SolidEdgePart
Imports SolidEdgeFramework
Public Class OccurencePropertiesUpdateForm

    Dim objApp As SolidEdgeFramework.Application = Nothing
    Dim objAssemblyDocument As SolidEdgeAssembly.AssemblyDocument

    Public Class OccurenceProperties
        Public higherLevel As Boolean = False
        Public draftReference As Boolean = False
        Public reportPartList As Boolean = False
        Public drawingViews As Boolean = False
        Public physicalProperties As Boolean = False
        Public interferenceAnalysis As Boolean = False

        'Public referenceOnly As Boolean = False
    End Class
    Dim mainObj As New MainClass
    Private Function IsValid() As Boolean

        SetSolidEdgeInstance()
        DisableTableLayoutPanel1()
        If objApp Is Nothing Then
            Return False
        Else
            Return True
        End If

    End Function
    Public Sub DisableTableLayoutPanel1()
        If objApp Is Nothing Then
            TableLayoutPanel1.Enabled = False
        Else
            TableLayoutPanel1.Enabled = True
        End If
    End Sub
    Private Sub OccurencePropertiesUpdateForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CustomLogUtil.Heading("Occurence Properties Form Open.....")
        If IsValid() Then
            objApp = SetSolidEdgeInstance()
            objAssemblyDocument = objApp.ActiveDocument
            Me.Text = $"{Me.Text} ({GlobalEntity.Version})"
        Else

            MessageBox.Show("Please Open Solid-Edge Assembly and Restart the Application", "Message")
            CustomLogUtil.Log("Please Open Solid-Edge Assembly and Restart the Application", "", "")
            Exit Sub
        End If




    End Sub
    Public Sub Closefn(mainObj As MainClass)
        mainObj.SolidEdgeinstance = "Close"
    End Sub
    '  Dim objSelectSet As SolidEdgeFramework.SelectSet = GetReferenceAssemblySelectSet()

    ' ReadOccurenceProperty(objSelectSet)



    Private Function SetSolidEdgeInstance()
        Try
            objApp = Marshal.GetActiveObject("SolidEdge.Application")


        Catch ex As Exception

            '  MessageBox.Show("Please Open SolidEdge")
            '  Me.Close()

            ' MsgBox($"Error in fetching the Solid-Edge instance {ex.Message} {vbNewLine} {ex.StackTrace}")
        End Try
        Return objApp

    End Function

    Private Sub BtnUpdateProperties_Click(sender As Object, e As EventArgs) Handles btnUpdateProperties.Click

        Dim objSelectSet As SolidEdgeFramework.SelectSet = GetReferenceAssemblySelectSet()

        Dim occProp As OccurenceProperties = UpdateOccurencePropertyDetails()

        UpdateProperties(objSelectSet, occProp)



        objAssemblyDocument.Save()
        MsgBox("Update properties completed.")
        CustomLogUtil.Heading("Update Occurence Properties completed.....")


    End Sub

    Private Sub ReadOccurenceProperty(ByRef objSelectSet As SolidEdgeFramework.SelectSet)

        For Each occur As Object In objSelectSet

            Dim occurenceType As ObjectType = occur.Type
            If occurenceType = ObjectType.igSubAssembly Then



                'chkHigherLevel.Checked = occur.DisplayinSubAssembly
                'chkReportPartList.Checked = occur.includeinBOM
                'chkDrawingViews.Checked = occur.DisplayinDrawings
                'chkPhysicalProperties.Checked = occur.IncludeinphysicalProperties
                'chkInteferenceAnalysis.Checked = occur.Includeininterference
                ''chkreferenceonly.Checked = occur.ReferenceOnly
            End If
            If occurenceType = ObjectType.igPart Then
                MessageBox.Show("Please select assembly", "Message")
                Me.Close()
            End If
        Next

    End Sub




    Private Function UpdateOccurencePropertyDetails() As OccurenceProperties

        'occProp.draftReference = chkDraftReference.Checked
        Dim occProp As New OccurenceProperties With {
            .higherLevel = rbYes_HighLevel.Checked,
            .reportPartList = rbYes_ReportPartList.Checked,
            .drawingViews = rbYes_DrawingViews.Checked,
            .physicalProperties = rbYes_PhysicalProperties.Checked,
            .interferenceAnalysis = rbYes_InterfernceAnalysis.Checked
        }
        'occProp.referenceOnly = chkreferenceonly.Checked

        Return occProp

    End Function

    Private Sub UpdateProperties(ByRef objSelectSet As SolidEdgeFramework.SelectSet, ByVal occProp As OccurenceProperties)
        For Each occur As Object In objSelectSet

            Dim occurenceType As ObjectType = occur.Type
            If occurenceType = ObjectType.igSubAssembly Or occurenceType = ObjectType.igPart Then

                If rbYes_ApplyColor.Checked Then
                    occur.style = "White (clear)"
                End If

                occur.DisplayinSubAssembly = occProp.higherLevel
                occur.includeinBOM = occProp.reportPartList
                occur.DisplayinDrawings = occProp.drawingViews
                occur.IncludeinphysicalProperties = occProp.physicalProperties
                occur.Includeininterference = occProp.interferenceAnalysis
                'occur.Referenceonly = occProp.referenceOnly
            End If
            If occurenceType = ObjectType.igReference Then
                If rbYes_ApplyColor.Checked Then
                    occur.object.style = "White (clear)"
                End If

                occur.object.DisplayinSubAssembly = occProp.higherLevel
                occur.object.includeinBOM = occProp.reportPartList
                occur.object.DisplayinDrawings = occProp.drawingViews
                occur.object.IncludeinphysicalProperties = occProp.physicalProperties
                occur.object.Includeininterference = occProp.interferenceAnalysis

            End If
        Next
    End Sub

    Private Function GetReferenceAssemblySelectSet() As SolidEdgeFramework.SelectSet
        Dim objReferenceDocSelectSet As SolidEdgeFramework.SelectSet = objAssemblyDocument.RelationshipsSelectSet
        Return objReferenceDocSelectSet
    End Function


End Class