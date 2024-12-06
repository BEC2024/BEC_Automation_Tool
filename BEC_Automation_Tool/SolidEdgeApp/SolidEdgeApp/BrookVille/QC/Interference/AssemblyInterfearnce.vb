Imports System.Runtime.InteropServices
'Imports SolidEdge.Framework.Interop
Imports SolidEdgeFramework
Public Class AssemblyInterfearnce
    Dim objApp As SolidEdgeFramework.Application = Nothing
    Dim objAsm As SolidEdgeAssembly.AssemblyDocument = Nothing

    Private Sub AssemblyInterfearnce_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        objApp = Marshal.GetActiveObject("SolidEdge.Application")
        objAsm = objApp.ActiveDocument
        Me.Text = $"{Me.Text} ({GlobalEntity.Version})"

    End Sub

    Dim cnt As Integer = 1

    Public Sub interfearnce()
        Try
            Dim nComparisonMethod As Integer
            Dim nSet1 As Integer
            Dim nSet2 As Integer
            Dim objOccurrences As SolidEdgeAssembly.Occurrences = Nothing
            Dim objOccurrence As SolidEdgeAssembly.Occurrence = Nothing
            Dim objInterfOcc As SolidEdgeAssembly.Occurrence = Nothing
            Dim a_objSet1() As Object
            Dim objTemp As Object
            Dim nNumInterferences As Long
            Dim nStatus As SolidEdgeAssembly.InterferenceStatusConstants
            ' Dim reportPath As String = $"C:\Users\vimalb\Desktop\New Text Document{cnt.ToString()}.txt"
            Dim fullpath As String = objAsm.FullName
            Dim reportname As Object = IO.Path.GetDirectoryName(fullpath) + "\" + "report.txt"
            objApp.DisplayAlerts = False
            objAsm = objApp.ActiveDocument
            nComparisonMethod = SolidEdgeConstants.InterferenceComparisonConstants.seInterferenceComparisonSet1vsAllOther
            nSet1 = 0
            nSet2 = 0
            objOccurrences = objAsm.Occurrences

            ReDim a_objSet1(0)
            For nIndex = 1 To objOccurrences.Count
                objTemp = objOccurrences.Item(nIndex)
                Dim type As ObjectType = objTemp.type
                a_objSet1(0) = objTemp
                nSet1 = 1

                'Add intereference as part in assembly
                Call objAsm.CheckInterference2(
                NumElementsSet1:=nSet1,
                Set1:=a_objSet1,
                Status:=nStatus,
                ComparisonMethod:=nComparisonMethod,
                AddInterferenceAsOccurrence:=True,
                NumInterferences:=nNumInterferences,
                InterferenceOccurrence:=objInterfOcc, ReportFilename:=reportname) ', ReportType:="TEXT") ',
                RenameFile(cnt, reportname)

                cnt = cnt + 1
                ' a_objSet1 = Nothing
            Next nIndex
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RenameFile(ByVal cnt As Integer, ByVal reportname As String)

        System.IO.File.Move(reportname, $"{reportname.Replace(".txt", $"{cnt.ToString}.txt")}")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        interfearnce()
    End Sub

End Class