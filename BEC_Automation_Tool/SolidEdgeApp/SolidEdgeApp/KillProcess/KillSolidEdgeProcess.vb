Public Class KillSolidEdgeProcess
    Public Shared Sub Kill()
        Try
            For Each p As Process In System.Diagnostics.Process.GetProcessesByName("Edge")
                p.Kill()
            Next
        Catch ex As Exception

        End Try

    End Sub

    Public Shared Sub killSilent()
        Dim count As Integer = Nothing
        For Each p As Process In System.Diagnostics.Process.GetProcessesByName("Edge")
            count += 1
        Next

        If count > 1 Then
            For Each p As Process In System.Diagnostics.Process.GetProcessesByName("Edge")
                count -= 1
                p.Kill()
                If count = 1 Then
                    Exit Sub
                End If
            Next
        ElseIf count = 1 Then
            For Each p As Process In System.Diagnostics.Process.GetProcessesByName("Edge")
                p.Kill()
            Next
        End If



    End Sub
End Class
