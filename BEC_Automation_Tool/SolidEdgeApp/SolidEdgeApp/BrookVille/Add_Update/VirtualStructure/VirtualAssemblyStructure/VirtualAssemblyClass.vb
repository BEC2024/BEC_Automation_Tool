Public Class VirtualAssemblyClass

    Public mainAssemblyName As String
    Public dict1 As Dictionary(Of String, List(Of String)) = New Dictionary(Of String, List(Of String))
    Public dicSubAssemblyDetails As Dictionary(Of String, VirtualAssemblyClass) = New Dictionary(Of String, VirtualAssemblyClass)()

End Class