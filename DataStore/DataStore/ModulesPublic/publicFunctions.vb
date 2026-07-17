Imports System.IO
Imports System.Reflection

Public Module publicFunctions

    ' Reads an embedded text resource by file name (suffix match), no LINQ needed.
    Public Function ReadEmbeddedText(ByVal resourceFileName As String) As String
        Dim asm As Assembly = Assembly.GetExecutingAssembly()
        Dim resName As String = Nothing
        For Each n As String In asm.GetManifestResourceNames()
            If n.EndsWith(resourceFileName, StringComparison.OrdinalIgnoreCase) Then
                resName = n
                Exit For
            End If
        Next
        If String.IsNullOrEmpty(resName) Then Return Nothing

        Using s As Stream = asm.GetManifestResourceStream(resName)
            If s Is Nothing Then Return Nothing
            Using sr As New StreamReader(s)
                Return sr.ReadToEnd()
            End Using
        End Using
    End Function

End Module
