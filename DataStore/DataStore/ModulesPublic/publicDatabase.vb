Imports System.IO
Imports System.Reflection
Imports System.Data.OleDb
Imports System.Configuration
Imports ADOX
Imports FactEngineForServices
Imports System.IO.Path
Imports System.Globalization
Imports System.Data.Common

Namespace Database

    Public Module DatabaseModule

        Public Function OpenDatabase(ByVal asDatabaseConnectionString) As Boolean

            Dim lsMessage As String = ""

            Try
                '============================================================================
                pdbConnection = New FactEngine.SQLiteConnection(Nothing, asDatabaseConnectionString, 1000, True)

                '------------------------------------------------
                'Open the (database) connection
                '------------------------------------------------
                If Not pdbConnection.Open(asDatabaseConnectionString) Then
                    Throw New Exception("Failed to Open database. Method: Open.")
                End If

                Return True

            Catch lo_ex As Exception

                Dim mb As MethodBase = MethodInfo.GetCurrentMethod()

                lsMessage = "Error: " & mb.ReflectedType.Name & "." & mb.Name
                lsMessage.AppendDoubleLineBreak("Error: There was an error opening the database: ")
                lsMessage.AppendDoubleLineBreak(asDatabaseConnectionString)
                lsMessage.AppendDoubleLineBreak("'" & Trim(lo_ex.Message) & "'" & vbCrLf & lo_ex.StackTrace)
                Throw New Exception(lsMessage)
            End Try

        End Function

        Public Function MakeStringSafe(ByVal asString As String) As String

            Dim lsReturnString As String = ""

            lsReturnString = asString.Replace("""", """")
            lsReturnString = asString.Replace("'", "''")


            Return lsReturnString

        End Function

        Public Function RevertString(ByVal asString As String) As String

            Dim lsReturnString As String = ""

            lsReturnString = asString.Replace("''", "`")

            Return lsReturnString

        End Function

    End Module

End Namespace




