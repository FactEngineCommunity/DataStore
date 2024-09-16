Imports FactEngineForServices

Public Module tableDataStore

    Public Function AddData(ByVal arData As DataStore.Data) As Boolean

        Dim lsSQLQuery As String = ""

        Try
            lsSQLQuery = "INSERT INTO DataStore (ID, Data, DataClass, EntryDate, Source, Author, Tags, Location, Notes)"
            lsSQLQuery &= " VALUES ("
            lsSQLQuery &= "'" & arData.ID & "'"
            lsSQLQuery &= " ,'" & Trim(Replace(arData.Data, "'", "`")) & "'"
            lsSQLQuery &= " ,'" & Trim(Replace(arData.DataClass, "'", "`")) & "'"
            lsSQLQuery &= " ," & pdbConnection.DateWrap(Now.ToString("yyyy/MM/dd HH:mm:ss"))
            lsSQLQuery &= " ,'" & Trim(Replace(arData.Source, "'", "`")) & "'"
            lsSQLQuery &= " ,'" & Trim(Replace(arData.Author, "'", "`")) & "'"
            lsSQLQuery &= " ,'" & Trim(Replace(arData.Tags, "'", "`")) & "'"
            lsSQLQuery &= " ,'" & Trim(Replace(arData.Location, "'", "`")) & "'"
            lsSQLQuery &= " ,'" & Trim(Replace(arData.Notes, "'", "`")) & "'"
            lsSQLQuery &= ")"

            Dim lrRecordset As ORMQL.Recordset = pdbConnection.Execute(lsSQLQuery)

            If lrRecordset.ErrorReturned Then
                Throw New Exception(lrRecordset.ErrorString)
            End If

            Return True

        Catch ex As Exception
            Dim lsMessage As String
            lsMessage = "Error: AddData"
            lsMessage &= vbCrLf & vbCrLf & ex.Message
            prApplication.ThrowErrorMessage(lsMessage, pcenumErrorType.Critical, ex.StackTrace,,)

            Return False

        End Try

    End Function

End Module
