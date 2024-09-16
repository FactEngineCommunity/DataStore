Imports System.Data.SQLite
Imports System.Linq.Expressions
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Serialization
Imports System.IO
Imports System.Text.RegularExpressions
Imports FactEngineForServices
Imports System.Text
Imports System.Runtime.CompilerServices
Imports System.Reflection

Namespace DataStore
    Public Class [Store]

        'EXAMPLE
        ' Define the LINQ expression for the condition
        'Dim lrDataStore As New DataStore.Store
        'Dim whereClause As Expression(Of Func(Of FBM.DictionaryEntry, Boolean)) = Function(p) p.Symbol = "Satellite"
        'Dim larDictionaryEntry As List(Of FBM.DictionaryEntry) = lrDataStore.GetData(Of FBM.DictionaryEntry)(whereClause)

        ''' <summary>
        '''  AddObject method to add a new record to the DataStore table
        ''' </summary>
        ''' <param name="arObject"></param>
        Public Sub Add(arObject As Object)
            Try
                'Usage Example:
                'Dim lrEntityType As New FBM.EntityType(Nothing, pcenumLanguage.ORMModel, "Test Entity Type", Nothing, True)
                'Dim lrDataStore As New DataStore.Store
                'Call lrDataStore.AddObject(lrEntityType)

                ' Create the custom JSON serializer settings
                Dim settings As New JsonSerializerSettings With {
                    .Formatting = Formatting.Indented,
                   .TypeNameHandling = TypeNameHandling.Objects,
                   .ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                   }

                ' Serialize the object into JSON with custom settings
                Dim jsonData As String = JsonConvert.SerializeObject(arObject, settings)

                ' Get the type name of the object
                Dim typeName As String = arObject.GetType.FullName

                ' Create the SQL query to insert the new record into the DataStore table
                Dim lrData As New DataStore.Data(jsonData, typeName, Now, "", "", "", "", "")

                ' Execute the insert query
                Call tableDataStore.AddData(lrData)

            Catch ex As Exception
                ' Handle the exception appropriately
                ' ...
            End Try
        End Sub

        ''' <summary>
        ''' Delete method
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="whereClause"></param>
        Public Sub Delete(Of T)(whereClause As Expression(Of Func(Of T, Boolean)))
            Try
                'Usage Examples
                'Dim whereClause As Expression(Of Func(Of FBM.DictionaryEntry, Boolean)) = Function(p) p.Symbol = "Satellite"
                'Dim larDictionaryEntry As List(Of FBM.DictionaryEntry) = lrDataStore.GetData(Of FBM.DictionaryEntry)(whereClause)
                'or
                '            Dim lrEntityType As New FBM.EntityType(Nothing, pcenumLanguage.ORMModel, "Test Entity Type", Nothing, True)
                'Dim whereClause As Expression(Of Func(Of FBM.EntityType, Boolean)) = Function(p) p.Id = "Test Entity Type"
                'Dim lrDataStore As New DataStore.Store
                'Call lrDataStore.DeleteData(Of FBM.EntityType)(whereClause)

                ' Get the record that matches the provided whereClause using GetData function
                Dim asID As String = Nothing
                Dim dataList As List(Of T) = Me.Get(Of T)(whereClause, asID)

                ' Assuming there should be only one record that matches the condition
                If dataList.Count >= 1 And asID IsNot Nothing Then

                    For Each lrDataItem In dataList
                        Dim recordToUpdate As T = lrDataItem

                        ' Now you can update the record back in the database using prConnection.Execute
                        ' Here's the SQL query to update the Data field for the specific record identified by ID
                        Dim lsSQLQuery As String = "DELETE FROM DataStore WHERE ID = '" & asID & "'"

                        ' Execute the update query
                        Call pdbConnection.Execute(lsSQLQuery)
                    Next
                End If

            Catch ex As Exception
                ' Handle the exception appropriately
                ' ...
            End Try
        End Sub

        Public Function [Get](Of T)(Optional whereClause As Expression(Of Func(Of T, Boolean)) = Nothing, Optional ByRef asID As String = Nothing) As List(Of T)

            Dim dataList As New List(Of T)()

            Try
                'Usage Example:
                'Dim whereClause As Expression(Of Func(Of FBM.EntityType, Boolean)) = Function(p) p.Id = "Test Entity Type"
                'Dim lrDataStore As New DataStore.Store
                'Dim larDictionaryEntry As List(Of FBM.EntityType) = lrDataStore.GetData(Of FBM.EntityType)(whereClause)

                Dim typeName As String = GetType(T).FullName

                Dim lsSQLQuery As String
                If whereClause Is Nothing Then
                    lsSQLQuery = "SELECT ID, Data FROM DataStore WHERE json_extract(Data, '$.$type') = '" & typeName & ", " & Assembly.GetExecutingAssembly().GetName().Name & "'"
                Else
                    lsSQLQuery = "SELECT ID, Data FROM DataStore WHERE json_extract(Data, '$.$type') = '" & typeName & ", " & Assembly.GetExecutingAssembly().GetName().Name & "'"
                    lsSQLQuery &= " AND " & Me.GenerateJsonWhereClause(whereClause)
                End If

                Dim lrRecordset As New RecordsetProxy
                lrRecordset.ActiveConnection = pdbConnection
                lrRecordset.CursorType = pcOpenStatic

                Call lrRecordset.Open(lsSQLQuery)

                While Not lrRecordset.EOF
                    Dim jsonData As String = lrRecordset("Data").Value
                    Dim jsonReader As JsonTextReader = New JsonTextReader(New StringReader(jsonData))
                    Dim data As T = JsonSerializer.Create().Deserialize(Of T)(jsonReader)

                    If whereClause Is Nothing OrElse whereClause.Compile()(data) Then
                        dataList.Add(data)
                    End If

                    asID = lrRecordset("ID").value

                    lrRecordset.MoveNext()
                End While

                lrRecordset.Close()

                Return dataList

            Catch ex As Exception
                Dim lsMessage As String
                Dim mb As MethodBase = MethodInfo.GetCurrentMethod()

                lsMessage = "Error: " & mb.ReflectedType.Name & "." & mb.Name
                lsMessage &= vbCrLf & vbCrLf & ex.Message
                prApplication.ThrowErrorMessage(lsMessage, pcenumErrorType.Critical, ex.StackTrace,,)
            End Try

        End Function

        Private Function GenerateJsonWhereClause(Of T)(whereClause As Expression(Of Func(Of T, Boolean))) As String
            Dim jsonWhereBuilder As New StringBuilder()

            Dim conditions = ExtractConditions(whereClause.Body)
            For Each condition In conditions
                If jsonWhereBuilder.Length > 0 Then
                    jsonWhereBuilder.Append(" AND ")
                End If
                AppendCondition(condition, jsonWhereBuilder)
            Next

            Return jsonWhereBuilder.ToString()
        End Function

        Private Function ExtractConditions(expression As Expression) As IEnumerable(Of BinaryExpression)
            Dim conditions As New List(Of BinaryExpression)
            CollectConditions(expression, conditions)
            Return conditions
        End Function

        Private Sub CollectConditions(expression As Expression, conditions As List(Of BinaryExpression))
            If TypeOf expression Is BinaryExpression Then
                Dim binaryExpression = DirectCast(expression, BinaryExpression)
                If binaryExpression.NodeType = ExpressionType.And OrElse binaryExpression.NodeType = ExpressionType.AndAlso Then
                    CollectConditions(binaryExpression.Left, conditions)
                    CollectConditions(binaryExpression.Right, conditions)
                ElseIf binaryExpression.NodeType = ExpressionType.Equal Then
                    conditions.Add(binaryExpression)
                End If
            End If
            ' Add handling for other types of expressions if needed
        End Sub

        Private Sub AppendCondition(binaryExpression As BinaryExpression, jsonWhereBuilder As StringBuilder)
            If binaryExpression.Left.NodeType = ExpressionType.MemberAccess AndAlso binaryExpression.Right.NodeType = ExpressionType.Constant Then
                AppendMemberAccessCondition(binaryExpression, jsonWhereBuilder)
            ElseIf binaryExpression.Left.NodeType = ExpressionType.Call AndAlso binaryExpression.Right.NodeType = ExpressionType.Constant Then
                AppendMethodCallCondition(binaryExpression, jsonWhereBuilder)
            End If
            ' Add handling for other types of expressions if needed
        End Sub

        Private Sub AppendMemberAccessCondition(binaryExpression As BinaryExpression, jsonWhereBuilder As StringBuilder)
            Dim memberExpression = DirectCast(binaryExpression.Left, MemberExpression)
            Dim memberName = memberExpression.Member.Name
            Dim constantValue = DirectCast(DirectCast(binaryExpression.Right, ConstantExpression).Value, String)

            jsonWhereBuilder.Append("json_extract(Data, '$.")
            jsonWhereBuilder.Append(memberName)
            jsonWhereBuilder.Append("') = '")
            jsonWhereBuilder.Append(constantValue)
            jsonWhereBuilder.Append("'")
        End Sub

        Public Function GetNestedPropertyValue(obj As Object, propertyPath As String) As Object
            Dim parts As String() = propertyPath.Split("."c)
            Dim currentObj As Object = obj
            Dim lasParts = parts.ToList
            lasParts.RemoveAt(parts.Count - 1)
            parts = lasParts.ToArray

            For Each part As String In parts
                If currentObj Is Nothing Then Return Nothing
                Dim propInfo As PropertyInfo = currentObj.GetType().GetProperty(part, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)
                If propInfo Is Nothing Then Return currentObj
                currentObj = propInfo.GetValue(currentObj, Nothing)
            Next

            Return currentObj
        End Function

        Public Shared Function GetMemberExpressionValue(expression As MemberExpression) As Object
            Dim dependencyChain As New List(Of MemberExpression)()
            Dim pointingExpression As MemberExpression = expression
            While pointingExpression IsNot Nothing
                dependencyChain.Add(pointingExpression)
                pointingExpression = TryCast(pointingExpression.Expression, MemberExpression)
            End While

            Dim baseExpression As ConstantExpression = TryCast(dependencyChain.Last().Expression, ConstantExpression)
            If baseExpression Is Nothing Then
                Throw New Exception($"Last expression {dependencyChain.Last().Expression} of dependency chain of {expression} is not a constant." & "Thus the expression value cannot be found.")
            End If

            Dim resolvedValue As Object = baseExpression.Value

            For i As Integer = dependencyChain.Count To 1 Step -1
                Dim expr As MemberExpression = dependencyChain(i - 1)
                resolvedValue = New PropOrField(expr.Member).GetValue(resolvedValue)
            Next

            Return resolvedValue
        End Function

        Private Sub AppendMethodCallCondition(binaryExpression As BinaryExpression, jsonWhereBuilder As StringBuilder)

            Try

                Dim methodCallExpression = DirectCast(binaryExpression.Left, MethodCallExpression)
                Dim visitor = New ClosureExtractorVisitor()
                visitor.Visit(methodCallExpression)

                ' Extract the member name and value
                Dim extractedMemberName As String = Nothing
                Dim extractedMemberValue As Object = Nothing
                Dim SecondMemberName As String = "<Error>"
                Dim SecondMemberType As Object = Nothing
                Dim SecondMemberValue As Object = Nothing

                Dim constantValue As String = "<Error>"

                If visitor.ExtractedValues.Count > 0 Then
                    ' Assuming the first extracted value is the one we need
                    Dim firstExtractedValue = visitor.ExtractedValues.First()
                    If TypeOf firstExtractedValue.Key Is MemberExpression Then
                        Dim memberExpr = DirectCast(firstExtractedValue.Key, MemberExpression)
                        extractedMemberName = memberExpr.Member.Name

                        extractedMemberValue = visitor.ExtractedValues.ToList(1).Value

                        SecondMemberType = visitor.ExtractedValues.ToList(1).Key.Type

                        SecondMemberValue = Me.GetMemberExpressionValue(visitor.ExtractedValues.ToList(1).Key)

                    End If

                    If binaryExpression.Right.NodeType = ExpressionType.Constant Then
                        If SecondMemberValue IsNot Nothing Then
                            Select Case SecondMemberValue.GetType
                                Case Is = GetType(System.String),
                                          GetType(System.Int16)
                                    constantValue = SecondMemberValue
                                    GoTo FoundValues
                                Case Else

                            End Select
                        End If
                    End If

                    Dim lsFullSecondMemberPath As String = Nothing
                    Dim input As String = binaryExpression.ToString
                    Dim pattern As String = ".*\.([A-Za-z0-9_]+)" 'Was "\.(?<MemberName>[A-Za-z0-9_]+)(?![^.]*\.[A-Za-z0-9_])"
                    Dim match As Match = Regex.Match(input, pattern)

                    If match.Success Then
                        SecondMemberName = match.Groups(1).Value 'Was "MemberName").Value
                    End If


                    Try
                        lsFullSecondMemberPath = input.Split(",")(2).Split(":")(1)
                        Dim splitParts As String() = lsFullSecondMemberPath.Split("."c)
                        lsFullSecondMemberPath = String.Join(".", splitParts, 1, splitParts.Length - 1)
                    Catch ex As Exception
                        lsFullSecondMemberPath = SecondMemberName
                    End Try



                    Try
                        If SecondMemberName <> lsFullSecondMemberPath And lsFullSecondMemberPath.Split(".").Count > 2 Then
                            SecondMemberName = lsFullSecondMemberPath
                        End If
                    Catch ex As Exception
                        'We tried
                    End Try


                End If

                If extractedMemberName IsNot Nothing Then

                    ' First try to get it as a property
                    Dim memberInfo As MemberInfo = extractedMemberValue.GetType().GetProperty(SecondMemberName, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)

                    ' If not found as a property, try to get it as a field
                    If memberInfo Is Nothing Then
                        memberInfo = extractedMemberValue.GetType().GetField(SecondMemberName, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)
                    End If

                    If memberInfo Is Nothing And SecondMemberName.Contains(".") Then
                        memberInfo = GetNestedPropertyValue(extractedMemberValue, SecondMemberName)
                    End If

                    ' If the member is found (either as a property or field), get its value
                    If memberInfo IsNot Nothing Then
                        Dim value As Object = Nothing
                        If TypeOf memberInfo Is PropertyInfo Then
                            value = DirectCast(memberInfo, PropertyInfo).GetValue(extractedMemberValue, Nothing)
                        ElseIf TypeOf memberInfo Is FieldInfo Then
                            value = DirectCast(memberInfo, FieldInfo).GetValue(extractedMemberValue)
                        End If

                        If value IsNot Nothing Then
                            constantValue = value.ToString()
                        End If
                    Else
                        Try
                            memberInfo = SecondMemberValue.GetType().GetProperty(SecondMemberName, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)

                            If memberInfo Is Nothing Then
                                Try
                                    memberInfo = SecondMemberValue.GetType().GetField(SecondMemberName)
                                Catch ex As Exception
                                    'We tried
                                End Try
                            End If

                            If memberInfo IsNot Nothing Then
                                Dim value As Object = Nothing
                                If TypeOf memberInfo Is PropertyInfo Then
                                    value = DirectCast(memberInfo, PropertyInfo).GetValue(SecondMemberValue, Nothing)
                                ElseIf TypeOf memberInfo Is FieldInfo Then
                                    value = DirectCast(memberInfo, FieldInfo).GetValue(SecondMemberValue)
                                End If

                                If value IsNot Nothing Then
                                    constantValue = value.ToString()
                                End If
                            End If

                        Catch ex As Exception
                            'We tried
                        End Try
                    End If
FoundValues:
                    jsonWhereBuilder.Append("json_extract(Data, '$.")
                    jsonWhereBuilder.Append(extractedMemberName)
                    jsonWhereBuilder.Append("') = '")
                    jsonWhereBuilder.Append(constantValue)
                    jsonWhereBuilder.Append("'")
                Else
                    ' Handle the case where the member name could not be extracted
                    ' This may involve logging an error or throwing an exception
                End If

            Catch ex As Exception
                Dim lsMessage As String
                Dim mb As MethodBase = MethodInfo.GetCurrentMethod()

                lsMessage = "Error: " & mb.ReflectedType.Name & "." & mb.Name
                lsMessage &= vbCrLf & vbCrLf & ex.Message
                prApplication.ThrowErrorMessage(lsMessage, pcenumErrorType.Critical, ex.StackTrace,,)
            End Try
        End Sub

        Private Function ExtractMemberNameFromMethodCall(methodCallExpression As MethodCallExpression) As String
            ' Assuming the method call is structured in a way where one of the arguments
            ' is a MemberExpression which contains the member name we need.
            For Each arg As Expression In methodCallExpression.Arguments
                If TypeOf arg Is MemberExpression Then
                    Dim memberExpr = DirectCast(arg, MemberExpression)
                    Return memberExpr.Member.Name
                End If
                ' Depending on the structure, you might need to check for other types of expressions here
            Next

            ' If the member name isn't found, return a default value or handle accordingly
            Return String.Empty
        End Function

        ''' <summary>
        ''' Update function to replace the JSON in the Data field with Newtonsoft serialization
        ''' NB Only operates where one record is returned from inner Get. I.e. You should aim to update only one Document.
        ''' </summary>
        ''' <typeparam name="t"></typeparam>
        ''' <param name="arObject"></param>
        ''' <param name="whereClause"></param>
        Public Sub Update(Of t)(arObject As Object, whereClause As Expression(Of Func(Of t, Boolean)))
            Try
                'Usage examples
                'Dim whereClause As Expression(Of Func(Of FBM.DictionaryEntry, Boolean)) = Function(p) p.Symbol = "Satellite"
                'Dim larDictionaryEntry As List(Of FBM.DictionaryEntry) = lrDataStore.GetData(Of FBM.DictionaryEntry)(whereClause)
                'or
                'Dim lrEntityType As New FBM.EntityType(Nothing, pcenumLanguage.ORMModel, "Test Entity Type", Nothing, True)
                'Dim whereClause As Expression(Of Func(Of FBM.EntityType, Boolean)) = Function(p) p.Id = "Test Entity Type"
                'lrEntityType.Symbol = "Testaddb"
                'Dim lrDataStore As New DataStore.Store
                'Call lrDataStore.UpdateData(Of FBM.EntityType)(lrEntityType, whereClause)

                ' Get the record that matches the provided whereClause using GetData function
                Dim asID As String = Nothing

                Dim dataList As List(Of t) = Me.Get(Of t)(whereClause, asID) 'Id is byRef and Updated by Get.

                ' Assuming there should be only one record that matches the condition
                If dataList.Count = 1 AndAlso asID IsNot Nothing Then
                    Dim recordToUpdate = dataList(0)

                    ' Create the custom JSON serializer settings
                    Dim settings As New JsonSerializerSettings With {
                                .Formatting = Formatting.Indented,
                                .TypeNameHandling = TypeNameHandling.Objects,
                                .ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                            }

                    ' Serialize the object into JSON with custom settings
                    Dim jsonData As String = JsonConvert.SerializeObject(arObject, settings)

                    ' Now you can update the record back in the database using prConnection.Execute
                    ' Here's the SQL query to update the Data field for the specific record identified by ID
                    Dim lsSQLQuery As String = "UPDATE DataStore SET Data = '" & jsonData & "' WHERE ID = '" & asID & "'"

                    ' Execute the update query
                    Dim lrRecordset As ORMQL.Recordset = pdbConnection.Execute(lsSQLQuery)

                    If lrRecordset.ErrorReturned Then
                        Throw New Exception(lrRecordset.ErrorString)
                    End If

                End If

            Catch ex As Exception
                Dim lsMessage As String
                Dim mb As MethodBase = MethodInfo.GetCurrentMethod()

                lsMessage = "Error: " & mb.ReflectedType.Name & "." & mb.Name
                lsMessage &= vbCrLf & vbCrLf & ex.Message
                prApplication.ThrowErrorMessage(lsMessage, pcenumErrorType.Critical, ex.StackTrace,,)
            End Try
        End Sub

        Public Sub UpsertWrapper(arObject As Object, whereClause As Object)
            Dim objectType = arObject.GetType()
            Dim upsertMethod = Me.GetType().GetMethod("Upsert").MakeGenericMethod(objectType)
            upsertMethod.Invoke(Me, {arObject, whereClause})
        End Sub

        ''' <summary>
        ''' Upsert method to either update or insert a record in the DataStore table
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="arObject"></param>
        ''' <param name="whereClause"></param>
        Public Sub Upsert(Of T)(arObject As Object, whereClause As Expression(Of Func(Of T, Boolean)))
            Try
                ' Get the record that matches the provided whereClause using GetData function
                Dim asID As String = Nothing
                Dim dataList As List(Of T) = Me.Get(Of T)(whereClause, asID)

                ' Serialize the object into JSON with custom settings
                Dim settings As New JsonSerializerSettings With {
                    .Formatting = Formatting.Indented,
                    .TypeNameHandling = TypeNameHandling.Objects,
                    .ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                }
                Dim jsonData As String = JsonConvert.SerializeObject(arObject, settings)
                jsonData = Database.MakeStringSafe(jsonData)

                If dataList.Count >= 1 AndAlso asID IsNot Nothing Then
                    ' If the record exists, update it in the database
                    Dim lsSQLQuery As String = "UPDATE DataStore SET Data = '" & jsonData & "' WHERE ID = '" & asID & "'"
                    Dim lrRecordset As ORMQL.Recordset = pdbConnection.Execute(lsSQLQuery)

                    If lrRecordset.ErrorReturned Then
                        If lrRecordset.ErrorString.Contains("unique") Then
                            'Ignore 
                        Else
                            Throw New Exception(lrRecordset.ErrorString)
                        End If
                    End If
                Else
                    ' If the record doesn't exist, insert it into the database
                    Dim lrData As New DataStore.Data(jsonData, GetType(T).FullName, Now, "", "", "", "", "")
                    Call tableDataStore.AddData(lrData)
                End If

            Catch ex As Exception
                Dim lsMessage As String
                Dim mb As MethodBase = MethodInfo.GetCurrentMethod()

                lsMessage = "Error: " & mb.ReflectedType.Name & "." & mb.Name
                lsMessage &= vbCrLf & vbCrLf & ex.Message
                prApplication.ThrowErrorMessage(lsMessage, pcenumErrorType.Critical, ex.StackTrace,,)
            End Try
        End Sub

    End Class

    ' Use a custom ExpressionVisitor to extract values from closures
    Public Class ClosureExtractorVisitor
        Inherits ExpressionVisitor

        Private _extractedValues As New Dictionary(Of Expression, Object)

        Public ReadOnly Property ExtractedValues As Dictionary(Of Expression, Object)
            Get
                Return _extractedValues
            End Get
        End Property

        Protected Overrides Function VisitMember(node As MemberExpression) As Expression
            If TypeOf node.Expression Is ConstantExpression Then
                ' Extracting the value from a constant object
                Dim constantExpr = DirectCast(node.Expression, ConstantExpression)
                Dim closureInstance = constantExpr.Value
                Dim memberInfo = TryCast(node.Member, PropertyInfo)
                If memberInfo IsNot Nothing Then
                    ' If the member is a property, extract the property value
                    Dim value = memberInfo.GetValue(closureInstance, Nothing)
                    _extractedValues(node) = value
                Else
                    _extractedValues(node) = constantExpr.Value
                    ' If it's not a property, handle accordingly (e.g., fields)
                    ' Implementation depends on your specific scenario
                End If
            ElseIf TypeOf node.Expression Is ParameterExpression Then
                ' Handling parameter expressions (Note: actual value cannot be determined here)
                Dim parameterExpr = DirectCast(node.Expression, ParameterExpression)
                _extractedValues(node) = parameterExpr
            Else
                ' Recursively handle other types of expressions
                Visit(node.Expression)
            End If

            Return MyBase.VisitMember(node)
        End Function

    End Class

    Public Class PropOrField
        Public ReadOnly MemberInfo As MemberInfo

        Public Sub New(memberInfo As MemberInfo)
            If Not (TypeOf memberInfo Is PropertyInfo) AndAlso Not (TypeOf memberInfo Is FieldInfo) Then
                Throw New Exception($"{NameOf(memberInfo)} must either be {NameOf(PropertyInfo)} or {NameOf(FieldInfo)}")
            End If

            Me.MemberInfo = memberInfo
        End Sub

        Public Function GetValue(source As Object) As Object
            If TypeOf MemberInfo Is PropertyInfo Then
                Return CType(MemberInfo, PropertyInfo).GetValue(source)
            ElseIf TypeOf MemberInfo Is FieldInfo Then
                Return CType(MemberInfo, FieldInfo).GetValue(source)
            End If

            Return Nothing
        End Function

        Public Sub SetValue(target As Object, source As Object)
            If TypeOf MemberInfo Is PropertyInfo Then
                CType(MemberInfo, PropertyInfo).SetValue(target, source)
            ElseIf TypeOf MemberInfo Is FieldInfo Then
                CType(MemberInfo, FieldInfo).SetValue(target, source)
            End If
        End Sub

        Public Function GetMemberType() As Type
            If TypeOf MemberInfo Is PropertyInfo Then
                Return CType(MemberInfo, PropertyInfo).PropertyType
            ElseIf TypeOf MemberInfo Is FieldInfo Then
                Return CType(MemberInfo, FieldInfo).FieldType
            End If

            Return Nothing
        End Function
    End Class

End Namespace
