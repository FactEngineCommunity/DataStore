Imports System.Linq.Expressions
Imports System.Reflection

Namespace DataStore

    <System.AttributeUsage(System.AttributeTargets.Property)>
    Public Class PrimaryKeyFieldAttribute
        Inherits Attribute

        ' Constructor
        Public Sub New()
            ' Add any additional properties or methods if needed
        End Sub
    End Class

    Public Class DataStorePrototype
        Inherits DataStore.Data

        ''' <summary>
        ''' Upserts the object in the DataStore
        ''' </summary>
        Public Sub UpSert()
            Dim lrDataStore As New DataStore.Store
            Dim whereClause = BuildWhereClause(Me)
            lrDataStore.UpsertWrapper(Me, whereClause)
        End Sub


        Public Sub AutoCopyProperties(sourceObject As Object, ByRef destinationObject As Object)
            Dim type As Type = sourceObject.GetType()
            If type = destinationObject.GetType() Then
                For Each propInfo As PropertyInfo In type.GetProperties()
                    If propInfo.CanRead AndAlso propInfo.CanWrite Then
                        Dim value As Object = propInfo.GetValue(sourceObject, Nothing)
                        propInfo.SetValue(destinationObject, value, Nothing)
                    End If
                Next
            Else
                Throw New ArgumentException("Objects must be of the same type")
            End If
        End Sub

    End Class


    Module LinqExpressions

        Public Function BuildWhereClause(obj As Object) As LambdaExpression
            Dim objectType = obj.GetType()
            Dim parameter = Expression.Parameter(objectType, "data")
            Dim body As Expression = Expression.Constant(True) ' Initialize with True

            ' Reflectively find properties with the PrimaryKeyFieldAttribute and build the expression
            For Each prop In objectType.GetProperties()
                If Attribute.IsDefined(prop, GetType(PrimaryKeyFieldAttribute)) Then
                    ' Create an equality expression for the property
                    Dim propAccess = Expression.Property(parameter, prop.Name)
                    Dim valueAccess = Expression.Constant(prop.GetValue(obj), prop.PropertyType)
                    Dim equals = Expression.Equal(propAccess, valueAccess)

                    ' Combine with current expression body using AndAlso
                    body = Expression.AndAlso(body, equals)
                End If
            Next

            ' Return lambda expression
            Return Expression.Lambda(body, parameter)
        End Function

    End Module

End Namespace
