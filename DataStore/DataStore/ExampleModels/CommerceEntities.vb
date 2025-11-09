Imports System
Imports System.Collections.Generic

Namespace DataStore.ExampleModels

    Public Enum OrderStatus
        Draft
        Open
        Completed
        Cancelled
    End Enum

    ''' <summary>
    ''' Simple customer document used by the console demo. It is compiled into the DataStore assembly so queries can
    ''' filter on <c>$type</c> metadata without any additional configuration.
    ''' </summary>
    Public Class Customer

        <PrimaryKeyField>
        Public Property CustomerNumber As String
        Public Property Name As String
        Public Property Email As String
        Public Property RegisteredOn As DateTime
        Public Property Tags As List(Of String)
        Public Property Addresses As List(Of Address)

    End Class

    Public Class Address
        Public Property Label As String
        Public Property Street As String
        Public Property City As String
        Public Property State As String
        Public Property PostalCode As String
    End Class

    Public Class PurchaseOrder

        <PrimaryKeyField>
        Public Property OrderNumber As String
        Public Property CustomerNumber As String
        Public Property PlacedOn As DateTime
        Public Property Status As OrderStatus
        Public Property Items As List(Of OrderLine)
        Public Property Timeline As List(Of TimelineNote)

        Public ReadOnly Property Total As Decimal
            Get
                If Items Is Nothing Then Return 0D
                Dim accumulator As Decimal = 0D
                For Each line In Items
                    accumulator += line.Quantity * line.UnitPrice
                Next
                Return accumulator
            End Get
        End Property

    End Class

    Public Class OrderLine
        Public Property Sku As String
        Public Property Description As String
        Public Property Quantity As Integer
        Public Property UnitPrice As Decimal
    End Class

    Public Class TimelineNote
        Public Property OccurredAt As DateTime
        Public Property Author As String
        Public Property Comment As String
    End Class

End Namespace
