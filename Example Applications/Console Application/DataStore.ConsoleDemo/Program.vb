Imports DataStore.DataStore
Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq.Expressions

Namespace ConsoleDemo

    Module Program

        Private ReadOnly SampleOrderPrefix As String = "PO-2024-"
        Private ReadOnly SampleCustomerPrefix As String = "CUST-20"

        Sub Main()
            Dim databasePath As String = Path.Combine(Environment.CurrentDirectory, "DataStoreConsoleDemo.sqlite")
            Console.WriteLine($"Connecting to demo database at {databasePath} ...")

            Dim store As New Store
            store.Connect(databasePath, 3)

            SeedSampleData(store)
            DisplayStoredCustomers(store)
            DisplayOpenOrders(store)
            'ShowJoinedView(store)
            'UpdateCustomerEmail(store)
            'ArchiveOldOrders(store)

            Console.WriteLine()
            Console.WriteLine("Demo complete. You can inspect the SQLite file to explore the stored JSON payloads.")
            Console.WriteLine("Press [Enter] to exit...")
            Console.ReadLine()
        End Sub

        Private Sub SeedSampleData(store As Store)
            Console.WriteLine()
            Console.WriteLine("Seeding customers and purchase orders ...")

            ' Clean up anything from previous runs so the demo is deterministic.
            store.Delete(Of ExampleModels.PurchaseOrder)(Function(po) po.OrderNumber.StartsWith(SampleOrderPrefix))
            store.Delete(Of ExampleModels.Customer)(Function(c) c.CustomerNumber.StartsWith(SampleCustomerPrefix))

            Dim customers As List(Of ExampleModels.Customer) = SampleData.CreateCustomers()
            For Each customer In customers
                store.Upsert(Of ExampleModels.Customer)(customer, Function(c) c.CustomerNumber = customer.CustomerNumber)
            Next

            Dim orders As List(Of ExampleModels.PurchaseOrder) = SampleData.CreateOrders(customers)
            For Each order In orders
                store.Upsert(Of ExampleModels.PurchaseOrder)(order, Function(po) po.OrderNumber = order.OrderNumber)
            Next

            Console.WriteLine($"Seeded {customers.Count} customers and {orders.Count} orders.")
        End Sub

        Private Sub DisplayStoredCustomers(store As Store)
            Console.WriteLine()
            Console.WriteLine("Customers currently on file:")

            Dim customers As List(Of ExampleModels.Customer) = store.Get(Of ExampleModels.Customer)()
            For Each customer In customers
                Console.WriteLine($" - {customer.CustomerNumber}: {customer.Name} ({customer.Email})")
                If customer.Addresses IsNot Nothing Then
                    For Each address In customer.Addresses
                        Console.WriteLine($"     {address.Label}: {address.Street}, {address.City}, {address.State} {address.PostalCode}")
                    Next
                End If
            Next
        End Sub

        Private Sub DisplayOpenOrders(store As Store)
            Console.WriteLine()
            Console.WriteLine("Open purchase orders:")

            Dim openOrders As List(Of ExampleModels.PurchaseOrder) = store.Get(Of ExampleModels.PurchaseOrder)(Function(po) po.Status = ExampleModels.OrderStatus.Open)
            For Each order In openOrders
                Console.WriteLine($" - {order.OrderNumber} for customer {order.CustomerNumber}: {order.Total:c}")
                For Each line In order.Items
                    Console.WriteLine($"     {line.Quantity} x {line.Description} @ {line.UnitPrice:c}")
                Next
            Next
        End Sub

        Private Sub ShowJoinedView(store As Store)
            Console.WriteLine()
            Console.WriteLine("Joining customers with their recent orders:")

            Dim cutoff As DateTime = DateTime.Today.AddDays(-30)

            Dim lWhereAll As Expression(Of Func(Of ExampleModels.Customer, ExampleModels.PurchaseOrder, Boolean)) =
                Function(lcust, lpo) _
                    lcust.CustomerNumber = lpo.CustomerNumber AndAlso
                    lpo.PlacedOn >= cutoff


            Dim joined As List(Of Object()) = store.Get(lWhereAll)

            For Each tuple In joined
                Dim customer As ExampleModels.Customer = DirectCast(tuple(0), ExampleModels.Customer)
                Dim order As ExampleModels.PurchaseOrder = DirectCast(tuple(1), ExampleModels.PurchaseOrder)
                Console.WriteLine($" - {customer.Name} placed order {order.OrderNumber} on {order.PlacedOn:d} totaling {order.Total:c}")
            Next
        End Sub

        Private Sub UpdateCustomerEmail(store As Store)
            Console.WriteLine()
            Console.WriteLine("Updating a customer's preferred email address via Upsert ...")

            Dim targetId As String = SampleCustomerPrefix & "01"
            Dim customer As ExampleModels.Customer = store.Get(Of ExampleModels.Customer)(Function(c) c.CustomerNumber = targetId).FirstOrDefault()

            If customer Is Nothing Then
                Console.WriteLine("   Customer not found; skipping update.")
                Return
            End If

            customer.Email = "alice.smith+demo@example.com"
            store.Upsert(Of ExampleModels.Customer)(customer, Function(c) c.CustomerNumber = targetId)

            Dim updated As ExampleModels.Customer = store.Get(Of ExampleModels.Customer)(Function(c) c.CustomerNumber = targetId).First()
            Console.WriteLine($"   Updated email stored as {updated.Email}.")
        End Sub

        Private Sub ArchiveOldOrders(store As Store)
            Console.WriteLine()
            Console.WriteLine("Archiving completed orders older than 90 days ...")

            Dim cutoff As DateTime = DateTime.Today.AddDays(-90)
            Dim archivedBefore As Integer = store.Get(Of ExampleModels.PurchaseOrder)(Function(po) po.Status = ExampleModels.OrderStatus.Completed AndAlso po.PlacedOn < cutoff).Count

            store.Delete(Of ExampleModels.PurchaseOrder)(Function(po) po.Status = ExampleModels.OrderStatus.Completed AndAlso po.PlacedOn < cutoff)

            Dim archivedAfter As Integer = store.Get(Of ExampleModels.PurchaseOrder)(Function(po) po.Status = ExampleModels.OrderStatus.Completed AndAlso po.PlacedOn < cutoff).Count

            Console.WriteLine($"   Removed {archivedBefore - archivedAfter} historical orders.")
        End Sub

        <System.Runtime.CompilerServices.Extension>
        Private Function FirstOrDefault(Of T)(items As IList(Of T)) As T
            If items Is Nothing OrElse items.Count = 0 Then
                Return Nothing
            End If
            Return items(0)
        End Function

        <System.Runtime.CompilerServices.Extension>
        Private Function First(Of T)(items As IList(Of T)) As T
            If items Is Nothing OrElse items.Count = 0 Then
                Throw New InvalidOperationException("Sequence contains no elements.")
            End If
            Return items(0)
        End Function

    End Module

End Namespace
