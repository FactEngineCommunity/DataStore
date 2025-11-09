Imports DataStore
Imports DataStore.ExampleModels
Imports System
Imports System.Collections.Generic
Imports System.IO

Namespace DataStore.ConsoleDemo

    Module Program

        Private ReadOnly SampleOrderPrefix As String = "PO-2024-"
        Private ReadOnly SampleCustomerPrefix As String = "CUST-20"

        Sub Main()
            Dim databasePath As String = Path.Combine(Environment.CurrentDirectory, "DataStoreConsoleDemo.sqlite")
            Console.WriteLine($"Connecting to demo database at {databasePath} ...")

            Dim store As New Store()
            store.Connect(databasePath, 3)

            SeedSampleData(store)
            DisplayStoredCustomers(store)
            DisplayOpenOrders(store)
            ShowJoinedView(store)
            UpdateCustomerEmail(store)
            ArchiveOldOrders(store)

            Console.WriteLine()
            Console.WriteLine("Demo complete. You can inspect the SQLite file to explore the stored JSON payloads.")
        End Sub

        Private Sub SeedSampleData(store As Store)
            Console.WriteLine()
            Console.WriteLine("Seeding customers and purchase orders ...")

            ' Clean up anything from previous runs so the demo is deterministic.
            store.Delete(Of PurchaseOrder)(Function(po) po.OrderNumber.StartsWith(SampleOrderPrefix))
            store.Delete(Of Customer)(Function(c) c.CustomerNumber.StartsWith(SampleCustomerPrefix))

            Dim customers As List(Of Customer) = SampleData.CreateCustomers()
            For Each customer In customers
                store.Upsert(Of Customer)(customer, Function(c) c.CustomerNumber = customer.CustomerNumber)
            Next

            Dim orders As List(Of PurchaseOrder) = SampleData.CreateOrders(customers)
            For Each order In orders
                store.Upsert(Of PurchaseOrder)(order, Function(po) po.OrderNumber = order.OrderNumber)
            Next

            Console.WriteLine($"Seeded {customers.Count} customers and {orders.Count} orders.")
        End Sub

        Private Sub DisplayStoredCustomers(store As Store)
            Console.WriteLine()
            Console.WriteLine("Customers currently on file:")

            Dim customers As List(Of Customer) = store.Get(Of Customer)()
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

            Dim openOrders As List(Of PurchaseOrder) = store.Get(Of PurchaseOrder)(Function(po) po.Status = OrderStatus.Open)
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
            Dim joined As List(Of Object()) = store.Get(Function(c As Customer, po As PurchaseOrder) c.CustomerNumber = po.CustomerNumber AndAlso po.PlacedOn >= cutoff)

            For Each tuple In joined
                Dim customer As Customer = DirectCast(tuple(0), Customer)
                Dim order As PurchaseOrder = DirectCast(tuple(1), PurchaseOrder)
                Console.WriteLine($" - {customer.Name} placed order {order.OrderNumber} on {order.PlacedOn:d} totaling {order.Total:c}")
            Next
        End Sub

        Private Sub UpdateCustomerEmail(store As Store)
            Console.WriteLine()
            Console.WriteLine("Updating a customer's preferred email address via Upsert ...")

            Dim targetId As String = SampleCustomerPrefix & "01"
            Dim customer As Customer = store.Get(Of Customer)(Function(c) c.CustomerNumber = targetId).FirstOrDefault()

            If customer Is Nothing Then
                Console.WriteLine("   Customer not found; skipping update.")
                Return
            End If

            customer.Email = "alice.smith+demo@example.com"
            store.Upsert(Of Customer)(customer, Function(c) c.CustomerNumber = targetId)

            Dim updated As Customer = store.Get(Of Customer)(Function(c) c.CustomerNumber = targetId).First()
            Console.WriteLine($"   Updated email stored as {updated.Email}.")
        End Sub

        Private Sub ArchiveOldOrders(store As Store)
            Console.WriteLine()
            Console.WriteLine("Archiving completed orders older than 90 days ...")

            Dim cutoff As DateTime = DateTime.Today.AddDays(-90)
            Dim archivedBefore As Integer = store.Get(Of PurchaseOrder)(Function(po) po.Status = OrderStatus.Completed AndAlso po.PlacedOn < cutoff).Count

            store.Delete(Of PurchaseOrder)(Function(po) po.Status = OrderStatus.Completed AndAlso po.PlacedOn < cutoff)

            Dim archivedAfter As Integer = store.Get(Of PurchaseOrder)(Function(po) po.Status = OrderStatus.Completed AndAlso po.PlacedOn < cutoff).Count

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
