Imports System
Imports System.Collections.Generic

Namespace ConsoleDemo

    Friend Module SampleData

        Friend Function CreateCustomers() As List(Of ExampleModels.Customer)
            Dim customers As New List(Of ExampleModels.Customer) From {
                New ExampleModels.Customer With {
                    .CustomerNumber = "CUST-2001",
                    .Name = "Alice Smith",
                    .Email = "alice.smith@example.com",
                    .RegisteredOn = DateTime.Today.AddDays(-120),
                    .Tags = New List(Of String) From {"retail", "vip"},
                    .Addresses = New List(Of ExampleModels.Address) From {
                        New ExampleModels.Address With {
                            .Label = "Head Office",
                            .Street = "100 Main Street",
                            .City = "Springfield",
                            .State = "IL",
                            .PostalCode = "62701"
                        }
                    }
                },
                New ExampleModels.Customer With {
                    .CustomerNumber = "CUST-2002",
                    .Name = "Bluebird Bikes",
                    .Email = "orders@bluebirdbikes.example",
                    .RegisteredOn = DateTime.Today.AddDays(-45),
                    .Tags = New List(Of String) From {"wholesale"},
                    .Addresses = New List(Of ExampleModels.Address) From {
                        New ExampleModels.Address With {
                            .Label = "Warehouse",
                            .Street = "2500 Industrial Way",
                            .City = "Madison",
                            .State = "WI",
                            .PostalCode = "53703"
                        }
                    }
                }
            }

            Return customers
        End Function

        Friend Function CreateOrders(customers As IEnumerable(Of ExampleModels.Customer)) As List(Of ExampleModels.PurchaseOrder)
            Dim orders As New List(Of ExampleModels.PurchaseOrder)

            Dim alice As ExampleModels.Customer = customers.FirstOrDefault(Function(c) c.CustomerNumber = "CUST-2001")
            If alice IsNot Nothing Then
                orders.Add(New ExampleModels.PurchaseOrder With {
                    .OrderNumber = "PO-2024-0001",
                    .CustomerNumber = alice.CustomerNumber,
                    .PlacedOn = DateTime.Today.AddDays(-5),
                    .Status = ExampleModels.OrderStatus.Open,
                    .Items = New List(Of ExampleModels.OrderLine) From {
                        New ExampleModels.OrderLine With {.Sku = "LAPTOP-15", .Description = "15-Ultrabook", .Quantity = 5, .UnitPrice = 899D
                },
                        New ExampleModels.OrderLine With {.Sku = "DOCK-USB", .Description = "USB-C Dock", .Quantity = 5, .UnitPrice = 129D
                }
                    },
                    .Timeline = New List(Of ExampleModels.TimelineNote) From {
                        New ExampleModels.TimelineNote With {.OccurredAt = DateTime.Today.AddDays(-5), .Author = "alice.smith@example.com", .Comment = "Submitted purchase order."},
                        New ExampleModels.TimelineNote With {.OccurredAt = DateTime.Today.AddDays(-4), .Author = "logistics@example.com", .Comment = "Confirmed inventory availability."}
                    }
                })

                orders.Add(New ExampleModels.PurchaseOrder With {
                    .OrderNumber = "PO-2023-0999",
                    .CustomerNumber = alice.CustomerNumber,
                    .PlacedOn = DateTime.Today.AddDays(-140),
                    .Status = ExampleModels.OrderStatus.Completed,
                    .Items = New List(Of ExampleModels.OrderLine) From {
                        New ExampleModels.OrderLine With {.Sku = "MONITOR-27", .Description = "27-4K Monitor", .Quantity = 10, .UnitPrice = 449D
                }
                    },
                    .Timeline = New List(Of ExampleModels.TimelineNote) From {
                        New ExampleModels.TimelineNote With {.OccurredAt = DateTime.Today.AddDays(-140), .Author = "alice.smith@example.com", .Comment = "Requested rush delivery."},
                        New ExampleModels.TimelineNote With {.OccurredAt = DateTime.Today.AddDays(-135), .Author = "logistics@example.com", .Comment = "Shipped via priority carrier."}
                    }
                })
            End If

            Dim bluebird As ExampleModels.Customer = customers.FirstOrDefault(Function(c) c.CustomerNumber = "CUST-2002")
            If bluebird IsNot Nothing Then
                orders.Add(New ExampleModels.PurchaseOrder With {
                    .OrderNumber = "PO-2024-0002",
                    .CustomerNumber = bluebird.CustomerNumber,
                    .PlacedOn = DateTime.Today.AddDays(-12),
                    .Status = ExampleModels.OrderStatus.Completed,
                    .Items = New List(Of ExampleModels.OrderLine) From {
                        New ExampleModels.OrderLine With {.Sku = "BIKE-FRAME", .Description = "Carbon Frame", .Quantity = 8, .UnitPrice = 499D
                },
                        New ExampleModels.OrderLine With {.Sku = "WHEEL-700C", .Description = "700C Wheelset", .Quantity = 8, .UnitPrice = 259D
                }
                    },
                    .Timeline = New List(Of ExampleModels.TimelineNote) From {
                        New ExampleModels.TimelineNote With {.OccurredAt = DateTime.Today.AddDays(-12), .Author = "orders@bluebirdbikes.example", .Comment = "Placed order."},
                        New ExampleModels.TimelineNote With {.OccurredAt = DateTime.Today.AddDays(-10), .Author = "logistics@example.com", .Comment = "Order fulfilled."}
                    }
                })
            End If

            Return orders
        End Function

        <System.Runtime.CompilerServices.Extension>
        Private Function FirstOrDefault(Of T)(source As IEnumerable(Of T), predicate As Func(Of T, Boolean)) As T
            For Each item In source
                If predicate(item) Then
                    Return item
                End If
            Next

            Return Nothing
        End Function

    End Module

End Namespace
