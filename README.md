# DataStore
.Net JSON Database in SQLite, Single Table/Single Field design, POCO Storage

# Connect to SQLite database and create the DataStore table if it does not exist

            Dim databasePath As String = Path.Combine(Environment.CurrentDirectory, "DataStoreConsoleDemo.sqlite")

            Dim store As New Store
            store.Connect(databasePath, 3)

===
Single Table structure:

	CREATE TABLE DataStore (
		ID TEXT PRIMARY KEY,
		Data TEXT,
		DataClass TEXT,
		EntryDate DATETIME,
		Source TEXT,
		Author TEXT,
		Tags TEXT,
		Location TEXT,
		Notes TEXT,
		UNIQUE (Data) ON CONFLICT FAIL)

	NB Predominantly used are the ID and Data fields. Data field stores JSON objects (Newtonsoft.Json serialised POCO objects).

# Features

DataStore is a lightweight .NET document database built on SQLite: one table, one JSON document column, and strongly typed POCOs at the application boundary. It keeps the deployment simplicity of a local SQLite file while providing a surprisingly expressive, LINQ-shaped document API.

## Document-first storage, SQLite-simple deployment

- **Single-table / single-document design.** All application objects live in the `DataStore` table, with the JSON payload stored in `Data` and supporting metadata available for ID, class, entry date, source, author, tags, location, and notes.
- **POCO-native persistence.** Store ordinary .NET classes directly; DataStore serializes and deserializes them with Newtonsoft.Json.
- **Runtime type preservation.** JSON is written with `TypeNameHandling.Objects`, allowing a heterogeneous document collection to retain concrete .NET type information.
- **Safe object graphs.** Serialization is configured to ignore reference loops, making it practical to persist real-world domain models without hand-flattening every relationship.
- **Zero-server operation.** Connect to an existing SQLite connection string, or point DataStore at a file and have it create the backing table on first use.
- **Duplicate-document protection.** The backing table includes a uniqueness constraint on the JSON `Data` payload.

## Typed document lifecycle

- **Add documents** with `Add(object)` or insert a prepared `DataStore.Data` record when lower-level control is needed.
- **Retrieve typed collections** with `Get(Of T)(...)`; the store selects the correct document type and materializes results back into `T`.
- **Update matching documents** by replacing their serialized JSON payload.
- **Upsert in one call.** `Upsert(Of T)` finds a document using a typed predicate, then updates it or inserts it when no match exists.
- **Delete by expression.** `Delete(Of T)` removes documents selected by the same strongly typed predicate style used for reads and updates.

## LINQ-shaped JSON querying

- **Write predicates against your POCO, not JSON strings.** A lambda such as `Function(p) p.UserId = userId` is translated into SQLite `json_extract` filtering and also evaluated against deserialized objects for faithful result filtering.
- **Predicate pushdown.** Type filtering and recognised conditions are applied in SQLite before objects are materialized, keeping ordinary lookups lean.
- **Comparison support.** Equality, inequality, greater/less-than comparisons, null checks, booleans, enums, numeric values, dates, GUIDs, and string values are represented as SQLite-safe literals.
- **Friendly string operations.** `Equals`, `Contains`, `StartsWith`, and `EndsWith` are translated into JSON-field comparisons / `LIKE` conditions.
- **Captured values work naturally.** Values held in surrounding variables or member expressions can be evaluated from the expression tree, so query code stays readable.

## Multi-collection WHERE joins - virtual tables without schema joins

The standout feature is an N-ary `Get` overload that accepts a multi-parameter `LambdaExpression`. Each lambda parameter represents a typed virtual document collection. DataStore fetches the candidate documents for each collection, forms tuples, applies the complete predicate in-process, and returns the matching tuples as `List(Of Object())`.

This lets one `WHERE` span multiple unrelated document types, without creating physical tables or writing SQL joins:

```vb
Dim whereAll As Expression(Of Func(Of TestPlan, TestSet, TestCycle, TestCase, Boolean)) =
    Function(plan, testSet, cycle, testCase) _
        plan.Identifier = testRun.TestPlanIdentifier AndAlso
        testSet.Identifier = testRun.TestSetIdentifier AndAlso
        cycle.Identifier = testRun.TestCycleIdentifier AndAlso
        testCase.Identifier = testRun.TestCaseIdentifier

Dim tuples As List(Of Object()) = store.Get(CType(whereAll, LambdaExpression))
Dim match = tuples.Single()

Dim plan = DirectCast(match(0), TestPlan)
Dim testSet = DirectCast(match(1), TestSet)
Dim cycle = DirectCast(match(2), TestCycle)
Dim testCase = DirectCast(match(3), TestCase)
```

Under the hood, DataStore harvests simple equality conditions for each lambda parameter and pushes those per-type constraints into the individual document reads where possible. It then evaluates the original full predicate over the candidate tuples. The result is expressive application-level joins with typed inputs, predictable tuple ordering, and no relational schema ceremony.

## Relationship-aware deletion

- **Declarative document relationships.** Decorate a property or field with `ForeignKeyReferenceAttribute` to describe its target type and target member.
- **Reflection-built dependency index.** DataStore scans configured application assembly prefixes once and caches a target-type-to-dependent map using thread-safe lazy initialization.
- **Dependent cleanup on delete.** When a root document is deleted, matching dependent documents are located and deleted first, enabling cascading document cleanup across POCO types.
- **Properties and fields supported.** Relationship and value access work with either public instance properties or fields.

## Built for pragmatic .NET applications

- **No ORM mapping layer to maintain.** Your POCOs remain the model; JSON is the storage format.
- **SQLite JSON functions as the query engine.** Keep a compact database file while retaining filterable fields inside each document.
- **Reflection where it adds leverage.** Generic APIs, runtime type construction, and expression trees let one store service work across the domain model.
- **Useful escape hatch.** The `DataStore.Data` path exposes the backing record when document-level metadata or direct record handling is required.

> **Best fit:** rich, evolving application models where simplicity of storage and type-safe document access are more valuable than a fully normalized relational schema. The multi-collection query feature is especially useful for composing related views across independently stored document collections.

# Upserting Objects (Upsert)

            Dim lrDataStore As New DataStore.Store
            Dim whereClause As Expression(Of Func(Of Personalisation.Profile, Boolean)) = Function(t) t.UserId = Me.zrUser.Id

            Dim lrProfile As New Personalisation.Profile

            lrProfile.UserId = Me.zrUser.Id
            lrProfile.LogoFileLocation = Me.TextBoxLogoFileLocation.Text.Trim
            lrProfile.ShowEnterpriseExplorer = Me.CheckBoxShowEnterpriseExplorer.Checked
            lrProfile.ShowVirtualAnalyst = Me.CheckBoxShowVirtualAnalyst.Checked
            lrProfile.ShowTheBox = Me.CheckBoxShowTheBox.Checked

            Try
                lrProfile.DefaultOSMTaskId = Me.ComboBoxDefaultTask.SelectedItem.ItemData
            Catch ex As Exception
                lrProfile.DefaultOSMTaskId = ""
            End Try

            Call lrDataStore.Upsert(lrProfile, whereClause)

# Inserting / Adding Objecting (Add)

            Dim lrDocument As New VectorDB.Document
            lrDocument.DocumentFileLocation = asDocumentPath
            lrDataStore.Add(lrDocument)

# Getting / Retrieving Objects

            Dim lrDataStore As New DataStore.Store
            Dim whereClause As Expression(Of Func(Of Personalisation.Profile, Boolean)) = Function(t) t.UserId = lrUser.Id

            Dim larProfile = lrDataStore.Get(whereClause)

# Deleting Objects

            Dim whereClause As Expression(Of Func(Of FBM.EntityType, Boolean)) = Function(p) p.Id = "Test Entity Type"

            Dim lrDataStore As New DataStore.Store
            Call lrDataStore.Delete(Of FBM.EntityType)(whereClause)

# Updating Objects

            Dim whereClause As Expression(Of Func(Of FBM.EntityType, Boolean)) = Function(p) p.Id = "Test Entity Type"

            lrEntityType.Symbol = "Testaddb"

            Dim lrDataStore As New DataStore.Store
            Call lrDataStore.Update(Of FBM.EntityType)(lrEntityType, whereClause)


