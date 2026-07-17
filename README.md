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


