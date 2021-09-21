using System.IO;
using IntegratedDriver;
using NUnit.Framework;
using UITests.PageModel;
using UITests.PageModel.Passport;
using UITests.PageModel.Shared;
using static UITests.Constants;
using static UITests.TestHelpers;

namespace UITests.SmokeTesting.GA
{
    [TestFixture]
    public class GlobalDocumentsSmokeTestsWithPersistenceEnabled : UITestBase
    {
        private Outlook _outlook;
        private MattersListPage _mattersListPage;
        private PassportPreferencesPage _passportPreferencesPage;

        [SetUp]
        public void SetUp()
        {
            _outlook = new Outlook(TestEnvironment);
            _outlook.Launch();
            _outlook.Oc.BasicSettingsPage.LogIn();

            _mattersListPage = _outlook.Oc.MattersListPage;
            _mattersListPage.ItemList.AddMatter();
            _passportPreferencesPage = _outlook.Oc.PassportPreferencesPage;
            _passportPreferencesPage.SetPersistentListPagesTo(true);
            _passportPreferencesPage.CloseWindowHandleSwitchToOc();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(GlobalDocumentsTestCategory)]
        [Category(DataDependentTestCategory)]
        [Description("Test case reference: Filters & Smoke Test Views")]
        public void RecentDocumentsFilters()
        {
            const string recentDocTab = "Recent Documents";

            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            globalDocumentsPage.Open();
            var selectedTab = globalDocumentsPage.Tabs.GetActiveTab();
            Assert.That(selectedTab, Is.EqualTo(recentDocTab), $"Default selected tab should be {recentDocTab} but is {selectedTab}");

            var documentSummary = _outlook.Oc.DocumentSummaryPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var recentDocList = globalDocumentsPage.ItemList;

            var unfilteredCount = globalDocumentsPage.ItemList.GetCount();
            //Prepare test documents in case no documents are present
            if (unfilteredCount == 0)
            {
                var matterDocList = _outlook.Oc.DocumentsListPage;
                _mattersListPage.Open();
                _mattersListPage.ItemList.OpenFirst();
                matterDetails.Tabs.Open("Documents");

                var testDoc1 = CreateDocument(OfficeApp.Notepad);
                var testDoc1NameWithoutExtension = Path.GetFileNameWithoutExtension(testDoc1.Name);
                DragAndDrop.FromFileSystem(testDoc1, matterDocList.DropPoint.GetElement());
                matterDocList.AddDocumentDialog.UploadDocument();
                unfilteredCount++;
                var matterDocument1 = matterDetails.ItemList.GetMatterDocumentListItemFromText(testDoc1.Name);
                matterDocument1.FileOptions.CheckOut();

                var notepad = new Notepad(testDoc1NameWithoutExtension);
                notepad.Close();

                var testDoc2 = CreateDocument(OfficeApp.Notepad);
                var testDoc2NameWithoutExtension = Path.GetFileNameWithoutExtension(testDoc2.Name);
                DragAndDrop.FromFileSystem(testDoc2, matterDocList.DropPoint.GetElement());
                matterDocList.AddDocumentDialog.UploadDocument();
                unfilteredCount++;
                var matterDocument2 = matterDetails.ItemList.GetMatterDocumentListItemFromText(testDoc2.Name);
                matterDocument2.FileOptions.CheckOut();

                notepad = new Notepad(testDoc2NameWithoutExtension);
                notepad.Close();

                globalDocumentsPage.Open();
                globalDocumentsPage.OpenRecentDocumentsList();
            }

            var randomDoc = GetRandomNumber(unfilteredCount - 1);
            var selectedDocument = recentDocList.GetGlobalDocumentListItemByIndex(randomDoc);

            //Fetch Document Properties
            var name = GetRandomSubstring(selectedDocument.Name);
            var status = selectedDocument.Status;
            var updatedBy = selectedDocument.CreatedByFullName;
            var matterName = GetRandomSubstring(selectedDocument.AssociatedEntityName.Replace("Matter: ", string.Empty));
            var dateUpdated = selectedDocument.UpdatedAt.Date;

            //Apply Filter - Name
            recentDocList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            var documentsFilterDialog = globalDocumentsPage.GlobalDocumentsListFilterDialog;
            documentsFilterDialog.Controls["Name"].Set(name);
            documentsFilterDialog.Apply();
            Assert.That(recentDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(recentDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(recentDocList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            var filteredList = recentDocList.GetAllGlobalDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by FileName - {name}");
            Assert.That(filteredList, Has.All.Property(nameof(GlobalDocumentListItem.Name)).Contains(name).IgnoreCase, $"Filtered list has items not containing FileName - {name}");

            // Save View
            recentDocList.OpenListOptionsMenu().SaveCurrentView();
            globalDocumentsPage.SaveCurrentViewDialog.Controls["Create New"].Set(ViewName);
            globalDocumentsPage.SaveCurrentViewDialog.Save();

            // Clear Filters
            recentDocList.OpenListOptionsMenu().RestoreDefaults();
            Assert.That(recentDocList.IsFilterIconVisible, Is.False, "Filter Icon is visible");

            // Apply Saved View
            recentDocList.OpenListOptionsMenu().ApplySavedView(ViewName);
            Assert.AreEqual(filteredList.Count, recentDocList.GetCount());

            // Set view as default
            recentDocList.OpenListOptionsMenu().SetCurrentViewAsDefault();

            // Verify current view is set as default
            recentDocList.OpenListOptionsMenu().RestoreDefaults();
            Assert.That(recentDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(recentDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");

            // Clear user default
            recentDocList.OpenListOptionsMenu().ClearUserDefault();
            Assert.That(recentDocList.IsFilterIconVisible, Is.False, "Filter Icon is visible");

            //Apply Filter - Status
            recentDocList.OpenListOptionsMenu().RestoreDefaults();
            recentDocList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Status"].Set(status);
            documentsFilterDialog.Apply();
            Assert.That(recentDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(recentDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(recentDocList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            filteredList = recentDocList.GetAllGlobalDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by Status - {status}");
            Assert.That(filteredList, Has.All.Property(nameof(GlobalDocumentListItem.Status)).EqualTo(status));

            //Apply Filter - Updated By
            recentDocList.OpenListOptionsMenu().RestoreDefaults();
            recentDocList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Updated By"].Set(updatedBy);
            documentsFilterDialog.Apply();
            Assert.That(recentDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(recentDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(recentDocList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            filteredList = recentDocList.GetAllGlobalDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by Updated By - {updatedBy}");
            Assert.That(filteredList, Has.All.Property(nameof(GlobalDocumentListItem.CreatedByFullName)).EqualTo(updatedBy), $"Filtered list has items with update by user other than- {updatedBy}");

            //Apply Filter - Matter Name
            recentDocList.OpenListOptionsMenu().RestoreDefaults();
            recentDocList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Matter Name"].Set(matterName);
            documentsFilterDialog.Apply();
            Assert.That(recentDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(recentDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(recentDocList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            filteredList = recentDocList.GetAllGlobalDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by Matter Name : {matterName}");
            Assert.That(filteredList, Has.All.Property(nameof(GlobalDocumentListItem.AssociatedEntityName)).Contains(matterName).IgnoreCase, $"Filtered list has items not attached with matter name - {matterName}");

            //Apply Filter - Updated At
            recentDocList.OpenListOptionsMenu().RestoreDefaults();
            var dateTo = dateUpdated.AddDays(3);
            recentDocList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Updated At"].Set(FormatDateRange(dateUpdated, dateTo));
            documentsFilterDialog.Apply();
            Assert.That(recentDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(recentDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(recentDocList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            var endDate = dateTo.AddSeconds(-1).AddDays(1);
            filteredList = recentDocList.GetAllGlobalDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by Updated At : {dateUpdated} - {dateTo}");
            Assert.That(filteredList, Has.All.Property(nameof(GlobalDocumentListItem.UpdatedAt)).InRange(dateUpdated, endDate), $"Filtered list has items out of Updated At : {dateUpdated} - {endDate}");

            //Get properties for matter attached with document
            randomDoc = GetRandomNumber(filteredList.Count - 1);
            selectedDocument = recentDocList.GetGlobalDocumentListItemByIndex(randomDoc);
            selectedDocument.NavigateToSummary();
            documentSummary.NavigateToParentMatter();

            var docMatterType = matterDetails.MatterType;
            var docMatterPABU = matterDetails.PracticeAreaBusinessUnit;

            globalDocumentsPage.Open();

            //Apply Filter - Matter Type & Pabu
            recentDocList.OpenListOptionsMenu().RestoreDefaults();
            recentDocList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Matter Type"].Set(docMatterType);
            documentsFilterDialog.Controls["Practice Area - Business Unit"].Set(docMatterPABU);
            documentsFilterDialog.Apply();
            Assert.That(recentDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(recentDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(recentDocList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            filteredList = recentDocList.GetAllGlobalDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by Matter Type - {docMatterType} and Pabu - {docMatterPABU}");

            randomDoc = GetRandomNumber(filteredList.Count - 1);
            selectedDocument = recentDocList.GetGlobalDocumentListItemByIndex(randomDoc);
            var docName = selectedDocument.Name;
            selectedDocument.NavigateToSummary();
            documentSummary.NavigateToParentMatter();
            Assert.That(matterDetails.MatterType, Is.EqualTo(docMatterType), $"Filtered list has document {docName} with different matter type applied filter - {docMatterType}");
            Assert.That(matterDetails.PracticeAreaBusinessUnit, Is.EqualTo(docMatterPABU), $"Filtered list has document {docName} with different Pabu applied filter - {docMatterPABU}");

            // Verify filter persist after navigation
            globalDocumentsPage.Open();
            Assert.That(recentDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(recentDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(GlobalDocumentsTestCategory)]
        [Description("Test case reference: Filters & Smoke Test Views")]
        public void CheckedOutDocumentsFilters()
        {
            var testDataCounter = 0;

            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            globalDocumentsPage.Open();
            globalDocumentsPage.OpenCheckedOutDocumentsList();

            var documentSummary = _outlook.Oc.DocumentSummaryPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var checkedOutDocList = globalDocumentsPage.ItemList;
            var matterDocList = _outlook.Oc.DocumentsListPage;
            var recentDocList = globalDocumentsPage.ItemList;

            var unfilteredCount = globalDocumentsPage.ItemList.GetCount();

            //Prepare test documents in case no documents are present
            if (unfilteredCount == 0)
            {
                _mattersListPage.Open();
                _mattersListPage.ItemList.OpenFirst();
                matterDetails.Tabs.Open("Documents");

                var testDoc1 = CreateDocument(OfficeApp.Notepad);
                var testDoc1NameWithoutExtension = Path.GetFileNameWithoutExtension(testDoc1.Name);
                DragAndDrop.FromFileSystem(testDoc1, documentSummary.DropPoint.GetElement());
                matterDocList.AddDocumentDialog.UploadDocument();
                unfilteredCount++;
                var matterDocument1 = matterDetails.ItemList.GetMatterDocumentListItemFromText(testDoc1.Name);
                matterDocument1.FileOptions.CheckOut();

                var notepad = new Notepad(testDoc1NameWithoutExtension);
                notepad.Close();

                var testDoc2 = CreateDocument(OfficeApp.Notepad);
                var testDoc2NameWithoutExtension = Path.GetFileNameWithoutExtension(testDoc2.Name);
                DragAndDrop.FromFileSystem(testDoc2, documentSummary.DropPoint.GetElement());
                matterDocList.AddDocumentDialog.UploadDocument();
                unfilteredCount++;
                var matterDocument2 = matterDetails.ItemList.GetMatterDocumentListItemFromText(testDoc2.Name);
                matterDocument2.FileOptions.CheckOut();

                notepad = new Notepad(testDoc2NameWithoutExtension);
                notepad.Close();

                globalDocumentsPage.Open();
                globalDocumentsPage.OpenCheckedOutDocumentsList();
                testDataCounter++;
            }

            //Select random Document
            var randomDoc = GetRandomNumber(unfilteredCount - 1);
            var selectedDocument = checkedOutDocList.GetGlobalDocumentListItemByIndex(randomDoc);

            //Fetch Document Properties
            var name = GetRandomSubstring(selectedDocument.Name);
            const string status = CheckInStatus.CheckedOut;
            var updatedBy = selectedDocument.CreatedByFullName;
            var matterName = GetRandomSubstring(selectedDocument.AssociatedEntityName.Replace("Matter: ", string.Empty));
            var dateUpdated = selectedDocument.UpdatedAt.Date;

            //Apply Filter - Name
            checkedOutDocList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            var documentsFilterDialog = globalDocumentsPage.GlobalDocumentsListFilterDialog;
            documentsFilterDialog.Controls["Name"].Set(name);
            documentsFilterDialog.Apply();
            Assert.That(checkedOutDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(checkedOutDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(checkedOutDocList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            var filteredList = checkedOutDocList.GetAllGlobalDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by FileName - {name}");
            Assert.That(filteredList, Has.All.Property(nameof(GlobalDocumentListItem.Name)).Contains(name).IgnoreCase, $"Filtered list has items not containing Name - {name}");

            // Save View
            checkedOutDocList.OpenListOptionsMenu().SaveCurrentView();
            globalDocumentsPage.SaveCurrentViewDialog.Controls["Create New"].Set(ViewName);
            globalDocumentsPage.SaveCurrentViewDialog.Save();

            // Clear Filters
            checkedOutDocList.OpenListOptionsMenu().RestoreDefaults();
            Assert.That(checkedOutDocList.IsFilterIconVisible, Is.False, "Filter Icon is visible");

            // Apply Saved View
            checkedOutDocList.OpenListOptionsMenu().ApplySavedView(ViewName);
            Assert.AreEqual(filteredList.Count, checkedOutDocList.GetCount());

            // Set view as default
            checkedOutDocList.OpenListOptionsMenu().SetCurrentViewAsDefault();

            // Verify current view is set as default
            checkedOutDocList.OpenListOptionsMenu().RestoreDefaults();
            Assert.That(checkedOutDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(checkedOutDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");

            // Clear user default
            checkedOutDocList.OpenListOptionsMenu().ClearUserDefault();
            Assert.That(checkedOutDocList.IsFilterIconVisible, Is.False, "Filter Icon is visible");

            //Apply Filter - Status
            checkedOutDocList.OpenListOptionsMenu().RestoreDefaults();
            checkedOutDocList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Status"].Set(status);
            documentsFilterDialog.Apply();
            Assert.That(checkedOutDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(checkedOutDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(checkedOutDocList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            filteredList = checkedOutDocList.GetAllGlobalDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by Status - {status}");
            Assert.That(filteredList, Has.All.Property(nameof(GlobalDocumentListItem.Status)).EqualTo(status));

            //Apply Filter - Updated By
            checkedOutDocList.OpenListOptionsMenu().RestoreDefaults();
            checkedOutDocList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Updated By"].Set(updatedBy);
            documentsFilterDialog.Apply();
            Assert.That(checkedOutDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(checkedOutDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(checkedOutDocList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            filteredList = checkedOutDocList.GetAllGlobalDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by Updated By - {updatedBy}");
            Assert.That(filteredList, Has.All.Property(nameof(GlobalDocumentListItem.CreatedByFullName)).EqualTo(updatedBy), $"Filtered list has items with update by user other than- {updatedBy}");

            //Apply Filter - Matter Name
            checkedOutDocList.OpenListOptionsMenu().RestoreDefaults();
            checkedOutDocList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Matter Name"].Set(matterName);
            documentsFilterDialog.Apply();
            Assert.That(checkedOutDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(checkedOutDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(checkedOutDocList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            filteredList = checkedOutDocList.GetAllGlobalDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by Matter Name : {matterName}");
            Assert.That(filteredList, Has.All.Property(nameof(GlobalDocumentListItem.AssociatedEntityName)).Contains(matterName).IgnoreCase, $"Filtered list has items not attached with matter name - {matterName}");

            //Apply Filter - Updated At
            checkedOutDocList.OpenListOptionsMenu().RestoreDefaults();
            var dateTo = dateUpdated.AddDays(3);
            checkedOutDocList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Updated At"].Set(FormatDateRange(dateUpdated, dateTo));
            documentsFilterDialog.Apply();
            Assert.That(checkedOutDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(checkedOutDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(checkedOutDocList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            var endDate = dateTo.AddSeconds(-1).AddDays(1);
            filteredList = checkedOutDocList.GetAllGlobalDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"(No records found on applying filter by Updated At : {dateUpdated} - {dateTo}");
            Assert.That(filteredList, Has.All.Property(nameof(GlobalDocumentListItem.UpdatedAt)).InRange(dateUpdated, endDate), $"Filtered list has items out of Updated At : {dateUpdated} - {endDate}");

            //Get properties for matter attached with document
            randomDoc = GetRandomNumber(filteredList.Count - 1);
            selectedDocument = checkedOutDocList.GetGlobalDocumentListItemByIndex(randomDoc);
            selectedDocument.NavigateToSummary();
            documentSummary.NavigateToParentMatter();

            var docMatterType = matterDetails.MatterType;
            var docMatterPABU = matterDetails.PracticeAreaBusinessUnit;

            _outlook.Oc.Header.NavigateBack();
            _outlook.Oc.Header.NavigateBack();

            //Apply Filter - Matter Type & Pabu
            checkedOutDocList.OpenListOptionsMenu().RestoreDefaults();
            checkedOutDocList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Matter Type"].Set(docMatterType);
            documentsFilterDialog.Controls["Practice Area - Business Unit"].Set(docMatterPABU);
            documentsFilterDialog.Apply();
            Assert.That(checkedOutDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(checkedOutDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(checkedOutDocList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            filteredList = checkedOutDocList.GetAllGlobalDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by Matter Type - {docMatterType} and Pabu - {docMatterPABU}");

            randomDoc = GetRandomNumber(filteredList.Count - 1);
            selectedDocument = checkedOutDocList.GetGlobalDocumentListItemByIndex(randomDoc);
            var docName = selectedDocument.Name;
            selectedDocument.NavigateToSummary();
            documentSummary.NavigateToParentMatter();
            Assert.That(matterDetails.MatterType, Is.EqualTo(docMatterType), $"Filtered list has document {docName} with different matter type applied filter - {docMatterType}");
            Assert.That(matterDetails.PracticeAreaBusinessUnit, Is.EqualTo(docMatterPABU), $"Filtered list has document {docName} with different Pabu applied filter - {docMatterPABU}");

            // Verify filter persist after navigation
            globalDocumentsPage.Open();
            globalDocumentsPage.OpenCheckedOutDocumentsList();
            Assert.That(checkedOutDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(checkedOutDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");

            //Test documents clean up
            checkedOutDocList.OpenListOptionsMenu().RestoreDefaults();
            if (testDataCounter > 0)
            {
                globalDocumentsPage.OpenRecentDocumentsList();
                var doc1 = recentDocList.GetGlobalDocumentListItemByIndex(0);
                var doc2 = recentDocList.GetGlobalDocumentListItemByIndex(1);

                doc1.FileOptions.DiscardCheckOutAndRemoveLocalCopy();
                doc1.Delete().Confirm();

                doc2.FileOptions.DiscardCheckOutAndRemoveLocalCopy();
                doc2.Delete().Confirm();
            }
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(GlobalDocumentsTestCategory)]
        [Description("Test case reference: Filters & Smoke Test Views")]
        public void AllDocumentsFilters()
        {
            var testDataCounter = 0;

            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.ShowResultAllDocumentsList();

            var documentSummary = _outlook.Oc.DocumentSummaryPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var allDocList = globalDocumentsPage.ItemList;
            var matterDocList = _outlook.Oc.DocumentsListPage;
            var recentDocList = globalDocumentsPage.ItemList;

            var unfilteredCount = globalDocumentsPage.ItemList.GetCount();

            //Prepare test documents in case no documents are present
            if (unfilteredCount == 0)
            {
                _mattersListPage.Open();
                _mattersListPage.ItemList.OpenFirst();
                matterDetails.Tabs.Open("Documents");

                var testDoc1 = CreateDocument(OfficeApp.Notepad);
                DragAndDrop.FromFileSystem(testDoc1, documentSummary.DropPoint.GetElement());
                matterDocList.AddDocumentDialog.UploadDocument();
                unfilteredCount++;

                var testDoc2 = CreateDocument(OfficeApp.Notepad);
                DragAndDrop.FromFileSystem(testDoc2, documentSummary.DropPoint.GetElement());
                matterDocList.AddDocumentDialog.UploadDocument();
                unfilteredCount++;

                globalDocumentsPage.Open();
                globalDocumentsPage.OpenAllDocumentsList();
                testDataCounter++;
            }

            //Select random Document
            var randomDoc = GetRandomNumber(unfilteredCount - 1);
            var selectedDocument = allDocList.GetGlobalDocumentListItemByIndex(randomDoc);

            //Fetch Document Properties
            var name = GetRandomSubstring(selectedDocument.Name);
            var status = selectedDocument.Status;
            var updatedBy = selectedDocument.CreatedByFullName;
            var matterName = GetRandomSubstring(selectedDocument.AssociatedEntityName.Replace("Matter: ", string.Empty));
            var dateUpdated = selectedDocument.UpdatedAt.Date;

            //Apply Filter - Name
            allDocList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            var documentsFilterDialog = globalDocumentsPage.GlobalDocumentsListFilterDialog;
            documentsFilterDialog.Controls["Name"].Set(name);
            documentsFilterDialog.Apply();
            Assert.That(allDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(allDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(allDocList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            var filteredList = allDocList.GetAllGlobalDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by FileName - {name}");
            Assert.That(filteredList, Has.All.Property(nameof(GlobalDocumentListItem.Name)).Contains(name).IgnoreCase, $"Filtered list has items not containing Name - {name}");

            // Save View
            allDocList.OpenListOptionsMenu().SaveCurrentView();
            globalDocumentsPage.SaveCurrentViewDialog.Controls["Create New"].Set(ViewName);
            globalDocumentsPage.SaveCurrentViewDialog.Save();

            // Clear Filters
            allDocList.OpenListOptionsMenu().RestoreDefaults();
            Assert.That(allDocList.IsFilterIconVisible, Is.False, "Filter Icon is visible");

            // Apply Saved View
            allDocList.OpenListOptionsMenu().ApplySavedView(ViewName);
            Assert.AreEqual(filteredList.Count, allDocList.GetCount());

            // Set view as default
            allDocList.OpenListOptionsMenu().SetCurrentViewAsDefault();

            // Verify current view is set as default
            allDocList.OpenListOptionsMenu().RestoreDefaults();
            Assert.That(allDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(allDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");

            // Clear user default
            allDocList.OpenListOptionsMenu().ClearUserDefault();
            Assert.That(allDocList.IsFilterIconVisible, Is.False, "Filter Icon is visible");

            //Apply Filter - Status
            allDocList.OpenListOptionsMenu().RestoreDefaults();
            allDocList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Status"].Set(status);
            documentsFilterDialog.Apply();
            Assert.That(allDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(allDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(allDocList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            filteredList = allDocList.GetAllGlobalDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by Status - {status}");
            Assert.That(filteredList, Has.All.Property(nameof(GlobalDocumentListItem.Status)).EqualTo(status));

            //Apply Filter - Updated By
            allDocList.OpenListOptionsMenu().RestoreDefaults();
            allDocList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Updated By"].Set(updatedBy);
            documentsFilterDialog.Apply();
            Assert.That(allDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(allDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(allDocList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            filteredList = allDocList.GetAllGlobalDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by Updated By - {updatedBy}");
            Assert.That(filteredList, Has.All.Property(nameof(GlobalDocumentListItem.CreatedByFullName)).EqualTo(updatedBy), $"Filtered list has items with update by user other than- {updatedBy}");

            //Apply Filter - Matter Name
            allDocList.OpenListOptionsMenu().RestoreDefaults();
            allDocList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Matter Name"].Set(matterName);
            documentsFilterDialog.Apply();
            Assert.That(allDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(allDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(allDocList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            filteredList = allDocList.GetAllGlobalDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by Matter Name : {matterName}");
            Assert.That(filteredList, Has.All.Property(nameof(GlobalDocumentListItem.AssociatedEntityName)).Contains(matterName).IgnoreCase, $"Filtered list has items not attached with matter name - {matterName}");

            //Apply Filter - Updated At
            allDocList.OpenListOptionsMenu().RestoreDefaults();
            var dateTo = dateUpdated.AddDays(1);
            allDocList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Updated At"].Set(FormatDateRange(dateUpdated, dateTo));
            documentsFilterDialog.Apply();
            Assert.That(allDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(allDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(allDocList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            var endDate = dateTo.AddSeconds(-1).AddDays(1);
            filteredList = allDocList.GetAllGlobalDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"(No records found on applying filter by Updated At : {dateUpdated} - {dateTo}");
            Assert.That(filteredList, Has.All.Property(nameof(GlobalDocumentListItem.UpdatedAt)).InRange(dateUpdated, endDate), $"Filtered list has items out of Updated At : {dateUpdated} - {endDate}");

            //Get properties for matter attached with document
            randomDoc = GetRandomNumber(filteredList.Count - 1);
            selectedDocument = allDocList.GetGlobalDocumentListItemByIndex(randomDoc);
            selectedDocument.NavigateToSummary();
            documentSummary.NavigateToParentMatter();

            var docMatterType = matterDetails.MatterType;
            var docMatterPABU = matterDetails.PracticeAreaBusinessUnit;

            _outlook.Oc.Header.NavigateBack();
            _outlook.Oc.Header.NavigateBack();

            //Apply Filter - Matter Type & Pabu
            allDocList.OpenListOptionsMenu().RestoreDefaults();
            allDocList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Matter Type"].Set(docMatterType);
            documentsFilterDialog.Controls["Practice Area - Business Unit"].Set(docMatterPABU);
            documentsFilterDialog.Apply();
            Assert.That(allDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(allDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(allDocList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            filteredList = allDocList.GetAllGlobalDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by Matter Type - {docMatterType} and Pabu - {docMatterPABU}");

            randomDoc = GetRandomNumber(filteredList.Count - 1);
            selectedDocument = allDocList.GetGlobalDocumentListItemByIndex(randomDoc);
            var docName = selectedDocument.Name;
            selectedDocument.NavigateToSummary();
            documentSummary.NavigateToParentMatter();
            Assert.That(matterDetails.MatterType, Is.EqualTo(docMatterType), $"Filtered list has document {docName} with different matter type applied filter - {docMatterType}");
            Assert.That(matterDetails.PracticeAreaBusinessUnit, Is.EqualTo(docMatterPABU), $"Filtered list has document {docName} with different Pabu applied filter - {docMatterPABU}");

            // Verify filter persist after navigation
            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            Assert.That(allDocList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(allDocList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");

            //Test documents clean up
            allDocList.OpenListOptionsMenu().RestoreDefaults();
            if (testDataCounter > 0)
            {
                globalDocumentsPage.OpenRecentDocumentsList();
                var doc1 = recentDocList.GetGlobalDocumentListItemByIndex(0);
                var doc2 = recentDocList.GetGlobalDocumentListItemByIndex(1);

                doc1.Delete().Confirm();
                doc2.Delete().Confirm();
            }
        }

        [TearDown]
        public void TearDown()
        {
            SaveScreenShotsAndLogs(_outlook);
            _outlook?.Destroy();
        }
    }
}
