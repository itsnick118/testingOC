using System;
using System.Collections.Generic;
using System.Linq;
using IntegratedDriver;
using NUnit.Framework;
using UITests.PageModel;
using UITests.PageModel.OfficeApps;
using UITests.PageModel.Shared;
using UITests.PageModel.Shared.Comparators;
using static UITests.Constants;
using static UITests.TestHelpers;

namespace UITests.RegressionTesting
{
    [TestFixture]
    public class GlobalDocumentsRegressionTests : UITestBase
    {
        private Outlook _outlook;
        private Word _word;

        [SetUp]
        public void SetUp()
        {
            _outlook = new Outlook(TestEnvironment);
            _outlook.Launch();
            _outlook.Oc.BasicSettingsPage.LogInAsStandardUser();
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_15767  Verify sorting on GDA All Documents and Checked out Tabs")]
        public void SortingOnGDLAllAndCheckedOutDocuments()
        {
            const int UniqueDocumentsNumber = 2;
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var documentsList = globalDocumentsPage.ItemList;
            var allDocumentsSortDialog = globalDocumentsPage.AllDocumentsSortDialog;
            var checkedOutDocumentsSortDialog = globalDocumentsPage.CheckedOutDocumentsSortDialog;

            globalDocumentsPage.Open();

            //All documents sort icon(visibility,color) and restore default visibility verification
            globalDocumentsPage.OpenAllDocumentsList();
            Assert.AreEqual(true, documentsList.IsSortIconVisible);
            Assert.AreEqual(BlackColorName, documentsList.GetSortIconColor().Name);
            Assert.AreEqual(false, allDocumentsSortDialog.IsSortRestoreDefaultPresent());
            allDocumentsSortDialog.Sort("Document Size", SortOrder.Ascending);
            Assert.AreEqual(false, allDocumentsSortDialog.IsSortRestoreDefaultPresent());
            Assert.AreEqual(BlueColorName, documentsList.GetSortIconColor().Name);

            //Checked out documents sort icon(visibility,color) and restore default visibility verification.
            globalDocumentsPage.OpenCheckedOutDocumentsList();
            Assert.AreEqual(true, documentsList.IsSortIconVisible);
            Assert.AreEqual(BlackColorName, documentsList.GetSortIconColor().Name);
            Assert.AreEqual(false, checkedOutDocumentsSortDialog.IsSortRestoreDefaultPresent());
            checkedOutDocumentsSortDialog.Sort("Document Size", SortOrder.Ascending);
            Assert.AreEqual(false, checkedOutDocumentsSortDialog.IsSortRestoreDefaultPresent());
            Assert.AreEqual(BlueColorName, documentsList.GetSortIconColor().Name);

            //validate sorting on checked out documents with different aspects.
            checkedOutDocumentsSortDialog.RestoreSortDefaults();
            var sortingParameters = new[] {
                nameof(GlobalDocumentListItem.Name),
                nameof(GlobalDocumentListItem.CreatedByFullName),
                nameof(GlobalDocumentListItem.UpdatedAt),
                nameof(GlobalDocumentListItem.DocumentSize)
                };

            //create checked out document test data if not present.
            var documentsCreated = new List<string>();
            var documents = documentsList.GetAllGlobalDocumentListItems();
            if (documents.GroupBy(x => x.Name).Count() < UniqueDocumentsNumber)
            {
                // no documents to check sorting, upload new documents
                mattersListPage.Open();
                mattersListPage.ItemList.OpenRandom();
                matterDetailsPage.Tabs.Open("Documents");

                for (var i = 0; i < UniqueDocumentsNumber; i++)
                {
                    var file = CreateDocument(OfficeApp.Notepad, GetRandomText((1024 * i) + 1));
                    DragAndDrop.FromFileSystem(file, matterDetailsPage.DropPoint.GetElement());
                    documentsListPage.AddDocumentDialog.UploadDocument();

                    var uploadedDocument = matterDetailsPage.ItemList.GetMatterDocumentListItemFromText(file.Name);
                    Assert.IsNotNull(uploadedDocument);

                    uploadedDocument.FileOptions.CheckOut();
                    var notepad = new Notepad(file.Name);
                    notepad.Close();
                    documentsCreated.Add(file.Name);
                }

                globalDocumentsPage.Open();
                globalDocumentsPage.OpenCheckedOutDocumentsList();
            }

            //validate sorting on checked out documents with different aspects.
            foreach (var sortParameter in sortingParameters)
            {
                checkedOutDocumentsSortDialog.Sort(TestHelpers.AddSpacesToTextAtCamelCase(sortParameter), SortOrder.Ascending);
                documents = documentsList.GetAllGlobalDocumentListItems();
                if (sortParameter == "DocumentSize")
                {
                    Assert.That(documents, Is.Ordered.Ascending.By(sortParameter).Using(new DocumentSizeComparer()));
                }
                else
                {
                    Assert.That(documents, Is.Ordered.Ascending.By(sortParameter));
                }

                checkedOutDocumentsSortDialog.Sort(TestHelpers.AddSpacesToTextAtCamelCase(sortParameter), SortOrder.Descending);
                documents = documentsList.GetAllGlobalDocumentListItems();
                if (sortParameter == "DocumentSize")
                {
                    Assert.That(documents, Is.Ordered.Descending.By(sortParameter).Using(new DocumentSizeComparer()));
                }
                else
                {
                    Assert.That(documents, Is.Ordered.Descending.By(sortParameter));
                }
            }

            //clean up
            foreach (var documentName in documentsCreated)
            {
                var document = documentsList.GetGlobalDocumentListItemFromText(documentName);
                document.Delete().Confirm();
            }

            //create all document test data if not present.
            documentsCreated = new List<string>();
            globalDocumentsPage.OpenAllDocumentsList();
            documents = globalDocumentsPage.ItemList.GetAllGlobalDocumentListItems();
            if (documents.GroupBy(x => x.Name).Count() < UniqueDocumentsNumber)
            {
                // no documents to check sorting, upload new documents
                mattersListPage.Open();
                mattersListPage.ItemList.OpenRandom();
                matterDetailsPage.Tabs.Open("Documents");

                for (var i = 0; i < UniqueDocumentsNumber; i++)
                {
                    var file = CreateDocument(OfficeApp.Notepad, GetRandomText((1024 * i) + 1));
                    DragAndDrop.FromFileSystem(file, matterDetailsPage.DropPoint.GetElement());
                    documentsListPage.AddDocumentDialog.UploadDocument();
                    var uploadedDocument = matterDetailsPage.ItemList.GetMatterDocumentListItemFromText(file.Name);
                    Assert.IsNotNull(uploadedDocument);
                    documentsCreated.Add(file.Name);
                }

                globalDocumentsPage.Open();
                globalDocumentsPage.OpenAllDocumentsList();
            }

            //validate sorting on all documents with different aspects.
            foreach (var sortParameter in sortingParameters)
            {
                allDocumentsSortDialog.Sort(TestHelpers.AddSpacesToTextAtCamelCase(sortParameter), SortOrder.Descending);
                documents = documentsList.GetAllGlobalDocumentListItems();
                if (sortParameter == "DocumentSize")
                {
                    Assert.That(documents, Is.Ordered.Descending.By(sortParameter).Using(new DocumentSizeComparer()));
                }
                else
                {
                    Assert.That(documents, Is.Ordered.Descending.By(sortParameter));
                }

                allDocumentsSortDialog.Sort(TestHelpers.AddSpacesToTextAtCamelCase(sortParameter), SortOrder.Ascending);
                documents = documentsList.GetAllGlobalDocumentListItems();
                if (sortParameter == "DocumentSize")
                {
                    Assert.That(documents, Is.Ordered.Ascending.By(sortParameter).Using(new DocumentSizeComparer()));
                }
                else
                {
                    Assert.That(documents, Is.Ordered.Ascending.By(sortParameter));
                }
            }

            //clean up
            foreach (var documentName in documentsCreated)
            {
                var document = documentsListPage.ItemList.GetGlobalDocumentListItemFromText(documentName);
                document.FileOptions.DiscardCheckOutAndRemoveLocalCopy();
                document.Delete().Confirm();
            }
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_17293  GDL: Verify to perform the list page opeartions ( View , Download, Delete)")]
        public void ViewDeleteDownloadOnGdl()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;

            var documentsList = documentsListPage.ItemList;
            var settingsPage = _outlook.Oc.SettingsPage;
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var globalDocumentsList = globalDocumentsPage.ItemList;

            // Open matter documents sub tab
            mattersList.OpenRandom();
            matterDetails.Tabs.Open("Documents");

            // Add new document
            var dndFileInfo = CreateDocument(OfficeApp.Word);
            DragAndDrop.FromFileSystem(dndFileInfo, matterDetails.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();

            var toastMessage = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(UploadSuccessMessage, toastMessage[0]);
            _outlook.Oc.CloseAllToastMessages();
            var uploadedDocument = documentsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            Assert.IsNotNull(uploadedDocument);

            // Go to global documents app then open all documents list page
            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.QuickSearch.SearchBy(dndFileInfo.Name);

            var addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(dndFileInfo.Name);
            Assert.IsNotNull(addedDocument);

            // Perform row click and open a document in read only mode
            addedDocument.Open();
            var fileName = addedDocument.Name;
            _word = new Word(TestEnvironment);
            _word.Attach(fileName);
            Assert.IsNotNull(_word.GetReadOnlyLabel());
            _word.Close();

            // Perform download action and save the document locally
            var fileInfo = addedDocument.Download($"{Guid.NewGuid()}.tmp");
            Assert.IsTrue(fileInfo.Exists);
            Assert.Greater(fileInfo.Length, 0);
            fileInfo.Delete();

            // Check out the same document and verify the banner message
            addedDocument.FileOptions.CheckOut();
            _word = new Word(TestEnvironment);
            _word.Attach(fileName);
            Assert.IsNull(_word.GetReadOnlyLabel());
            _word.Close();

            // Perform download action on a checked out document and save the document locally
            addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(dndFileInfo.Name);
            fileInfo = addedDocument.Download($"{Guid.NewGuid()}.tmp");
            Assert.IsTrue(fileInfo.Exists);
            Assert.Greater(fileInfo.Length, 0);
            fileInfo.Delete();

            // Perform Delete operation on checked out document
            addedDocument.Delete().Confirm();
            toastMessage = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(CheckedOutDocumentDeleteMessage, toastMessage[0]);
            _outlook.Oc.CloseAllToastMessages();
            addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(dndFileInfo.Name);
            Assert.IsNotNull(addedDocument);

            // Logout sbrown
            _outlook.Oc.OpenSettings();
            settingsPage.OpenConfiguration();
            settingsPage.LogOut().Confirm();

            // Log in office companion as dmaxwell
            _outlook.Oc.BasicSettingsPage.LogInAsAttorneyUser();
            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();

            globalDocumentsPage.QuickSearch.SearchBy(dndFileInfo.Name);
            addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(dndFileInfo.Name);
            Assert.IsNotNull(addedDocument);

            // Perform row click on a document which has been checked out by sbrown
            addedDocument.Open();
            toastMessage = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(ViewDocumentWarningMessage("sbrown"), toastMessage[0]);
            _word = new Word(TestEnvironment);
            _word.Attach(fileName);
            Assert.IsNotNull(_word.GetReadOnlyLabelForCheckedOutDocument());
            _word.Close();

            // Perform Delete operation on checked out document
            addedDocument.Delete().Confirm();
            toastMessage = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(DeleteDocumentMessageForDifferentUser("sbrown"), toastMessage[0]);
            _outlook.Oc.CloseAllToastMessages();
            addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(dndFileInfo.Name);
            Assert.IsNotNull(addedDocument);

            // Perform download action on a checked out document by different user and save the document locally
            fileInfo = addedDocument.Download($"{Guid.NewGuid()}.tmp");
            Assert.IsTrue(fileInfo.Exists);
            Assert.Greater(fileInfo.Length, 0);
            fileInfo.Delete();

            // Logout dmaxwell
            _outlook.Oc.OpenSettings();
            settingsPage.OpenConfiguration();
            settingsPage.LogOut().Confirm();

            // Log in office companion as sbrown
            _outlook.Oc.BasicSettingsPage.LogInAsStandardUser();
            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();

            // discard check out
            globalDocumentsPage.QuickSearch.SearchBy(dndFileInfo.Name);
            addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(dndFileInfo.Name);
            addedDocument.FileOptions.DiscardCheckOutAndRemoveLocalCopy();

            // Delete a checked in document
            addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(dndFileInfo.Name);
            addedDocument.Delete().Confirm();
            addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(dndFileInfo.Name);
            Assert.IsNull(addedDocument);
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_16963 : UI Verfication (JR - Recent Documents, Checked Out & All Documents)")]
        public void UIVerfication()
        {
            const int UniqueDocumentsNumber = 1;
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var documentsList = documentsListPage.ItemList;
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var documentSummaryPage = _outlook.Oc.DocumentSummaryPage;
            var globalDocumentsList = globalDocumentsPage.ItemList;

            const string ExpectedTab = "recent documents";

            //step1-Verify the Recent docs Lists in outlook
            globalDocumentsPage.Open();
            var selectTab = globalDocumentsPage.Tabs.GetActiveTab().ToLower();
            Assert.AreEqual(ExpectedTab, selectTab);

            //create recent document test data if not present.
            var documentsCreated = new List<string>();
            globalDocumentsPage.OpenAllDocumentsList();
            var documents = globalDocumentsPage.ItemList.GetAllGlobalDocumentListItems();
            if (documents.GroupBy(x => x.Name).Count() < UniqueDocumentsNumber)
            {
                // no documents to check sorting, upload new documents
                mattersListPage.Open();
                mattersListPage.ItemList.OpenRandom();
                matterDetailsPage.Tabs.Open("Documents");

                for (var i = 0; i < UniqueDocumentsNumber; i++)
                {
                    var file = CreateDocument(OfficeApp.Notepad, GetRandomText((1024 * i) + 1));
                    DragAndDrop.FromFileSystem(file, matterDetailsPage.DropPoint.GetElement());
                    documentsListPage.AddDocumentDialog.UploadDocument();
                    var uploadedDocument = matterDetailsPage.ItemList.GetMatterDocumentListItemFromText(file.Name);
                    Assert.IsNotNull(uploadedDocument);
                    documentsCreated.Add(file.Name);
                }

                globalDocumentsPage.Open();
            }

            //step2-Verify the fields in Narrow View
            var filteredList = globalDocumentsList.GetAllGlobalDocumentListItems();
            var randomDoc = GetRandomNumber(filteredList.Count - 1);
            var randomDocument = documentsList.GetGlobalDocumentListItemByIndex(randomDoc);
            Assert.IsNotNull(randomDocument.Name);
            Assert.True(randomDocument.IsDownloadIconVisible());
            Assert.IsNotNull(randomDocument.Status);
            Assert.IsNotNull(randomDocument.CreatedByFullName);
            Assert.True(randomDocument.IsDeleteButtonVisible());
            Assert.IsNotNull(randomDocument.DocumentSize);
            Assert.IsNotNull(randomDocument.UpdatedAt);
            Assert.IsNotNull(randomDocument.SecondaryElement);
            Assert.True(randomDocument.IsNavigateToSummaryVisible());

            //step3- Verify Tool tip
            Assert.AreEqual(randomDocument.DeleteButtonTooltip, "Delete document");
            Assert.AreEqual(randomDocument.NavigateToSummaryButtonTooltip, "Navigate to Summary");
            Assert.AreEqual(randomDocument.DownloadButtonTooltip, "Download document");

            //step4,5,6- Verify Document - Summary Page - Version History on Recent document tab
            randomDocument.NavigateToSummary();

            var versions = documentSummaryPage.ItemList.GetCount();
            Assert.AreEqual(versions, 1);

            var summaryInfo = documentSummaryPage.GetDocumentSummaryInfo();
            Assert.IsNotNull(summaryInfo);

            //clean up
            foreach (var documentName in documentsCreated)
            {
                var document = documentsListPage.ItemList.GetGlobalDocumentListItemFromText(documentName);
                document.Delete().Confirm();
            }
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_17295 : Verify breadcrumb navigations from document summary")]
        public void BreadcrumbNavigationFromDocumentSummary()
        {
            var folderName = GetRandomText(6);

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var documentsList = documentsListPage.ItemList;
            var documentSummary = _outlook.Oc.DocumentSummaryPage;
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var globalDocumentsList = globalDocumentsPage.ItemList;

            var dndFileInfo = CreateDocument(OfficeApp.Word);

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetails.Tabs.Open("Documents");
            documentsList.OpenAddFolderDialog();
            documentsListPage.AddFolderDialog.Controls["Name"].Set(folderName);
            documentsListPage.AddFolderDialog.Save();

            var testFolder = documentsList.GetMatterDocumentListItemFromText(folderName);
            Assert.IsNotNull(testFolder);

            testFolder.Open();

            //Upload document in folder
            DragAndDrop.FromFileSystem(dndFileInfo, matterDetails.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();

            var toastMessage = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(UploadSuccessMessage, toastMessage[0]);
            _outlook.Oc.CloseAllToastMessages();

            var uploadedDocument = documentsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            Assert.IsNotNull(uploadedDocument);

            var breadcrumbsPath = documentsListPage.BreadcrumbsControl.GetCurrentPath();
            Assert.IsTrue(breadcrumbsPath.EndsWith(folderName));

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.QuickSearch.SearchBy(dndFileInfo.Name);
            var addedDocumentToFolderLevel = globalDocumentsList.GetGlobalDocumentListItemFromText(dndFileInfo.Name);
            addedDocumentToFolderLevel.NavigateToSummary();

            var summaryInfo = documentSummary.GetDocumentSummaryInfo();
            Assert.IsNotEmpty(summaryInfo, "Document Summary fields are not retrieved or empty.");

            foreach (var field in summaryInfo)
            {
                Assert.IsNotEmpty(field.Text);
            }

            documentSummary.NavigateToParent();

            var breadCrumbsCurrentPath = documentsListPage.BreadcrumbsControl.GetCurrentPath();
            Assert.AreEqual(breadcrumbsPath, breadCrumbsCurrentPath);

            documentsListPage.BreadcrumbsControl.NavigateToTheRoot();

            testFolder = documentsList.GetMatterDocumentListItemFromText(folderName);
            testFolder.Delete().Confirm();

            // upload document to matter
            dndFileInfo = CreateDocument(OfficeApp.Word);
            DragAndDrop.FromFileSystem(dndFileInfo, matterDetails.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();

            uploadedDocument = documentsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            Assert.IsNotNull(uploadedDocument);

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.QuickSearch.SearchBy(dndFileInfo.Name);
            var addedDocumentToMatterLevel = globalDocumentsList.GetGlobalDocumentListItemFromText(dndFileInfo.Name);
            addedDocumentToMatterLevel.NavigateToSummary();

            summaryInfo = documentSummary.GetDocumentSummaryInfo();
            Assert.IsNotEmpty(summaryInfo, "Document Summary fields are not retrieved or empty.");

            foreach (var field in summaryInfo)
            {
                Assert.IsNotEmpty(field.Text);
            }

            documentSummary.NavigateToParentMatter();

            uploadedDocument = documentsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            Assert.IsNotNull(uploadedDocument);

            //Clean up
            uploadedDocument = documentsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            uploadedDocument.Delete().Confirm();
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_17019 : Verify the different filter criterias for the filters functionality on Global Documents App list page")]
        public void VerifyAllDocumentsFilters()
        {
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var documentSummary = _outlook.Oc.DocumentSummaryPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var allDocumentsList = globalDocumentsPage.ItemList;

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            allDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            // To select a document which is in checkedIn status only
            var documentsFilterDialog = globalDocumentsPage.GlobalDocumentsListFilterDialog;
            documentsFilterDialog.Controls["Status"].Set(CheckInStatus.CheckedIn);
            documentsFilterDialog.Controls["Name"].Set(".doc");
            documentsFilterDialog.Apply();

            //Select random Document
            var filteredCount = allDocumentsList.GetCount();
            var randomDocument = GetRandomNumber(filteredCount - 1);
            var selectedDocument = allDocumentsList.GetGlobalDocumentListItemByIndex(randomDocument);
            selectedDocument.NavigateToSummary();
            documentSummary.NavigateToParentMatter();

            var matterType = matterDetails.MatterType;
            var matterPABU = matterDetails.PracticeAreaBusinessUnit;
            const string OrganizationName = "Easterby";

            //Fetch Document Properties
            var fileName = selectedDocument.Name;
            var name = selectedDocument.Name;
            var status = selectedDocument.Status;
            var updatedBy = selectedDocument.CreatedByFullName;
            var matterName = GetRandomSubstring(selectedDocument.AssociatedEntityName.Replace("Matter: ", string.Empty));
            var dateUpdated = selectedDocument.UpdatedAt.Date;

            // Name
            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            allDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog = globalDocumentsPage.GlobalDocumentsListFilterDialog;
            documentsFilterDialog.Controls["Name"].Set(name);
            documentsFilterDialog.Apply();
            Assert.IsTrue(allDocumentsList.IsFilterIconVisible, "Filter Icon is not visible");
            Assert.AreEqual(allDocumentsList.GetFilterIconColor().Name, BlueColorName, "Filter Icon is not expected color");
            Assert.AreEqual(allDocumentsList.FilterIconToolTip, FilterIconToolTip, "Filter Icon tooltip differs from expected");

            var filteredList = allDocumentsList.GetAllGlobalDocumentListItems();
            Assert.IsNotEmpty(filteredList, $"No records found on applying filter by FileName - {name}");
            Assert.That(filteredList, Has.All.Property(nameof(GlobalDocumentListItem.Name)).Contains(name).IgnoreCase, $"Filtered list has items not containing Name - {name}");

            // check for 255 char limit on input text box
            allDocumentsList.OpenListOptionsMenu().RestoreDefaults();
            allDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            var randomString = GetRandomText(290);
            documentsFilterDialog.Controls["Name"].Set(randomString);
            var nameValue = documentsFilterDialog.Controls["Name"].GetValue();
            Assert.LessOrEqual(255, nameValue.Length);
            documentsFilterDialog.Cancel();

            // Clear Filters
            allDocumentsList.OpenListOptionsMenu().RestoreDefaults();
            Assert.IsFalse(allDocumentsList.IsFilterIconVisible, "Filter Icon is visible");

            // File Name
            allDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["File Name"].Set(fileName);
            documentsFilterDialog.Apply();
            Assert.IsTrue(allDocumentsList.IsFilterIconVisible, "Filter Icon is not visible");
            Assert.AreEqual(allDocumentsList.GetFilterIconColor().Name, BlueColorName, "Filter Icon is not expected color");
            Assert.AreEqual(allDocumentsList.FilterIconToolTip, FilterIconToolTip, "Filter Icon tooltip differs from expected");

            filteredList = allDocumentsList.GetAllGlobalDocumentListItems();
            Assert.IsNotEmpty(filteredList, $"No records found on applying filter by FileName - {fileName}");
            Assert.That(filteredList, Has.All.Property(nameof(GlobalDocumentListItem.Name)).Contains(fileName).IgnoreCase, $"Filtered list has items not containing Name - {fileName}");

            // check for 255 char limit on input text box
            allDocumentsList.OpenListOptionsMenu().RestoreDefaults();
            allDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            randomString = GetRandomText(290);
            documentsFilterDialog.Controls["File Name"].Set(randomString);
            var fileNameValue = documentsFilterDialog.Controls["File Name"].GetValue();
            Assert.LessOrEqual(255, fileNameValue.Length);
            documentsFilterDialog.Cancel();

            //Apply Filter - Status
            allDocumentsList.OpenListOptionsMenu().RestoreDefaults();
            allDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Status"].Set(status);
            documentsFilterDialog.Apply();
            Assert.IsTrue(allDocumentsList.IsFilterIconVisible, "Filter Icon is not visible");
            Assert.AreEqual(allDocumentsList.GetFilterIconColor().Name, BlueColorName, "Filter Icon is not expected color");
            Assert.AreEqual(allDocumentsList.FilterIconToolTip, FilterIconToolTip, "Filter Icon tooltip differs from expected");

            filteredList = allDocumentsList.GetAllGlobalDocumentListItems();
            Assert.IsNotEmpty(filteredList, $"No records found on applying filter by Status - {status}");
            Assert.That(filteredList, Has.All.Property(nameof(GlobalDocumentListItem.Status)).EqualTo(status));

            //Apply Filter - Updated At
            allDocumentsList.OpenListOptionsMenu().RestoreDefaults();
            var dateTo = dateUpdated.AddDays(3);
            allDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Updated At"].Set(FormatDateRange(dateUpdated, dateTo));
            documentsFilterDialog.Apply();
            Assert.IsTrue(allDocumentsList.IsFilterIconVisible, "Filter Icon is not visible");
            Assert.AreEqual(allDocumentsList.GetFilterIconColor().Name, BlueColorName, "Filter Icon is not expected color");
            Assert.AreEqual(allDocumentsList.FilterIconToolTip, FilterIconToolTip, "Filter Icon tooltip differs from expected");

            filteredList = allDocumentsList.GetAllGlobalDocumentListItems();
            Assert.IsNotEmpty(filteredList, $"No records found on applying filter by Updated At : {dateUpdated} - {dateTo}");

            //Apply Filter - Updated By
            allDocumentsList.OpenListOptionsMenu().RestoreDefaults();
            allDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Updated By"].Set(updatedBy);
            documentsFilterDialog.Apply();
            Assert.IsTrue(allDocumentsList.IsFilterIconVisible, "Filter Icon is not visible");
            Assert.AreEqual(allDocumentsList.GetFilterIconColor().Name, BlueColorName, "Filter Icon is not expected color");
            Assert.AreEqual(allDocumentsList.FilterIconToolTip, FilterIconToolTip, "Filter Icon tooltip differs from expected");

            filteredList = allDocumentsList.GetAllGlobalDocumentListItems();
            Assert.IsNotEmpty(filteredList, $"No records found on applying filter by Updated By - {updatedBy}");
            Assert.That(filteredList, Has.All.Property(nameof(GlobalDocumentListItem.CreatedByFullName)).EqualTo(updatedBy), $"Filtered list has items with update by user other than- {updatedBy}");

            // check for 255 char limit on input text box
            allDocumentsList.OpenListOptionsMenu().RestoreDefaults();
            allDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            randomString = GetRandomText(290);
            documentsFilterDialog.Controls["Updated By"].Set(randomString);
            var updatedByValue = documentsFilterDialog.Controls["Updated By"].GetValue();
            Assert.LessOrEqual(255, updatedByValue.Length);
            documentsFilterDialog.Cancel();

            // check for 255 char limit on input text box
            allDocumentsList.OpenListOptionsMenu().RestoreDefaults();
            allDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            randomString = GetRandomText(290);
            documentsFilterDialog.Controls["Comment"].Set(randomString);
            var commentValue = documentsFilterDialog.Controls["Comment"].GetValue();
            Assert.LessOrEqual(255, commentValue.Length);
            documentsFilterDialog.Cancel();

            // check for 2000 char limit on input text box
            allDocumentsList.OpenListOptionsMenu().RestoreDefaults();
            allDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            randomString = GetRandomText(2000);
            documentsFilterDialog.Controls["Content"].Set(randomString);
            documentsFilterDialog.Cancel();

            //Apply Filter - Matter Type
            allDocumentsList.OpenListOptionsMenu().RestoreDefaults();
            allDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Matter Type"].Set(matterType);
            documentsFilterDialog.Apply();
            Assert.IsTrue(allDocumentsList.IsFilterIconVisible, "Filter Icon is not visible");
            Assert.AreEqual(allDocumentsList.GetFilterIconColor().Name, BlueColorName, "Filter Icon is not expected color");
            Assert.AreEqual(allDocumentsList.FilterIconToolTip, FilterIconToolTip, "Filter Icon tooltip differs from expected");

            filteredList = allDocumentsList.GetAllGlobalDocumentListItems();
            Assert.IsNotEmpty(filteredList, $"No records found on applying filter by Matter Type - {matterType} ");

            //Apply Filter - PABU
            allDocumentsList.OpenListOptionsMenu().RestoreDefaults();
            allDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Practice Area - Business Unit"].Set(matterPABU);
            documentsFilterDialog.Apply();
            Assert.IsTrue(allDocumentsList.IsFilterIconVisible, "Filter Icon is not visible");
            Assert.AreEqual(allDocumentsList.GetFilterIconColor().Name, BlueColorName, "Filter Icon is not expected color");
            Assert.AreEqual(allDocumentsList.FilterIconToolTip, FilterIconToolTip, "Filter Icon tooltip differs from expected");

            filteredList = allDocumentsList.GetAllGlobalDocumentListItems();
            Assert.IsNotEmpty(filteredList, $"No records found on applying filter by Practice Area - Business Unit - {matterPABU} ");

            //Apply Filter - Matter Name
            allDocumentsList.OpenListOptionsMenu().RestoreDefaults();
            allDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Matter Name"].Set(matterName);
            documentsFilterDialog.Apply();
            Assert.IsTrue(allDocumentsList.IsFilterIconVisible, "Filter Icon is not visible");
            Assert.AreEqual(allDocumentsList.GetFilterIconColor().Name, BlueColorName, "Filter Icon is not expected color");
            Assert.AreEqual(allDocumentsList.FilterIconToolTip, FilterIconToolTip, "Filter Icon tooltip differs from expected");

            filteredList = allDocumentsList.GetAllGlobalDocumentListItems();
            Assert.IsNotEmpty(filteredList, $"No records found on applying filter by Matter Name : {matterName}");
            Assert.That(filteredList, Has.All.Property(nameof(GlobalDocumentListItem.AssociatedEntityName)).Contains(matterName).IgnoreCase, $"Filtered list has items not attached with matter name - {matterName}");

            // check for 255 char limit on input text box
            allDocumentsList.OpenListOptionsMenu().RestoreDefaults();
            allDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            randomString = GetRandomText(290);
            documentsFilterDialog.Controls["Matter Name"].Set(randomString);
            var matterNameValue = documentsFilterDialog.Controls["Matter Name"].GetValue();
            Assert.LessOrEqual(255, matterNameValue.Length);
            documentsFilterDialog.Cancel();

            // apply filter - organization name
            allDocumentsList.OpenListOptionsMenu().RestoreDefaults();
            allDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Organizations"].Set(OrganizationName);
            documentsFilterDialog.Apply();
            Assert.IsTrue(allDocumentsList.IsFilterIconVisible, "Filter Icon is not visible");

            filteredList = allDocumentsList.GetAllGlobalDocumentListItems();
            Assert.IsNotEmpty(filteredList, $"No records found on applying filter by Organization Name - {OrganizationName}");
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test Case 16175 part 1 : Verify to view banner message to expand the SPA(JR - All documents / Checked out/ Recent documents)[AP - MEDIUM, BP - MEDIUM]")]

        public void VerifyBannerMessageToExpandOC()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var settingsPage = _outlook.Oc.SettingsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var globalDocumentsList = globalDocumentsPage.ItemList;

            _outlook.Oc.OpenSettings();
            settingsPage.OpenAdvanced();
            settingsPage.SelectShowCollapsedOnStart();
            settingsPage.Apply();

            mattersListPage.Open();
            mattersListPage.ItemList.OpenRandom();
            matterDetailsPage.Tabs.Open("Documents");

            var testDocument = CreateDocument(OfficeApp.Word);
            DragAndDrop.FromFileSystem(testDocument, matterDetailsPage.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.QuickSearch.SearchBy(testDocument.Name);
            var addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(testDocument.Name);
            Assert.IsNotNull(addedDocument);
            Assert.AreEqual(CheckInStatus.CheckedIn, addedDocument.Status);

            addedDocument.Open();
            var fileName = addedDocument.Name;
            _word = new Word(TestEnvironment);
            _word.Attach(fileName);
            _word.ClickOnExpandBanner();
            Assert.IsNotNull(_word.GetReadOnlyLabel());

            _word.AttachToOc();
            _word.Oc.WaitForLoadComplete();
            _word.CheckOut();
            Assert.IsNull(_word.GetReadOnlyLabel());

            _word.Close();

            globalDocumentsPage.QuickSearch.SearchBy(testDocument.Name);
            addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(testDocument.Name);
            Assert.IsNotNull(addedDocument);
            addedDocument.Open();

            _word = new Word(TestEnvironment);
            _word.Attach(fileName);
            Assert.IsNull(_word.GetExpandBanner());
            _word.Close();

            _outlook.Oc.OpenSettings();
            settingsPage.OpenConfiguration();
            settingsPage.LogOut().Confirm();

            _outlook.Oc.BasicSettingsPage.LogIn();

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.QuickSearch.SearchBy(testDocument.Name);
            addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(testDocument.Name);
            Assert.IsNotNull(addedDocument);

            addedDocument.Open();
            _word = new Word(TestEnvironment);
            _word.Attach(fileName);
            Assert.IsNotNull(_word.GetExpandBanner());
            _word.ClickOnExpandBanner();
            Assert.NotNull(_word.GetReadOnlyLabelForCheckedOutDocument());
            _word.Close();

            // Clean Up
            _outlook.Oc.OpenSettings();
            settingsPage.OpenConfiguration();
            settingsPage.LogOut().Confirm();

            _outlook.Oc.BasicSettingsPage.LogInAsStandardUser();

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.QuickSearch.SearchBy(testDocument.Name);

            addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(testDocument.Name);
            addedDocument.FileOptions.DiscardCheckOutAndRemoveLocalCopy();

            addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(testDocument.Name);
            addedDocument.Delete().Confirm();
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_16175 part 2: Verify to view banner message to expand the SPA(JR - All documents / Checked out/ Recent documents)[AP - MEDIUM, BP - MEDIUM]")]
        public void VerifyCheckOutBannerBasedOnVersion()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var settingsPage = _outlook.Oc.SettingsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var globalDocumentsList = globalDocumentsPage.ItemList;
            var documentSummaryPage = _outlook.Oc.DocumentSummaryPage;
            var versionHistoryList = documentSummaryPage.ItemList;
            var checkInDialog = globalDocumentsPage.CheckInDocumentDialog;

            // Set show collapsed at start setting true
            _outlook.Oc.OpenSettings();
            settingsPage.OpenAdvanced();
            settingsPage.SelectShowCollapsedOnStart();
            settingsPage.Apply();

            mattersListPage.Open();
            mattersListPage.ItemList.OpenRandom();
            matterDetailsPage.Tabs.Open("Documents");

            var testDocument = CreateDocument(OfficeApp.Word);
            DragAndDrop.FromFileSystem(testDocument, matterDetailsPage.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.QuickSearch.SearchBy(testDocument.Name);
            var addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(testDocument.Name);
            Assert.IsNotNull(addedDocument);
            Assert.AreEqual(CheckInStatus.CheckedIn, addedDocument.Status);

            addedDocument.NavigateToSummary();
            documentSummaryPage.SummaryPanel.Toggle();

            var documentVersions = versionHistoryList.GetAllVersionHistoryListItems();
            documentVersions[0].Open();

            var fileName = addedDocument.Name;

            _word = new Word(TestEnvironment);
            _word.Attach(fileName);
            _word.ClickOnExpandBanner();
            Assert.IsNotNull(_word.GetReadOnlyLabel());
            _word.AttachToOc();
            _word.Oc.WaitForLoadComplete();
            _word.CheckOut();
            Assert.IsNull(_word.GetReadOnlyLabel());
            _word.Close();

            documentVersions = versionHistoryList.GetAllVersionHistoryListItems();
            documentVersions[0].Open();

            _word = new Word(TestEnvironment);
            _word.Attach(fileName);
            Assert.IsNull(_word.GetExpandBanner());
            _word.Close();

            _outlook.Oc.OpenSettings();
            settingsPage.OpenConfiguration();
            settingsPage.LogOut().Confirm();

            // Login as suser
            _outlook.Oc.BasicSettingsPage.LogIn();

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.QuickSearch.SearchBy(testDocument.Name);
            addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(testDocument.Name);
            Assert.IsNotNull(addedDocument);
            Assert.AreEqual(CheckInStatus.CheckedOut, addedDocument.Status);

            addedDocument.NavigateToSummary();
            documentSummaryPage.SummaryPanel.Toggle();

            documentVersions = versionHistoryList.GetAllVersionHistoryListItems();
            documentVersions[0].Open();

            _word = new Word(TestEnvironment);
            _word.Attach(fileName);
            Assert.IsNotNull(_word.GetExpandBanner());
            _word.ClickOnExpandBanner();
            Assert.NotNull(_word.GetReadOnlyLabelForCheckedOutDocument());
            _word.Close();

            //  Clean up
            _outlook.Oc.OpenSettings();
            settingsPage.OpenConfiguration();
            settingsPage.LogOut().Confirm();

            // Login as sbrown
            _outlook.Oc.BasicSettingsPage.LogInAsStandardUser();

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.QuickSearch.SearchBy(testDocument.Name);

            addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(testDocument.Name);
            addedDocument.FileOptions.CheckIn();
            checkInDialog.Controls["Comments"].Set("Clean Up Operation");
            documentsListPage.AddDocumentDialog.UploadDocument();

            // Delete document
            addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(testDocument.Name);
            addedDocument.Delete().Confirm();
        }

        [TearDown]
        public void TearDown()
        {
            SaveScreenShotsAndLogs(_outlook);
            _outlook?.Destroy();
        }
    }
}
