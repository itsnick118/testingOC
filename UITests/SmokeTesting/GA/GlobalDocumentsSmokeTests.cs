using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using IntegratedDriver;
using NUnit.Framework;
using UITests.PageModel;
using UITests.PageModel.OfficeApps;
using UITests.PageModel.Shared;
using UITests.PageModel.Shared.Comparators;
using static UITests.Constants;
using static UITests.TestHelpers;

namespace UITests.SmokeTesting.GA
{
    [TestFixture]
    public class GlobalDocumentsSmokeTests : UITestBase
    {
        private Outlook _outlook;
        private Word _word;
        private Excel _excel;
        private Powerpoint _powerpoint;

        [SetUp]
        public void SetUp()
        {
            _outlook = new Outlook(TestEnvironment);
            _outlook.Launch();
            _outlook.Oc.BasicSettingsPage.LogIn();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(GlobalDocumentsTestCategory)]
        public void DocumentDownload()
        {
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.ShowResultAllDocumentsList();
            var documentCount = globalDocumentsPage.ItemList.GetCount();
            while (documentCount < 1)
            {
                var wordFile = CreateDocument(OfficeApp.Word);
                var mattersListPage = _outlook.Oc.MattersListPage;
                var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
                var documentsListPage = _outlook.Oc.DocumentsListPage;

                DragAndDrop.FromFileSystem(wordFile, matterDetailsPage.DropPoint.GetElement());
                documentsListPage.AddDocumentDialog.UploadDocument();
                var uploadedDocument = documentsListPage.ItemList.GetMatterDocumentListItemFromText(wordFile.Name);
                Assert.IsNotNull(uploadedDocument);

                mattersListPage.Open();
                mattersListPage.ItemList.OpenRandom();
                matterDetailsPage.Tabs.Open("Documents");

                DragAndDrop.FromFileSystem(wordFile, matterDetailsPage.DropPoint.GetElement());
                documentsListPage.AddDocumentDialog.UploadDocument();

                globalDocumentsPage.Open();
                globalDocumentsPage.OpenAllDocumentsList();

                documentCount = globalDocumentsPage.ItemList.GetCount();
            }
            int randomDocNo = TestHelpers.GetRandomNumber(documentCount - 1);
            var file = globalDocumentsPage.ItemList.GetGlobalDocumentListItemByIndex(randomDocNo);
            var fileInfo = file.Download($"{Guid.NewGuid()}.tmp");
            Assert.Multiple(() =>
            {
                Assert.That(fileInfo.Exists, Is.True);
                Assert.That(fileInfo.Length, Is.GreaterThan(0));
            });

            File.Delete(fileInfo.FullName);
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(GlobalDocumentsTestCategory)]
        public void BannerOnLatestVersionOfCheckedInDocument()
        {
            var documentSummaryPage = _outlook.Oc.DocumentSummaryPage;
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var mattersListPage = _outlook.Oc.MattersListPage;

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.ItemList.OpenListOptionsMenu().OpenCreateListFilterDialog();

            var documentsFilterDialog = globalDocumentsPage.GlobalDocumentsListFilterDialog;
            documentsFilterDialog.Controls["Status"].Set(CheckInStatus.CheckedIn);
            documentsFilterDialog.Controls["File Name"].Set(".doc");
            documentsFilterDialog.Apply();

            var docCount = globalDocumentsPage.ItemList.GetCount();

            //Prepare test documents in case no documents are present
            if (docCount == 0)
            {
                var matterDocListPage = _outlook.Oc.DocumentsListPage;
                mattersListPage.Open();
                mattersListPage.ItemList.OpenRandom();
                matterDetailsPage.Tabs.Open("Documents");

                var testDoc1 = CreateDocument(OfficeApp.Word);
                DragAndDrop.FromFileSystem(testDoc1, documentSummaryPage.DropPoint.GetElement());
                matterDocListPage.AddDocumentDialog.UploadDocument();

                globalDocumentsPage.Open();
            }

            var selectedDocument = globalDocumentsPage.ItemList.GetGlobalDocumentListItemByIndex(0);
            Assert.IsNotNull(selectedDocument);

            var fileName = selectedDocument.Name;
            var documentStatus = selectedDocument.Status.ToLower();
            Assert.AreEqual(CheckInStatus.CheckedIn.ToLower(), documentStatus);

            selectedDocument.NavigateToSummary();
            documentSummaryPage.SummaryPanel.Toggle();
            var file = documentSummaryPage.ItemList.GetVersionHistoryListItemByIndex(0);

            file.Open();

            _word = new Word(TestEnvironment);
            _word.Attach(fileName);
            Assert.IsNotNull(_word.GetReadOnlyLabel());

            _word.Close();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(GlobalDocumentsTestCategory)]
        public void BannerOnLatestVersionOfCheckedInExcel()
        {
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var documentSummaryPage = _outlook.Oc.DocumentSummaryPage;
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.ItemList.OpenListOptionsMenu().OpenCreateListFilterDialog();

            var documentsFilterDialog = globalDocumentsPage.GlobalDocumentsListFilterDialog;
            documentsFilterDialog.Controls["Status"].Set(CheckInStatus.CheckedIn);
            documentsFilterDialog.Controls["File Name"].Set(".xlsx");
            documentsFilterDialog.Apply();

            var docCount = globalDocumentsPage.ItemList.GetCount();

            //Prepare test documents in case no documents are present
            if (docCount == 0)
            {
                var matterDocListPage = _outlook.Oc.DocumentsListPage;
                mattersListPage.Open();
                mattersListPage.ItemList.OpenRandom();
                matterDetailsPage.Tabs.Open("Documents");

                var testDoc1 = CreateDocument(OfficeApp.Excel);
                DragAndDrop.FromFileSystem(testDoc1, documentSummaryPage.DropPoint.GetElement());
                matterDocListPage.AddDocumentDialog.UploadDocument();

                globalDocumentsPage.Open();
            }

            var selectedDocument = globalDocumentsPage.ItemList.GetGlobalDocumentListItemByIndex(0);
            Assert.IsNotNull(selectedDocument);

            var fileName = selectedDocument.Name;
            var documentStatus = selectedDocument.Status.ToLower();
            Assert.AreEqual(CheckInStatus.CheckedIn.ToLower(), documentStatus);

            selectedDocument.NavigateToSummary();
            documentSummaryPage.SummaryPanel.Toggle();
            var file = documentSummaryPage.ItemList.GetVersionHistoryListItemByIndex(0);

            file.Open();

            _excel = new Excel(TestEnvironment);
            _excel.Attach(fileName);
            Assert.IsNotNull(_excel.GetReadOnlyLabel());

            _excel.Close();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(GlobalDocumentsTestCategory)]
        public void BannerOnLatestVersionOfCheckedInPowerpoint()
        {
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var documentSummaryPage = _outlook.Oc.DocumentSummaryPage;
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.ItemList.OpenListOptionsMenu().OpenCreateListFilterDialog();

            var documentsFilterDialog = globalDocumentsPage.GlobalDocumentsListFilterDialog;
            documentsFilterDialog.Controls["Status"].Set(CheckInStatus.CheckedIn);
            documentsFilterDialog.Controls["File Name"].Set(".pptx");
            documentsFilterDialog.Apply();

            var docCount = globalDocumentsPage.ItemList.GetCount();

            //Prepare test documents in case no documents are present
            if (docCount == 0)
            {
                var matterDocListPage = _outlook.Oc.DocumentsListPage;
                mattersListPage.Open();
                mattersListPage.ItemList.OpenRandom();
                matterDetailsPage.Tabs.Open("Documents");

                var testDoc1 = CreateDocument(OfficeApp.Powerpoint);
                DragAndDrop.FromFileSystem(testDoc1, documentSummaryPage.DropPoint.GetElement());
                matterDocListPage.AddDocumentDialog.UploadDocument();

                globalDocumentsPage.Open();
            }

            var selectedDocument = globalDocumentsPage.ItemList.GetGlobalDocumentListItemByIndex(0);
            Assert.IsNotNull(selectedDocument);

            var fileName = selectedDocument.Name;
            var documentStatus = selectedDocument.Status.ToLower();
            Assert.AreEqual(CheckInStatus.CheckedIn.ToLower(), documentStatus);

            selectedDocument.NavigateToSummary();
            documentSummaryPage.SummaryPanel.Toggle();
            var file = documentSummaryPage.ItemList.GetVersionHistoryListItemByIndex(0);

            file.Open();

            _powerpoint = new Powerpoint(TestEnvironment);
            _powerpoint.Attach(fileName);
            Assert.IsNotNull(_powerpoint.GetReadOnlyLabel());

            _powerpoint.Close();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(GlobalDocumentsTestCategory)]
        public void NavigateToMatterFromGDLDocSummary()
        {
            var wordFile = CreateDocument(OfficeApp.Word);

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var documentSummary = _outlook.Oc.DocumentSummaryPage;

            // Upload a document
            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetailsPage.Tabs.Open("Documents");

            DragAndDrop.FromFileSystem(wordFile, matterDetailsPage.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();

            var matterDocument = matterDetailsPage.ItemList.GetMatterDocumentListItemFromText(wordFile.Name);
            Assert.IsNotNull(matterDocument);

            // Verify scenario
            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.QuickSearch.SearchBy(wordFile.Name);

            var selectedDocument = globalDocumentsPage.ItemList.GetGlobalDocumentListItemFromText(wordFile.Name);
            Assert.IsNotNull(selectedDocument);

            var color = selectedDocument.GetColorOnHoverOverFileName();
            Assert.AreEqual(color.Name, BlueColorName);

            var textDecoration = selectedDocument.GetTextDecorationOnHoverOverFileName();
            StringAssert.Contains("underline", textDecoration);

            selectedDocument.NavigateToSummary();
            var summaryInfo = documentSummary.GetDocumentSummaryInfo();
            Assert.That(summaryInfo, Is.Not.Empty, "Document Summary fields are not retrieved or empty.");

            foreach (var webElement in summaryInfo)
            {
                Assert.IsNotEmpty(webElement.Text);
            }

            documentSummary.SummaryPanel.Toggle();

            summaryInfo = documentSummary.GetDocumentSummaryInfo();
            foreach (var webElement in summaryInfo)
            {
                Assert.That(webElement.Displayed, Is.False, "Summary Info is displayed on toggle");
            }

            documentSummary.NavigateToParentMatter();
            matterDocument = matterDetailsPage.ItemList.GetMatterDocumentListItemFromText(wordFile.Name);

            Assert.IsNotNull(matterDocument);
            Assert.AreEqual(matterDocument.Name, wordFile.Name);

            // Cleanup
            matterDocument.Delete().Confirm();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(GlobalDocumentsTestCategory)]
        [Category(DataDependentTestCategory)]
        public void NoBannerOnLatestVersionOfCheckedOutDocument()
        {
            var documentSummaryPage = _outlook.Oc.DocumentSummaryPage;
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var mattersListPage = _outlook.Oc.MattersListPage;

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.ItemList.OpenListOptionsMenu().OpenCreateListFilterDialog();

            var documentsFilterDialog = globalDocumentsPage.GlobalDocumentsListFilterDialog;
            documentsFilterDialog.Controls["Status"].Set(CheckInStatus.CheckedOut);
            documentsFilterDialog.Controls["File Name"].Set(".doc");
            documentsFilterDialog.Apply();

            var docCount = globalDocumentsPage.ItemList.GetCount();

            //Prepare test documents in case no documents are present
            if (docCount == 0)
            {
                var matterDocListPage = _outlook.Oc.DocumentsListPage;
                mattersListPage.Open();
                mattersListPage.ItemList.OpenRandom();
                matterDetailsPage.Tabs.Open("Documents");

                var testDoc1 = CreateDocument(OfficeApp.Word);
                DragAndDrop.FromFileSystem(testDoc1, documentSummaryPage.DropPoint.GetElement());
                matterDocListPage.AddDocumentDialog.UploadDocument();

                globalDocumentsPage.Open();
                var uploadedDocument = globalDocumentsPage.ItemList.GetGlobalDocumentListItemByIndex(0);
                uploadedDocument.FileOptions.CheckOut();
            }

            var selectedDocument = globalDocumentsPage.ItemList.GetGlobalDocumentListItemByIndex(0);
            Assert.IsNotNull(selectedDocument);

            var fileName = selectedDocument.Name;
            var documentStatus = selectedDocument.Status.ToLower();
            Assert.AreEqual(CheckInStatus.CheckedOut.ToLower(), documentStatus);

            selectedDocument.NavigateToSummary();
            documentSummaryPage.SummaryPanel.Toggle();
            var file = documentSummaryPage.ItemList.GetVersionHistoryListItemByIndex(0);

            file.Open();

            _word = new Word(TestEnvironment);
            _word.Attach(fileName);
            Assert.IsNull(_word.GetReadOnlyLabel(false));

            _word.Close();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(GlobalDocumentsTestCategory)]
        public void DocumentDelete()
        {
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.ItemList.OpenListOptionsMenu().OpenCreateListFilterDialog();

            var documentsFilterDialog = globalDocumentsPage.GlobalDocumentsListFilterDialog;
            documentsFilterDialog.Controls["Status"].Set(CheckInStatus.CheckedIn);
            documentsFilterDialog.Apply();

            var checkedInDocumentCount = globalDocumentsPage.ItemList.GetCount();
            while (checkedInDocumentCount < 1)
            {
                var testEmail = _outlook.AddTestEmailsToFolder(1, FileSize.Small, true).First();
                _outlook.OpenTestEmailFolder();
                _outlook.TurnOnReadingPane();

                var textfilename = new FileInfo(testEmail.Value).Name;
                var attachment = _outlook.GetAttachmentFromReadingPane(textfilename);

                var mattersListPage = _outlook.Oc.MattersListPage;
                var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
                var documentsListPage = _outlook.Oc.DocumentsListPage;

                mattersListPage.Open();
                mattersListPage.ItemList.OpenRandom();
                matterDetailsPage.Tabs.Open("Documents");

                DragAndDrop.FromElementToElement(attachment, matterDetailsPage.DropPoint.GetElement());
                documentsListPage.AddDocumentDialog.UploadDocument();

                globalDocumentsPage.Open();
                globalDocumentsPage.OpenAllDocumentsList();
                globalDocumentsPage.ItemList.OpenListOptionsMenu().OpenCreateListFilterDialog();

                documentsFilterDialog.Controls["Status"].Set(CheckInStatus.CheckedIn);
                documentsFilterDialog.Apply();

                checkedInDocumentCount = globalDocumentsPage.ItemList.GetCount();
            }

            var document = globalDocumentsPage.ItemList.GetGlobalDocumentListItemByIndex(0);

            var isDeleteButtonVisible = document.IsDeleteButtonVisible();
            Assert.That(isDeleteButtonVisible, Is.True, "Delete button is not visible without hovering over list item.");

            var buttonTooltip = document.DeleteButtonTooltip;
            var buttonColor = document.GetDeleteButtonColor();
            Assert.That(buttonTooltip, Is.EqualTo("Delete document"), "Delete button tooltip differs from expected");
            Assert.That(buttonColor.Name, Is.EqualTo(RedColorName), "Delete button color differs from expected.");

            document = globalDocumentsPage.ItemList.GetGlobalDocumentListItemByIndex(0);
            var nameBeforeRemoval = document.Name;

            document.Delete().Cancel();

            document = globalDocumentsPage.ItemList.GetGlobalDocumentListItemByIndex(0);
            var nameAfterCancelling = document.Name;
            Assert.AreEqual(nameBeforeRemoval, nameAfterCancelling, "Document was not found after clicking Cancel on delete dialog.");

            document.Delete().Confirm();

            var removedDocumentName = nameAfterCancelling;
            var firstDocumentInList = globalDocumentsPage.ItemList.GetGlobalDocumentListItemByIndex(0);
            Assert.That(removedDocumentName, Is.Not.EqualTo(firstDocumentInList.Name), "Document is present in the list after deleting it.");
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(GlobalDocumentsTestCategory)]
        public void DocumentSearch()
        {
            const string relatedObjectType = "matter";
            var wordFile = CreateDocument(OfficeApp.Word);

            var testCurrentUser = _outlook.CurrentUserDisplayName;

            // prepare well-known document
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;

            mattersListPage.Open();

            var matter = mattersListPage.ItemList.GetMatterDocumentListItemByIndex(0);
            matter.Open();
            matterDetailsPage.Tabs.Open("Documents");

            DragAndDrop.FromFileSystem(wordFile, matterDetailsPage.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();

            var uploadedDocument = documentsListPage.ItemList.GetMatterDocumentListItemFromText(wordFile.Name);
            Assert.IsNotNull(uploadedDocument);

            // search in global documents
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.ItemList.OpenListOptionsMenu().OpenCreateListFilterDialog();

            var documentsFilterDialog = globalDocumentsPage.GlobalDocumentsListFilterDialog;
            documentsFilterDialog.Controls["Status"].Set(CheckInStatus.CheckedIn);
            documentsFilterDialog.Apply();

            // search by file name
            globalDocumentsPage.QuickSearch.SearchBy(wordFile.Name);

            var foundDocumentsByFilename = globalDocumentsPage.ItemList.GetAllGlobalDocumentListItems();
            Assert.That(foundDocumentsByFilename.Count, Is.EqualTo(1), "Search by unique filename gave more or less than 1 document.");
            var actualDocumentName = foundDocumentsByFilename[0].Name;
            Assert.AreEqual(actualDocumentName, wordFile.Name, "Found document name differs from expected.");

            // search by document name
            globalDocumentsPage.QuickSearch.SearchBy(wordFile.Name);

            var foundDocumentsByDocumentName = globalDocumentsPage.ItemList.GetAllGlobalDocumentListItems();
            Assert.AreEqual(foundDocumentsByDocumentName.Count, 1, "Search by unique document name gave more or less than 1 document.");
            var actualDocumentName2 = foundDocumentsByDocumentName[0].Name;
            Assert.AreEqual(actualDocumentName2, wordFile.Name, "Found document name differs from expected.");

            // search by author
            globalDocumentsPage.QuickSearch.SearchBy(testCurrentUser);

            var foundDocumentsByAuthor = globalDocumentsPage.ItemList.GetAllGlobalDocumentListItems();
            Assert.That(foundDocumentsByAuthor.Count, Is.GreaterThan(0), "Search by author gave no results.");

            foreach (var document in foundDocumentsByAuthor)
            {
                Assert.AreEqual(document.CreatedByFullName, testCurrentUser, "Found document author differs from expected.");
            }

            // search by related object - matter
            globalDocumentsPage.QuickSearch.SearchBy(relatedObjectType);

            var foundDocumentsByMatter = globalDocumentsPage.ItemList.GetAllGlobalDocumentListItems();
            Assert.That(foundDocumentsByMatter.Count, Is.GreaterThan(0), "Search by 'matter' related object gave no results.");

            foreach (var document in foundDocumentsByMatter)
            {
                Assert.That(document.AssociatedEntityName, Does.StartWith("Matter:"), "Found document related object refers not to the 'matter' entity.");
            }
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(GlobalDocumentsTestCategory)]
        public void ViewDocumentFromVersionHistoryListPage()
        {
            const string docType = ".doc";

            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var documentSummaryPage = _outlook.Oc.DocumentSummaryPage;
            var versionHistoryPage = documentSummaryPage.ItemList;
            var globalDocumentsList = globalDocumentsPage.ItemList;
            _word = new Word(TestEnvironment);

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();

            // apply filter to select checked in document
            var documentsFilterDialog = globalDocumentsPage.GlobalDocumentsListFilterDialog;
            documentsFilterDialog.Controls["Status"].Set(CheckInStatus.CheckedIn);
            documentsFilterDialog.Controls["File Name"].Set(docType);
            documentsFilterDialog.Apply();

            var checkedInDocumentCount = globalDocumentsList.GetCount();
            while (checkedInDocumentCount < 1)
            {
                var matterDocListPage = _outlook.Oc.DocumentsListPage;
                var mattersListPage = _outlook.Oc.MattersListPage;
                var matterDetailsPage = _outlook.Oc.MatterDetailsPage;

                mattersListPage.Open();
                mattersListPage.ItemList.OpenRandom();
                matterDetailsPage.Tabs.Open("Documents");

                var testDoc1 = CreateDocument(OfficeApp.Word);
                DragAndDrop.FromFileSystem(testDoc1, matterDocListPage.DropPoint.GetElement());
                matterDocListPage.AddDocumentDialog.UploadDocument();

                globalDocumentsPage.Open();
                globalDocumentsPage.OpenAllDocumentsList();
                globalDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();

                documentsFilterDialog.Controls["Status"].Set(CheckInStatus.CheckedIn);
                documentsFilterDialog.Controls["File Name"].Set(docType);
                documentsFilterDialog.Apply();

                checkedInDocumentCount = globalDocumentsList.GetCount();
            }
            var selectedDocument = globalDocumentsList.GetGlobalDocumentListItemByIndex(0);
            Assert.IsNotNull(selectedDocument);

            var selectedDocumentName = selectedDocument.Name;

            var checkInDocumentDialog = globalDocumentsPage.CheckInDocumentDialog;
            var checkOutDocumentDialog = globalDocumentsPage.CheckOutDocumentDialog;

            // check out
            selectedDocument.FileOptions.CheckOut();
            checkOutDocumentDialog.Keep();
            _word.Attach(selectedDocumentName);
            _word.Close();

            globalDocumentsPage.OpenCheckedOutDocumentsList();
            globalDocumentsPage.QuickSearch.SearchBy(selectedDocumentName);

            // check in
            selectedDocument = globalDocumentsList.GetGlobalDocumentListItemByIndex(0);
            selectedDocument.FileOptions.CheckIn();
            checkInDocumentDialog.UploadDocument();

            // check out
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.QuickSearch.SearchBy(selectedDocumentName);
            selectedDocument = globalDocumentsList.GetGlobalDocumentListItemByIndex(0);
            selectedDocument.FileOptions.CheckOut();
            checkOutDocumentDialog.Keep();
            _word.Attach(selectedDocumentName);
            _word.Close();

            //check in
            globalDocumentsPage.OpenCheckedOutDocumentsList();
            globalDocumentsPage.QuickSearch.SearchBy(selectedDocumentName);
            selectedDocument = globalDocumentsList.GetGlobalDocumentListItemByIndex(0);
            selectedDocument.FileOptions.CheckIn();
            checkInDocumentDialog.UploadDocument();

            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.QuickSearch.SearchBy(selectedDocumentName);
            selectedDocument = globalDocumentsList.GetGlobalDocumentListItemByIndex(0);
            selectedDocument.NavigateToSummary();
            documentSummaryPage.SummaryPanel.Toggle();
            var file = versionHistoryPage.GetVersionHistoryListItemByIndex(0);

            // verify label on latest version in checked in state
            file.Open();
            _word.Attach(selectedDocumentName);
            Assert.That(_word.GetReadOnlyLabel(), Is.Not.Null, "Read Only Label is not displayed on latest checked in version");

            _word.Close();

            var ocHeader = _outlook.Oc.Header;
            ocHeader.NavigateBack();

            // check out
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.QuickSearch.SearchBy(selectedDocumentName);
            selectedDocument = globalDocumentsList.GetGlobalDocumentListItemByIndex(0);
            selectedDocument.FileOptions.CheckOut();
            checkOutDocumentDialog.Keep();
            _word.Attach(selectedDocumentName);
            _word.Close();

            globalDocumentsPage.OpenCheckedOutDocumentsList();
            globalDocumentsPage.QuickSearch.SearchBy(selectedDocumentName);
            selectedDocument = globalDocumentsList.GetGlobalDocumentListItemByIndex(0);
            selectedDocument.NavigateToSummary();
            file = versionHistoryPage.GetVersionHistoryListItemByIndex(0);

            // verify label on latest version in checked out state
            file.Open();
            _word = new Word(TestEnvironment);
            _word.Attach(selectedDocumentName);
            Assert.That(_word.GetReadOnlyLabel(false), Is.Null, "Read Only Label is displayed on latest checked out version");

            _word.Close();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(GlobalDocumentsTestCategory)]
        [Category(DataDependentTestCategory)]
        [Description("Test case reference: All Documents tab (View documents from multiple PABU)")]
        public void ViewDocumentsFromMultiplePABU()
        {
            // Assumption: There are matters with Primary and Secondary PABU defined in app.config exist
            var primaryPABU = _outlook.CurrentUserPrimaryPABU;
            var secondaryPABU = _outlook.CurrentUserSecondaryPABU;

            // Prepare well-known documents
            var emails = _outlook.AddTestEmailsToFolder(2, asAttachment: true);
            var filename1 = Path.GetFileNameWithoutExtension(emails.Values.ElementAt(0));
            var filename2 = Path.GetFileNameWithoutExtension(emails.Values.ElementAt(1));

            _outlook.OpenTestEmailFolder();
            _outlook.TurnOnReadingPane();

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var documentsFilter = globalDocumentsPage.GlobalDocumentsListFilterDialog;
            var mattersList = mattersListPage.ItemList;
            var globalDocumentsList = globalDocumentsPage.ItemList;

            // Upload documents
            mattersListPage.Open();
            mattersList.OpenListOptionsMenu().OpenCreateListFilterDialog();

            var mattersFilter = mattersListPage.MatterListFilterDialog;
            mattersFilter.Controls["Practice Area - Business Unit"].Set(primaryPABU);
            mattersFilter.Apply();

            mattersList.OpenFirst();
            matterDetailsPage.Tabs.Open("Documents");

            _outlook.SelectNthItem(0);
            var attachment1 = _outlook.GetAttachmentFromReadingPane(filename1);
            DragAndDrop.FromElementToElement(attachment1, matterDetailsPage.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();
            var uploadedDocument1 = documentsListPage.ItemList.GetMatterDocumentListItemFromText(filename1);
            uploadedDocument1.FileOptions.CheckOut();
            new Notepad(filename1).Close();

            mattersListPage.Open();
            mattersList.OpenListOptionsMenu().RestoreDefaults();
            mattersList.OpenListOptionsMenu().OpenCreateListFilterDialog();

            mattersFilter.Controls["Practice Area - Business Unit"].Set(secondaryPABU);
            mattersFilter.Apply();

            mattersList.OpenFirst();
            matterDetailsPage.Tabs.Open("Documents");

            _outlook.SelectNthItem(1);
            var attachment2 = _outlook.GetAttachmentFromReadingPane(filename2);
            DragAndDrop.FromElementToElement(attachment2, matterDetailsPage.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();
            var uploadedDocument2 = documentsListPage.ItemList.GetMatterDocumentListItemFromText(filename2);
            uploadedDocument2.FileOptions.CheckOut();
            new Notepad(filename2).Close();

            // Verify scenario
            globalDocumentsPage.Open();

            // All Documents - Primary PABU
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();

            documentsFilter.Controls["File Name"].Set(filename2);
            documentsFilter.Controls["Practice Area - Business Unit"].Set(primaryPABU);
            documentsFilter.Apply();

            var document2 = globalDocumentsList.GetGlobalDocumentListItemFromText(filename2, false);
            Assert.Null(document2, "The document from Secondary PABU is found by Primary PABU on All Documents tab.");

            globalDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilter.Controls["File Name"].Set(filename1);
            documentsFilter.Apply();

            var document1 = globalDocumentsList.GetGlobalDocumentListItemFromText(filename1);
            Assert.NotNull(document1, "The document is not found by Primary PABU and Filename on All Documents tab.");

            // Checked Out - Primary PABU
            globalDocumentsPage.OpenCheckedOutDocumentsList();
            globalDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();

            documentsFilter.Controls["File Name"].Set(filename2);
            documentsFilter.Controls["Practice Area - Business Unit"].Set(primaryPABU);
            documentsFilter.Apply();

            document2 = globalDocumentsList.GetGlobalDocumentListItemFromText(filename2, false);
            Assert.Null(document2, "The document from Secondary PABU is found by Primary PABU on Checked Out tab.");

            globalDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilter.Controls["File Name"].Set(filename1);
            documentsFilter.Apply();

            document1 = globalDocumentsList.GetGlobalDocumentListItemFromText(filename1);
            Assert.NotNull(document1, "The document is not found by Primary PABU and Filename on Checked Out tab.");

            // Recent Documents - Primary PABU
            globalDocumentsPage.OpenRecentDocumentsList();
            globalDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();

            documentsFilter.Controls["File Name"].Set(filename2);
            documentsFilter.Controls["Practice Area - Business Unit"].Set(primaryPABU);
            documentsFilter.Apply();

            document2 = globalDocumentsList.GetGlobalDocumentListItemFromText(filename2, false);
            Assert.Null(document2, "The document from Secondary PABU is found by Primary PABU on Recent Documents tab.");

            globalDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilter.Controls["File Name"].Set(filename1);
            documentsFilter.Apply();

            document1 = globalDocumentsList.GetGlobalDocumentListItemFromText(filename1);
            Assert.NotNull(document1, "The document is not found by Primary PABU and Filename on Recent Documents tab.");

            // Cleanup
            document1.FileOptions.DiscardCheckOutAndRemoveLocalCopy();
            document1.Delete().Confirm();

            // All Documents - Secondary PABU
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsList.OpenListOptionsMenu().RestoreDefaults();
            globalDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();

            documentsFilter.Controls["File Name"].Set(filename1);
            documentsFilter.Controls["Practice Area - Business Unit"].Set(secondaryPABU);
            documentsFilter.Apply();

            document1 = globalDocumentsList.GetGlobalDocumentListItemFromText(filename1, false);
            Assert.Null(document1, "The document from Secondary PABU is found by Primary PABU on All Documents tab.");

            globalDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilter.Controls["File Name"].Set(filename2);
            documentsFilter.Apply();

            document2 = globalDocumentsList.GetGlobalDocumentListItemFromText(filename2);
            Assert.NotNull(document2, "The document is not found by Secondary PABU and Filename on All Documents tab.");

            // Checked Out - Secondary PABU
            globalDocumentsPage.OpenCheckedOutDocumentsList();
            globalDocumentsList.OpenListOptionsMenu().RestoreDefaults();
            globalDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();

            documentsFilter.Controls["File Name"].Set(filename1);
            documentsFilter.Controls["Practice Area - Business Unit"].Set(secondaryPABU);
            documentsFilter.Apply();

            document1 = globalDocumentsList.GetGlobalDocumentListItemFromText(filename1, false);
            Assert.Null(document1, "The document from Secondary PABU is found by Primary PABU on Checked Out tab.");

            globalDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilter.Controls["File Name"].Set(filename2);
            documentsFilter.Apply();

            document2 = globalDocumentsList.GetGlobalDocumentListItemFromText(filename2);
            Assert.NotNull(document2, "The document is not found by Secondary PABU and Filename on Checked Out tab.");

            // Recent Documents - Secondary PABU
            globalDocumentsPage.OpenRecentDocumentsList();
            globalDocumentsList.OpenListOptionsMenu().RestoreDefaults();
            globalDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();

            documentsFilter.Controls["File Name"].Set(filename1);
            documentsFilter.Controls["Practice Area - Business Unit"].Set(secondaryPABU);
            documentsFilter.Apply();

            document1 = globalDocumentsList.GetGlobalDocumentListItemFromText(filename1, false);
            Assert.Null(document1, "The document from Secondary PABU is found by Primary PABU on Recent Documents tab.");

            globalDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilter.Controls["File Name"].Set(filename2);
            documentsFilter.Apply();

            document2 = globalDocumentsList.GetGlobalDocumentListItemFromText(filename2);
            Assert.NotNull(document2, "The document is not found by Secondary PABU and Filename on Recent Documents tab.");

            // Cleanup
            document2.FileOptions.DiscardCheckOutAndRemoveLocalCopy();
            document2.Delete().Confirm();
        }

        [Test]
        [Category(GlobalDocumentsTestCategory)]
        public void RecentDocumentsSort()
        {
            const int uniqueDocumentsNumber = 2;

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;

            globalDocumentsPage.Open();

            var documents = globalDocumentsPage.ItemList.GetAllGlobalDocumentListItems();

            // Test data creation
            var documentsCreated = new List<string>();
            if (documents.GroupBy(x => x.Name).Count() < uniqueDocumentsNumber)
            {
                // no documents to check sorting, upload new documents
                mattersListPage.Open();
                mattersListPage.ItemList.OpenFirst();
                matterDetailsPage.Tabs.Open("Documents");

                for (var i = 0; i < uniqueDocumentsNumber; i++)
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

            globalDocumentsPage.RecentDocumentsSortDialog.Sort("Document Size", SortOrder.Ascending);
            documents = globalDocumentsPage.ItemList.GetAllGlobalDocumentListItems();
            Assert.That(documents.GroupBy(x => x.Name).Count(), Is.GreaterThanOrEqualTo(uniqueDocumentsNumber),
                "Two or more documents with different names are required to verify sorting.");
            Assert.That(documents, Is.Ordered.Ascending.By(nameof(GlobalDocumentListItem.DocumentSize)).Using(new DocumentSizeComparer()));

            globalDocumentsPage.RecentDocumentsSortDialog.RestoreSortDefaults();
            documents = globalDocumentsPage.ItemList.GetAllGlobalDocumentListItems();
            Assert.That(documents, Is.Ordered.Descending.By(nameof(GlobalDocumentListItem.UpdatedAt)));

            // Cleanup
            foreach (var documentName in documentsCreated)
            {
                var document = globalDocumentsPage.ItemList.GetGlobalDocumentListItemFromText(documentName);
                document.Delete().Confirm();
            }
        }

        [Test]
        [Category(GlobalDocumentsTestCategory)]
        public void CheckedOutDocumentsSort()
        {
            const int uniqueDocumentsNumber = 2;

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var documentSummary = _outlook.Oc.DocumentSummaryPage;

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenCheckedOutDocumentsList();
            var documents = globalDocumentsPage.ItemList.GetAllGlobalDocumentListItems();

            // Test data creation
            var documentsCreated = new List<string>();
            if (documents.GroupBy(x => x.Name).Count() < uniqueDocumentsNumber)
            {
                // no documents to check sorting, upload new documents
                mattersListPage.Open();
                mattersListPage.ItemList.OpenFirst();
                matterDetailsPage.Tabs.Open("Documents");

                for (var i = 0; i < uniqueDocumentsNumber; i++)
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

            globalDocumentsPage.CheckedOutDocumentsSortDialog.Sort("Document Size", SortOrder.Ascending);
            documents = globalDocumentsPage.ItemList.GetAllGlobalDocumentListItems();
            Assert.That(documents.GroupBy(x => x.Name).Count(), Is.GreaterThanOrEqualTo(uniqueDocumentsNumber),
                "Two or more documents with different names are required to verify sorting.");
            Assert.That(documents, Is.Ordered.Ascending.By(nameof(GlobalDocumentListItem.DocumentSize)).Using(new DocumentSizeComparer()));

            globalDocumentsPage.CheckedOutDocumentsSortDialog.RestoreSortDefaults();
            documents = globalDocumentsPage.ItemList.GetAllGlobalDocumentListItems();
            Assert.That(documents, Is.Ordered.Descending.By(nameof(GlobalDocumentListItem.UpdatedAt)));

            // Cleanup
            if (documentsCreated.Count > 0)
            {
                var document = globalDocumentsPage.ItemList.GetGlobalDocumentListItemFromText(documentsCreated[0]);
                document.NavigateToSummary();
                documentSummary.NavigateToParentMatter();
            }

            foreach (var documentName in documentsCreated)
            {
                var document = documentsListPage.ItemList.GetGlobalDocumentListItemFromText(documentName);
                document.FileOptions.DiscardCheckOutAndRemoveLocalCopy();
                document.Delete().Confirm();
            }
        }

        [Test]
        [Category(GlobalDocumentsTestCategory)]
        public void AllDocumentsSort()
        {
            const int uniqueDocumentsNumber = 2;

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var documentListFilterDialog = globalDocumentsPage.GlobalDocumentsListFilterDialog;

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            documentsListPage.ItemList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentListFilterDialog.Apply();

            var documents = globalDocumentsPage.ItemList.GetAllGlobalDocumentListItems();

            // Test data creation
            var documentsCreated = new List<string>();
            if (documents.GroupBy(x => x.Name).Count() < uniqueDocumentsNumber)
            {
                // no documents to check sorting, upload new documents
                mattersListPage.Open();
                mattersListPage.ItemList.OpenFirst();
                matterDetailsPage.Tabs.Open("Documents");

                for (var i = 0; i < uniqueDocumentsNumber; i++)
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

            // Verify sorting
            globalDocumentsPage.AllDocumentsSortDialog.Sort("Document Size", SortOrder.Ascending);
            documents = globalDocumentsPage.ItemList.GetAllGlobalDocumentListItems();
            Assert.That(documents.GroupBy(x => x.DocumentSize).Count(), Is.GreaterThanOrEqualTo(uniqueDocumentsNumber),
                "Two or more documents with different size are required to verify sorting.");
            Assert.That(documents, Is.Ordered.Ascending.By(nameof(GlobalDocumentListItem.DocumentSize)).Using(new DocumentSizeComparer()));

            globalDocumentsPage.AllDocumentsSortDialog.RestoreSortDefaults();
            documents = globalDocumentsPage.ItemList.GetAllGlobalDocumentListItems();
            Assert.That(documents, Is.Ordered.Ascending.By(nameof(GlobalDocumentListItem.Name)));

            // Cleanup
            foreach (var documentName in documentsCreated)
            {
                var document = globalDocumentsPage.ItemList.GetGlobalDocumentListItemFromText(documentName);
                document.Delete().Confirm();
            }
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(GlobalDocumentsTestCategory)]
        [Description("17115 : View and download multiple documents versions from Document summary")]
        public void ViewAndDownloadMultipleDocumentFromVersionHistory()
        {
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var documentSummaryPage = _outlook.Oc.DocumentSummaryPage;
            var checkInDialog = documentSummaryPage.CheckInDocumentDialog;
            var versionHistoryList = documentSummaryPage.ItemList;
            var documentListPage = _outlook.Oc.DocumentsListPage;
            var documentList = documentListPage.ItemList;
            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;

            const int DocumentVersionsCreated = 2;
            const string DocumentContent = "This is a test document with version : ";

            var testDocument = CreateDocument(OfficeApp.Word, DocumentContent + "1");

            mattersListPage.Open();
            mattersList.OpenRandom();
            matterDetailsPage.Tabs.Open("Documents");

            var folderName = GetRandomText(5);
            documentList.OpenAddFolderDialog();
            documentListPage.AddFolderDialog.Controls["Name"].Set(folderName);
            documentListPage.AddFolderDialog.Save();

            var testFolder = documentList.GetMatterDocumentListItemFromText(folderName);
            Assert.IsNotNull(testFolder);
            testFolder.Open();

            DragAndDrop.FromFileSystem(testDocument, matterDetailsPage.DropPoint.GetElement());
            checkInDialog.Controls["Comments"].Set($"{AutomatedComment} : Version 1");
            documentListPage.AddDocumentDialog.UploadDocument();

            documentListPage.QuickSearch.SearchBy(testDocument.Name);
            var uploadedMatterDocument = documentList.GetMatterDocumentListItemFromText(testDocument.Name);
            Assert.IsNotNull(uploadedMatterDocument);

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.ShowResultAllDocumentsList();
            documentListPage.QuickSearch.SearchBy(testDocument.Name);
            var uploadedDocument = documentList.GetGlobalDocumentListItemFromText(testDocument.Name);
            Assert.IsNotNull(uploadedDocument);

            // Creating multiple versions of uploaded document
            for (var i = 0; i < DocumentVersionsCreated; i++)
            {
                uploadedDocument = documentList.GetGlobalDocumentListItemFromText(testDocument.Name);
                uploadedDocument.FileOptions.CheckOut();

                _word = new Word(TestEnvironment);
                _word.Attach(testDocument.Name);
                _word.ReplaceTextWith($"{DocumentContent}{i + 2}");
                _word.SaveDocument();
                _word.Close();

                uploadedDocument = documentList.GetGlobalDocumentListItemFromText(testDocument.Name);
                uploadedDocument.FileOptions.CheckIn();
                checkInDialog.Controls["Comments"].Set($"{AutomatedComment} : Version {i + 2}");
                documentListPage.AddDocumentDialog.UploadDocument();
            }

            documentListPage.QuickSearch.SearchBy(testDocument.Name);
            uploadedDocument = documentList.GetGlobalDocumentListItemFromText(testDocument.Name);
            uploadedDocument.NavigateToSummary();
            documentSummaryPage.SummaryPanel.Toggle();

            var documentVersions = versionHistoryList.GetAllListItems();
            Assert.AreEqual(DocumentVersionsCreated + 1, documentVersions.Count);

            // View and validate content of all Versions of Checked In Document
            var currentVersionOfDocument = versionHistoryList.GetListItemByIndex(0);
            Assert.IsNotNull(currentVersionOfDocument);

            for (var i = documentVersions.Count - 1; i > 0; i--)
            {
                currentVersionOfDocument = versionHistoryList.GetListItemByIndex(documentVersions.Count - i);
                currentVersionOfDocument.Open();

                _word = new Word(TestEnvironment);
                _word.Attach(testDocument.Name);

                var documentContent = _word.ReadActiveFileContent();
                Assert.AreEqual(DocumentContent + i, documentContent);

                _word.Close();
            }

            Windows.ClearWorkingTempFolder();

            // To download and validate the latest version
            var documentAllVersions = versionHistoryList.GetAllVersionHistoryListItems();
            Assert.IsNotNull(documentAllVersions[0]);

            var localFile = documentAllVersions[0].Download(testDocument.Name);

            _word = new Word(TestEnvironment);

            _word.OpenDocumentFromExplorer(localFile.FullName);
            Assert.IsTrue(_word.IsDocumentOpened);

            _word.Close();

            Assert.AreEqual(DocumentContent + documentAllVersions[0].Version, _word.ReadWordContent(localFile.FullName));

            // Clean up
            documentSummaryPage.NavigateToParentMatter();

            // Clean up
            testFolder = documentList.GetMatterDocumentListItemFromText(folderName);
            testFolder.Delete().Confirm();
        }

        [TearDown]
        public void TearDown()
        {
            SaveScreenShotsAndLogs(_outlook);
            _word?.Destroy();
            _outlook?.Destroy();
        }
    }
}
