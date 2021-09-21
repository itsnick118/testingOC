using System;
using System.IO;
using System.Linq;
using IntegratedDriver;
using NUnit.Framework;
using UITests.PageModel;
using UITests.PageModel.OfficeApps;
using static UITests.Constants;
using static UITests.TestHelpers;

namespace UITests.RegressionTesting
{
    [TestFixture]
    public class WordContextRegressionTests : UITestBase
    {
        private Word _word;
        private FileInfo _quickFileInfo;

        [SetUp]
        public void SetUp()
        {
            _word = new Word(TestEnvironment);
            _quickFileInfo = CreateDocument(OfficeApp.Word);
            _word.OpenDocumentFromExplorer(_quickFileInfo.FullName);
            _word.AttachToOc();
            _word.Oc.BasicSettingsPage.LogInAsStandardUser();
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_16378 Verify the People Tab in office App(Word)")]
        public void VerifyPeopleTabInWord()
        {
            const string expectedTab = "people";
            var testDocFile = CreateDocument(OfficeApp.Word);
            var testEmailFile = EmailGenerator.GetTestEmailTemplate();

            var mattersListPage = _word.Oc.MattersListPage;
            var matterDetails = _word.Oc.MatterDetailsPage;
            var documentsListPage = _word.Oc.DocumentsListPage;
            var peopleListPage = _word.Oc.PeopleListPage;
            var peopleList = peopleListPage.ItemList;
            var documentList = documentsListPage.ItemList;

            //Verify People Tab shown correctly
            mattersListPage.Open();
            mattersListPage.ItemList.OpenRandom();
            var selectedTab = matterDetails.Tabs.GetActiveTab().ToLower();
            Assert.AreEqual(expectedTab, selectedTab);

            //DnD Email in people
            DragAndDrop.FromFileSystem(testEmailFile, matterDetails.DropPoint.GetElement());
            var emailCount = _word.Oc.GetQueuedEmailCount();
            Assert.IsNotNull(emailCount, "No email in queue to upload");
            Assert.AreEqual(1, emailCount);

            //DnD Doc in people
            DragAndDrop.FromFileSystem(testDocFile, matterDetails.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();

            //Search Uploaded document in Documents tab
            matterDetails.Tabs.Open("Documents");
            var foundDocumentsByDocumentName = documentList.GetMatterDocumentListItemFromText(testDocFile.Name);
            Assert.IsNotNull(foundDocumentsByDocumentName, "Uploaded document does not exists");

            //Add Person in People Tab
            matterDetails.Tabs.Open("People");
            var personPIC = peopleListPage.RemoveAllPersonsExceptPIC();

            peopleList.OpenAddDialog();
            var addPersonDialog = peopleListPage.AddPersonDialog;
            addPersonDialog.Controls["Person Type"].Set("Internal");
            addPersonDialog.Controls["Person"].SetValueOtherthan(personPIC.PersonName);
            addPersonDialog.Controls["Role/Involvement Type"].SetByIndex(1);
            addPersonDialog.Controls["Start Date"].Set(FormatDate(DateTime.Now.AddDays(0)));
            addPersonDialog.Controls["End Date"].Set(FormatDate(DateTime.Now.AddDays(60)));
            addPersonDialog.Controls["Comments"].Set("comments_test");

            var personName = addPersonDialog.Controls["Person"].GetValue();
            var roleInvolvementType = addPersonDialog.Controls["Role/Involvement Type"].GetValue();
            addPersonDialog.Save();

            var addedPerson = peopleList.GetPeopleListItemFromText(personName);
            Assert.IsNotNull(addedPerson, "Newly added person is not listed after saving");
            Assert.AreEqual(true, peopleList.IsSortIconVisible);
            Assert.AreEqual(personName, addedPerson.PersonName);
            Assert.AreEqual(roleInvolvementType, addedPerson.Role);
            Assert.That(addedPerson.PersonType.StartsWith("Internal"));
            Assert.AreEqual(true, addedPerson.IsRemovePersonButtonVisible());
            Assert.AreEqual(true, addedPerson.IsEditButtonVisible());
            Assert.AreEqual(true, addedPerson.IsContactIconVisible());
            Assert.AreEqual(true, addedPerson.IsEmailIconVisible());
            Assert.AreEqual(BlackColorName, addedPerson.GetPersonNameColor().Name);

            //Edit person.
            addedPerson.Edit();
            Assert.AreEqual(addPersonDialog.GetDialogButtons(), editPopupButtons);
            Assert.AreEqual(addPersonDialog.HeaderText, "Edit Person");
            addPersonDialog.Controls["Person"].SetValueOtherthan(personPIC.PersonName, true);
            addPersonDialog.Controls["Role/Involvement Type"].SetByIndex(2, true);
            personName = addPersonDialog.Controls["Person"].GetValue();
            roleInvolvementType = addPersonDialog.Controls["Role/Involvement Type"].GetValue();

            addPersonDialog.Save();
            addedPerson = peopleList.GetPeopleListItemFromText(personName);
            Assert.AreEqual(personName, addedPerson.PersonName);
            Assert.AreEqual(roleInvolvementType, addedPerson.Role);

            //Remove Person
            addedPerson.Remove().Confirm();
            addedPerson = peopleList.GetPeopleListItemFromText(personName);
            Assert.IsNull(addedPerson, "Removed person is still visible in the list");

            //validate to check PIC should not be removed.
            personPIC = peopleListPage.RemoveAllPersonsExceptPIC();
            Assert.AreEqual(1, peopleList.GetCount());
            personPIC.Remove().Confirm();
            var messages = _word.Oc.GetAllToastMessages();
            Assert.AreEqual(1, messages.Length);
            Assert.AreEqual(MessageonTryingToRemovePIC(personPIC.PersonName), messages[0]);
            _word.Oc.CloseAllToastMessages();
            personPIC = peopleList.GetPeopleListItemByIndex(0);
            Assert.IsNotNull(personPIC);
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_16833 working with emails in other office apps(Word)")]
        public void VerifyEmailTabInWord()
        {
            const string expectedTab = "emails";
            var folderName = GetRandomText(6);
            var subFolderName = GetRandomText(7);
            var testDocFile = CreateDocument(OfficeApp.Word);
            var testEmailFile = EmailGenerator.GetTestEmailTemplate();

            var mattersListPage = _word.Oc.MattersListPage;
            var matterDetails = _word.Oc.MatterDetailsPage;
            var documentsListPage = _word.Oc.DocumentsListPage;
            var emailsListPage = _word.Oc.EmailListPage;
            var emailsList = emailsListPage.ItemList;
            var addFolderDialog = emailsListPage.AddFolderDialog;
            var documentSummary = _word.Oc.DocumentSummaryPage;
            var checkInDialog = documentSummary.CheckInDocumentDialog;
            var documentList = documentsListPage.ItemList;
            var ocHeader = _word.Oc.Header;

            //Verify Emails Tab shown correctly
            mattersListPage.Open();
            mattersListPage.ItemList.OpenRandom();
            matterDetails.Tabs.Open("Emails");
            var selectedTab = matterDetails.Tabs.GetActiveTab().ToLower();
            Assert.AreEqual(expectedTab, selectedTab);

            //DnD email in Emails Tab
            DragAndDrop.FromFileSystem(testEmailFile, matterDetails.DropPoint.GetElement());
            var emailCount = _word.Oc.GetQueuedEmailCount();
            Assert.IsNotNull(emailCount, "No email in queue to upload");
            Assert.AreEqual(1, emailCount);
            ocHeader.OpenUploadQueue();
            ocHeader.CancelAllQueued();
            Assert.IsFalse(ocHeader.IsUploadEmailWaitingQueueDisplayed(), "Upload email waiting queue is not cleared");

            //DnD Doc in Emails Tab and verify in Document Tab
            DragAndDrop.FromFileSystem(testDocFile, matterDetails.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();

            matterDetails.Tabs.Open("Documents");
            var foundDocumentsByDocumentName = documentList.GetMatterDocumentListItemFromText(testDocFile.Name);
            Assert.IsNotNull(foundDocumentsByDocumentName, "No document exists");
            foundDocumentsByDocumentName.Delete().Confirm();
            foundDocumentsByDocumentName = documentList.GetMatterDocumentListItemFromText(testDocFile.Name, false);
            Assert.IsNull(foundDocumentsByDocumentName, "Document still exists");

            //Verify that user can add folder and sub folder under Emails tab
            matterDetails.Tabs.Open("Emails");
            emailsList.OpenAddFolderDialog();
            addFolderDialog.Controls["Name"].Set(folderName);
            addFolderDialog.Save();

            var foundFolderByFolderName = emailsList.GetEmailListItemFromText(folderName);
            Assert.IsTrue(foundFolderByFolderName.IsFolder);
            foundFolderByFolderName.Open();
            emailsList.OpenAddFolderDialog();
            addFolderDialog.Controls["Name"].Set(subFolderName);
            addFolderDialog.Save();

            var foundSubFolderBySubFolderName = emailsList.GetEmailListItemFromText(subFolderName);
            Assert.IsTrue(foundSubFolderBySubFolderName.IsFolder);

            //Verify that Quick File is not visible under Emails Tab for folders
            matterDetails.Tabs.Open("Emails");
            foundFolderByFolderName = emailsList.GetEmailListItemFromText(folderName);
            Assert.IsFalse(foundFolderByFolderName.HasQuickFileButton);
            foundFolderByFolderName.Open();
            foundSubFolderBySubFolderName = emailsList.GetEmailListItemFromText(subFolderName);
            Assert.IsFalse(foundSubFolderBySubFolderName.HasQuickFileButton);

            //Verify that "Quick file" on matter summary uploads current document under documents tab
            matterDetails.QuickFile();
            checkInDialog.UploadDocument();
            matterDetails.Tabs.Open("Documents");
            var foundQuickFileDocumentByDocumentName = documentList.GetMatterDocumentListItemFromText(_quickFileInfo.Name);
            Assert.IsNotNull(foundQuickFileDocumentByDocumentName);
            foundQuickFileDocumentByDocumentName.Delete().Confirm();
            foundQuickFileDocumentByDocumentName = documentList.GetMatterDocumentListItemFromText(_quickFileInfo.Name, false);
            Assert.IsNull(foundQuickFileDocumentByDocumentName, "Document still exists");

            //Cleanup test data
            matterDetails.Tabs.Open("Emails");
            foundFolderByFolderName = emailsList.GetEmailListItemFromText(folderName);
            foundFolderByFolderName.Delete().Confirm();
            foundFolderByFolderName = emailsList.GetEmailListItemFromText(folderName, false);
            Assert.IsNull(foundFolderByFolderName, "Folder still exists");
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_16174  GDL: Verify Document Summary Page / Version History reloads properly")]
        public void VerifyDocumentSummaryPageVersionHistoryReload()
        {
            var mattersListPage = _word.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetails = _word.Oc.MatterDetailsPage;
            var documentsListPage = _word.Oc.DocumentsListPage;
            var documentSummaryPage = _word.Oc.DocumentSummaryPage;
            var checkInDialog = documentSummaryPage.CheckInDocumentDialog;

            //Step 1
            var documentsList = documentsListPage.ItemList;
            var globalDocumentsPage = _word.Oc.GlobalDocumentsPage;
            var globalDocumentsList = globalDocumentsPage.ItemList;

            // Open matter documents sub tab
            mattersList.OpenRandom();
            matterDetails.Tabs.Open("Documents");

            // Add new document
            var dndFileInfo = CreateDocument(OfficeApp.Word);
            DragAndDrop.FromFileSystem(dndFileInfo, matterDetails.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();

            var toastMessage = _word.Oc.GetAllToastMessages();
            Assert.AreEqual(UploadSuccessMessage, toastMessage[0]);
            _word.Oc.CloseAllToastMessages();

            var uploadedDocument = documentsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            Assert.IsNotNull(uploadedDocument);

            // Go to global documents app then open all documents list page
            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.QuickSearch.SearchBy(dndFileInfo.Name);
            var addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(dndFileInfo.Name);
            Assert.IsNotNull(addedDocument);
            Assert.AreEqual(1, globalDocumentsList.GetCount());

            // verify document summary
            addedDocument.NavigateToSummary();
            var summaryInfo = documentSummaryPage.GetDocumentSummaryInfo();
            Assert.IsNotEmpty(summaryInfo, "Document Summary fields are not retrieved or empty.");

            //Step 2
            _word.Oc.ReloadOc();
            summaryInfo = documentSummaryPage.GetDocumentSummaryInfo();
            Assert.IsNotEmpty(summaryInfo, "Document Summary fields are not retrieved or empty.");
            foreach (var field in summaryInfo)
                Assert.IsNotEmpty(field.Text);

            //Step 4
            documentSummaryPage.SummaryPanel.Toggle();
            var fileName = addedDocument.Name;
            var versionHistoryList = documentSummaryPage.ItemList;
            var file = versionHistoryList.GetVersionHistoryListItemByIndex(0);
            file.Open();
            _word = new Word(TestEnvironment);
            _word.Attach(fileName);
            Assert.IsNotNull(_word.GetReadOnlyLabel());

            //Step 5
            _word.CheckOut();
            _word.CloseDocument();
            _word.ReplaceTextWith(GetRandomText(10));
            _word.SaveDocument();
            documentSummaryPage.QuickFile();
            checkInDialog.Proceed();
            checkInDialog.UploadDocument();
            documentSummaryPage.SummaryPanel.Toggle();

            // check in
            var checkedIn = CheckInStatus.CheckedIn.ToLower();
            var documentStatus = documentSummaryPage.IsStatusEqualsTo(checkedIn);
            Assert.IsTrue(documentStatus);

            // delete document
            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.QuickSearch.SearchBy(dndFileInfo.Name);

            addedDocument = globalDocumentsList.GetGlobalDocumentListItemByIndex(0);
            addedDocument.Delete().Confirm();
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC17294 -  Verify to perform download and view operation from version history")]
        public void DownloadandViewFromVersionHistory()
        {
            var documentSummaryPage = _word.Oc.DocumentSummaryPage;
            var globalDocumentsPage = _word.Oc.GlobalDocumentsPage;
            var globalDocumentsList = globalDocumentsPage.ItemList;
            var checkInDocumentDialog = globalDocumentsPage.CheckInDocumentDialog;
            var versionHistoryList = documentSummaryPage.ItemList;
            // close initially generated random document
            _word.CloseDocument();

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();

            globalDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            var documentsFilterDialog = globalDocumentsPage.GlobalDocumentsListFilterDialog;
            documentsFilterDialog.Controls["Status"].Set(CheckInStatus.CheckedIn);
            documentsFilterDialog.Controls["File Name"].Set(".doc");
            documentsFilterDialog.Apply();

            var unfilteredCount = globalDocumentsList.GetCount();
            var randomDocumentIndex = GetRandomNumber(unfilteredCount - 1);

            var selectedDocument = globalDocumentsList.GetGlobalDocumentListItemByIndex(randomDocumentIndex);
            selectedDocument.FileOptions.CheckOut();

            var randomString = GetRandomText(10);
            _word.ReplaceTextWith(randomString);
            _word.SaveDocument();
            _word.CloseDocument();

            globalDocumentsPage.OpenCheckedOutDocumentsList();
            var checkedOutDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(selectedDocument.Name);
            checkedOutDocument.FileOptions.CheckIn();
            checkInDocumentDialog.Controls["Comments"].Set("Test Comment - Document Version 0");
            checkInDocumentDialog.UploadDocument();

            globalDocumentsPage.OpenAllDocumentsList();
            globalDocumentsPage.QuickSearch.SearchBy(selectedDocument.Name);
            var checkedInDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(selectedDocument.Name);
            checkedInDocument.NavigateToSummary();
            documentSummaryPage = _word.Oc.DocumentSummaryPage;
            documentSummaryPage.SummaryPanel.Toggle();

            // Verify versions history list and sort order
            var versionsUpdated = versionHistoryList.GetAllVersionHistoryListItems()
                                    .Select(x => x.Version).ToList();
            var descendingListTemplate = versionsUpdated.OrderByDescending(x => x).ToList();
            Assert.AreEqual(versionsUpdated, descendingListTemplate);

            // Step 1 : Verify to download specific version of document from version history

            var latestVersionOfDocument = versionHistoryList.GetVersionHistoryListItemByIndex(0);
            Assert.IsNotNull(latestVersionOfDocument);
            latestVersionOfDocument.Download(selectedDocument.Name);
            var fileFullPath = Path.Combine(Windows.GetWorkingTempFolder().FullName, selectedDocument.Name);
            var fileContent = _word.ReadWordContent(fileFullPath);
            Assert.AreEqual(randomString.ToLower(), fileContent.ToLower());

            // Step 2 : Verify to download specific version of document from version history
            //          (checked out by different user)
            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();

            globalDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog = globalDocumentsPage.GlobalDocumentsListFilterDialog;
            documentsFilterDialog.Controls["Status"].Set(CheckInStatus.CheckedOut);
            documentsFilterDialog.Controls["File Name"].Set(".doc");
            documentsFilterDialog.Controls["Updated By"].Set("super user");
            documentsFilterDialog.Apply();

            var checkedOutDocumentBySuperUser = globalDocumentsList.GetGlobalDocumentListItemByIndex(0);
            checkedOutDocumentBySuperUser.NavigateToSummary();
            var documentAllVersions = versionHistoryList.GetAllVersionHistoryListItems();
            Assert.IsNotNull(documentAllVersions[0]);
            documentAllVersions[0].Download(checkedOutDocumentBySuperUser.Name);

            var toastMessages = _word.Oc.GetAllToastMessages();
            if (toastMessages != null)
            {
                Assert.AreEqual(toastMessages[0], ViewDocumentWarningMessage("suser"));
            }

            // Step 3 :  Verify to view a document from version history list page

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();

            selectedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(selectedDocument.Name);
            selectedDocument.NavigateToSummary();
            documentAllVersions = versionHistoryList.GetAllVersionHistoryListItems();
            Assert.IsNotNull(documentAllVersions[0]);
            // row click operation
            documentAllVersions[0].Open();

            // Step 4 : Verify to view a document from version history list page
            //          (Checked out by other user)
            globalDocumentsPage.Open();
            globalDocumentsPage.OpenAllDocumentsList();

            globalDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog = globalDocumentsPage.GlobalDocumentsListFilterDialog;
            documentsFilterDialog.Controls["Status"].Set(CheckInStatus.CheckedOut);
            documentsFilterDialog.Controls["File Name"].Set(".doc");
            documentsFilterDialog.Controls["Updated By"].Set("super user");
            documentsFilterDialog.Apply();

            checkedOutDocumentBySuperUser = globalDocumentsList.GetGlobalDocumentListItemByIndex(0);
            checkedOutDocumentBySuperUser.NavigateToSummary();

            documentAllVersions = versionHistoryList.GetAllVersionHistoryListItems();
            Assert.IsNotNull(documentAllVersions[0]);
            // row click operation
            documentAllVersions[0].Open();

            toastMessages = _word.Oc.GetAllToastMessages();
            if (toastMessages != null)
            {
                Assert.AreEqual(toastMessages[0], ViewDocumentWarningMessage("suser"));
            }
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Editing Documents - QuickFile")]
        public void QuickFileDocument()
        {
            var mattersListPage = _word.Oc.MattersListPage;
            var matterDetails = _word.Oc.MatterDetailsPage;
            var globalDocuments = _word.Oc.GlobalDocumentsPage;
            var documentSummary = _word.Oc.DocumentSummaryPage;
            var checkInDialog = documentSummary.CheckInDocumentDialog;

            // Upload a document
            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetails.Tabs.Open("Documents");

            // Quick Filing version 1 of document
            matterDetails.QuickFile();
            checkInDialog.Controls["Comments"].Set("Document: QuickFile : Check In Operation: Version 1");
            checkInDialog.UploadDocument();

            _word.CloseDocument();

            // Verify scenario
            globalDocuments.Open();
            globalDocuments.QuickSearch.SearchBy(_quickFileInfo.Name);

            var testDocument = globalDocuments.ItemList.GetGlobalDocumentListItemFromText(_quickFileInfo.Name);
            testDocument.FileOptions.CheckOut();
            testDocument = globalDocuments.ItemList.GetGlobalDocumentListItemFromText(_quickFileInfo.Name);
            testDocument.NavigateToSummary();
            Assert.AreEqual(documentSummary.ItemList.GetCount(), 1, "There are more or less than 1 version of the document.");

            _word.ReplaceTextWith(GetRandomText(10));
            // Quick Filing version 2 of document
            documentSummary.QuickFile();
            //Save
            checkInDialog.Save();
            checkInDialog.Controls["Comments"].Set("Document: QuickFile : Check In Operation: Version 2");
            checkInDialog.UploadDocument();

            var toastMessage = _word.Oc.GetAllToastMessages();
            Assert.AreEqual(DocumentSummaryQuickFileMessage, toastMessage[0]);

            documentSummary.WaitForStatusChangeTo(CheckInStatus.CheckedIn);
            Assert.AreEqual(documentSummary.ItemList.GetCount(), 2, "There are more or less than 2 versions of the document.");

            _word.Oc.Header.NavigateBack();
            globalDocuments.QuickSearch.SearchBy(_quickFileInfo.Name);
            testDocument = globalDocuments.ItemList.GetGlobalDocumentListItemFromText(_quickFileInfo.Name);
            testDocument.FileOptions.CheckOut();
            testDocument = globalDocuments.ItemList.GetGlobalDocumentListItemFromText(_quickFileInfo.Name);
            testDocument.NavigateToSummary();
            _word.ReplaceTextWith(GetRandomText(10));
            documentSummary.QuickFile();
            //Save As
            checkInDialog.SaveAs(_quickFileInfo.Name, false);
            checkInDialog.Controls["Comments"].Set("Document: QuickFile : Check In Operation: Version 3");
            checkInDialog.UploadDocument();

            // Cleanup
            _word.Oc.Header.NavigateBack();
            globalDocuments.QuickSearch.SearchBy(_quickFileInfo.Name);
            testDocument = globalDocuments.ItemList.GetGlobalDocumentListItemFromText(_quickFileInfo.Name);
            testDocument.Delete().Confirm();
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Editing Documents - DiscardCheckout")]
        public void DiscardCheckOutDocument()
        {
            var mattersListPage = _word.Oc.MattersListPage;
            var matterDetails = _word.Oc.MatterDetailsPage;
            var globalDocumentsPage = _word.Oc.GlobalDocumentsPage;
            var globalDocumentsList = globalDocumentsPage.ItemList;
            var checkInDialog = globalDocumentsPage.CheckInDocumentDialog;
            var checkOutDocumentDialog = globalDocumentsPage.CheckOutDocumentDialog;

            // Upload a document
            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetails.Tabs.Open("Documents");

            // Quick Filing version 1 of document
            matterDetails.QuickFile();
            checkInDialog.Controls["Comments"].Set("Test Document: QuickFile : Check In");
            checkInDialog.UploadDocument();
            _word.CloseDocument();

            // Verify scenario
            globalDocumentsPage.Open();
            globalDocumentsPage.QuickSearch.SearchBy(_quickFileInfo.Name);
            var testDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(_quickFileInfo.Name);

            Assert.NotNull(testDocument);
            Assert.AreEqual(CheckInStatus.CheckedIn.ToLower(), testDocument.Status.ToLower());

            testDocument.FileOptions.CheckOut();
            testDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(_quickFileInfo.Name);

            Assert.NotNull(testDocument);
            Assert.AreEqual(CheckInStatus.CheckedOut.ToLower(), testDocument.Status.ToLower());

            testDocument.FileOptions.DiscardCheckOut();
            checkOutDocumentDialog.Keep();

            // Clean up
            _word.CloseDocument();
            testDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(_quickFileInfo.Name);
            testDocument.Delete().Confirm();
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("TC-17152 part 2 : View and download specific document versions from Document summary")]
        public void ViewOtherUsersDocFromSummary()
        {
            var settingsPage = _word.Oc.SettingsPage;

            // Log out standard user from office companion
            _word.Oc.OpenSettings();
            settingsPage.OpenConfiguration();
            settingsPage.LogOut().Confirm();

            // Log out Attorney user from office companion
            _word.Oc.BasicSettingsPage.LogInAsAttorneyUser();

            var mattersListPage = _word.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetailsPage = _word.Oc.MatterDetailsPage;
            var documentSummaryPage = _word.Oc.DocumentSummaryPage;
            var checkInDialog = documentSummaryPage.CheckInDocumentDialog;
            var documentListPage = _word.Oc.DocumentsListPage;
            var documentList = documentListPage.ItemList;
            var versionHistoryList = documentSummaryPage.ItemList;
            var documentSummary = _word.Oc.DocumentSummaryPage;

            var testDocument = CreateDocument(OfficeApp.Word);
            mattersListPage.Open();
            mattersList.OpenFirst();
            matterDetailsPage.Tabs.Open("Documents");

            var folderName = GetRandomText(5);
            documentList.OpenAddFolderDialog();
            documentListPage.AddFolderDialog.Controls["Name"].Set(folderName);
            documentListPage.AddFolderDialog.Save();

            var testFolder = documentList.GetMatterDocumentListItemFromText(folderName);
            Assert.IsNotNull(testFolder);

            testFolder.Open();

            DragAndDrop.FromFileSystem(testDocument, matterDetailsPage.DropPoint.GetElement());
            checkInDialog.UploadDocument();

            documentListPage.QuickSearch.SearchBy(testDocument.Name);
            var uploadedDocument = documentList.GetMatterDocumentListItemFromText(testDocument.Name);

            Assert.AreEqual(CheckInStatus.CheckedIn.ToLower(), uploadedDocument.Status.ToLower());

            _word.CloseDocument();

            uploadedDocument.FileOptions.CheckOut();

            _word.Oc.OpenSettings();
            settingsPage.OpenConfiguration();
            settingsPage.LogOut().Confirm();

            // Log in as standard user from office companion
            _word.Oc.BasicSettingsPage.LogInAsStandardUser();

            documentListPage = _word.Oc.DocumentsListPage;
            documentList = documentListPage.ItemList;

            mattersListPage.Open();
            mattersList.OpenFirst();
            matterDetailsPage.Tabs.Open("Documents");
            documentListPage.QuickSearch.SearchBy(folderName);
            testFolder = documentList.GetMatterDocumentListItemFromText(folderName);
            testFolder.Open();

            documentListPage.QuickSearch.SearchBy(testDocument.Name);
            uploadedDocument = documentList.GetMatterDocumentListItemFromText(testDocument.Name);
            uploadedDocument.NavigateToSummary();

            var documentVersions = versionHistoryList.GetAllListItems();
            Assert.AreEqual(1, documentVersions.Count);
            _word.CloseDocument();

            // View Checked out Document by other user
            var latestVersionOfDocument = versionHistoryList.GetListItemByIndex(0);
            latestVersionOfDocument.Open();

            // Need to change readonly label ,as we are checking out from dmaxwell and sbrown is hardcoded
            Assert.True(_word.IsReadOnly);

            documentSummary.NavigateToParentMatter();
            testFolder = documentList.GetMatterDocumentListItemFromText(folderName);
            testFolder.Delete().Confirm();

            var toastMessages = _word.Oc.GetAllToastMessages();
            Assert.AreEqual(DeleteFolderHasCheckedOutDocumentErrorMessage, toastMessages[0]);

            testFolder.Open();

            uploadedDocument = documentList.GetMatterDocumentListItemFromText(testDocument.Name);
            uploadedDocument.Delete().Confirm();

            toastMessages = _word.Oc.GetAllToastMessages();
            Assert.AreEqual(DeleteDocumentMessageForDifferentUser(TestEnvironment.AttorneyUser), toastMessages[0]);
        }

        [TearDown]
        public void TearDown()
        {
            SaveScreenShotsAndLogs(_word);
            _word.Close();
            _word.Destroy();
        }
    }
}
