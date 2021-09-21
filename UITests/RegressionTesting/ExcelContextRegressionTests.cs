using System;
using System.IO;
using IntegratedDriver;
using NUnit.Framework;
using UITests.PageModel;
using UITests.PageModel.OfficeApps;
using static UITests.Constants;
using static UITests.TestHelpers;

namespace UITests.RegressionTesting
{
    [TestFixture]
    public class ExcelContextRegressionTests : UITestBase
    {
        private Excel _excel;
        private FileInfo _quickFileInfo;

        [SetUp]
        public void SetUp()
        {
            _excel = new Excel(TestEnvironment);
            _quickFileInfo = CreateDocument(OfficeApp.Excel);
            _excel.OpenDocumentFromExplorer(_quickFileInfo.FullName);
            _excel.AttachToOc();
            _excel.Oc.BasicSettingsPage.LogInAsStandardUser();
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_16378 Verify the People Tab in office App(Excel)")]
        public void VerifyPeopleTabInExcel()
        {
            const string expectedTab = "people";
            var testDocFile = CreateDocument(OfficeApp.Word);
            var testEmailFile = EmailGenerator.GetTestEmailTemplate();

            var mattersListPage = _excel.Oc.MattersListPage;
            var matterDetails = _excel.Oc.MatterDetailsPage;
            var documentsListPage = _excel.Oc.DocumentsListPage;
            var peopleListPage = _excel.Oc.PeopleListPage;
            var peopleList = peopleListPage.ItemList;
            var documentList = documentsListPage.ItemList;

            //Verify People Tab shown correctly
            mattersListPage.Open();
            mattersListPage.ItemList.OpenRandom();
            var selectedTab = matterDetails.Tabs.GetActiveTab().ToLower();
            Assert.AreEqual(expectedTab, selectedTab);

            //DnD Email in people
            DragAndDrop.FromFileSystem(testEmailFile, matterDetails.DropPoint.GetElement());
            var emailCount = _excel.Oc.GetQueuedEmailCount();
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
            var messages = _excel.Oc.GetAllToastMessages();
            Assert.AreEqual(1, messages.Length);
            Assert.AreEqual(MessageonTryingToRemovePIC(personPIC.PersonName), messages[0]);
            _excel.Oc.CloseAllToastMessages();
            personPIC = peopleList.GetPeopleListItemByIndex(0);
            Assert.IsNotNull(personPIC);
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_16833 working with emails in other office apps(Excel)")]
        public void VerifyEmailTabInExcel()
        {
            const string expectedTab = "emails";
            var folderName = GetRandomText(6);
            var subFolderName = GetRandomText(7);
            var testDocFile = CreateDocument(OfficeApp.Word);
            var testEmailFile = EmailGenerator.GetTestEmailTemplate();

            var mattersListPage = _excel.Oc.MattersListPage;
            var matterDetails = _excel.Oc.MatterDetailsPage;
            var documentsListPage = _excel.Oc.DocumentsListPage;
            var emailsListPage = _excel.Oc.EmailListPage;
            var emailsList = emailsListPage.ItemList;
            var addFolderDialog = emailsListPage.AddFolderDialog;
            var documentSummary = _excel.Oc.DocumentSummaryPage;
            var checkInDialog = documentSummary.CheckInDocumentDialog;
            var documentList = documentsListPage.ItemList;
            var ocHeader = _excel.Oc.Header;

            //Verify Emails Tab shown correctly
            mattersListPage.Open();
            mattersListPage.ItemList.OpenRandom();
            matterDetails.Tabs.Open("Emails");
            var selectedTab = matterDetails.Tabs.GetActiveTab().ToLower();
            Assert.AreEqual(expectedTab, selectedTab);

            //DnD email in Emails Tab
            DragAndDrop.FromFileSystem(testEmailFile, matterDetails.DropPoint.GetElement());
            var emailCount = _excel.Oc.GetQueuedEmailCount();
            Assert.IsNotNull(emailCount, "No email in queue to upload");
            Assert.AreEqual(1, emailCount);
            ocHeader.OpenUploadQueue();
            ocHeader.CancelAllQueued();
            Assert.That(ocHeader.IsUploadEmailWaitingQueueDisplayed(), Is.False, "Upload email waiting queue is not cleared");

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

            //Verify that "Quick file" on the matter summary uploads current document under documents tab
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

        [TearDown]
        public void TearDown()
        {
            SaveScreenShotsAndLogs(_excel);
            _excel.Close();
            _excel.Destroy();
        }
    }
}
