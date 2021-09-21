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
    public class MatterSmokeTests : UITestBase
    {
        private Outlook _outlook;
        private Word _word;

        [SetUp]
        public void SetUp()
        {
            _outlook = new Outlook(TestEnvironment);
            _outlook.Launch();
            _outlook.Oc.BasicSettingsPage.LogIn();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(DataDependentTestCategory)]
        public void FavoritesMattersList()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;

            mattersListPage.SetNthMatterAsFavorite(0);
            mattersListPage.SetNthMatterAsFavorite(1);

            mattersListPage.OpenFavoritesList();
            var favoritesMatterCount = mattersList.GetCount();
            Assert.AreEqual(2, favoritesMatterCount);

            mattersListPage.ClearFavorites(2);
            Assert.That(mattersList.GetCount(), Is.Zero, "Favorite list has matters after removing");
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(DataDependentTestCategory)]
        public void AllMattersSort()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            mattersListPage.Open();

            mattersListPage.MatterSortDialog.Sort("Full Name", SortOrder.Descending);

            var matterItems = mattersList.GetAllMatterListItems();
            Assert.That(matterItems.Count, Is.GreaterThanOrEqualTo(2), "No matters in the list to check sorting. Need two matters at least.");
            Assert.That(matterItems.Select(x => x.PrimaryInternalContact).All(x => !string.IsNullOrEmpty(x)));
            Assert.That(matterItems, Is.Ordered.Descending.By(nameof(MatterListItem.PrimaryInternalContact)));

            mattersListPage.MatterSortDialog.RestoreSortDefaults();

            matterItems = mattersList.GetAllMatterListItems();
            Assert.That(matterItems.Select(x => x.Name).All(x => !string.IsNullOrEmpty(x)));
            Assert.That(matterItems, Is.Ordered.Ascending.By(nameof(MatterListItem.Name)));
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(DataDependentTestCategory)]
        public void MyMattersSort()
        {
            var myMattersListPage = _outlook.Oc.MattersListPage;
            myMattersListPage.OpenMyMattersList();

            Assert.Warn("Sort options not shown http://mingle/projects/growth/cards/20624");

            myMattersListPage.MyMattersSortDialog.Sort("Matter Name", SortOrder.Descending);

            var matterItems = myMattersListPage.ItemList.GetAllMatterListItems();
            Assert.That(matterItems.Count, Is.GreaterThanOrEqualTo(2), "No matters in the list to check sorting. Need two matters at least.");
            Assert.That(matterItems.Select(x => x.Name).All(x => !string.IsNullOrEmpty(x)));
            Assert.That(matterItems, Is.Ordered.Descending.By(nameof(MatterListItem.Name)));

            myMattersListPage.MyMattersSortDialog.RestoreSortDefaults();

            matterItems = myMattersListPage.ItemList.GetAllMatterListItems();
            Assert.That(matterItems.Select(x => x.StatusDate).All(x => x.HasValue));
            Assert.That(matterItems, Is.Ordered.Descending.By(nameof(MatterListItem.StatusDate)));
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(DataDependentTestCategory)]
        public void PeopleSort()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var peopleListPage = _outlook.Oc.PeopleListPage;
            var peopleList = peopleListPage.ItemList;

            mattersListPage.ItemList.OpenRandom();
            matterDetails.Tabs.Open("People");

            var peopleItems = peopleList.GetAllPeopleListItems();
            var uniquePeopleCount = peopleItems.GroupBy(x => x.PersonName).Count();
            while (uniquePeopleCount < 2)
            {
                var addPersonDialog = peopleListPage.AddPersonDialog;
                peopleList.OpenAddDialog();

                addPersonDialog.Controls["Person Type"].SetByIndex(3);
                addPersonDialog.Controls["Comments"].Set($"comments_test + {GetRandomNumber(3)}");
                addPersonDialog.Controls["Role/Involvement Type"].SetByIndex(GetRandomNumber(6, 1));
                addPersonDialog.Controls["Person"].SetByIndex(GetRandomNumber(14, 1));
                addPersonDialog.Controls["Start Date"].Set(FormatDate(DateTime.Now));
                addPersonDialog.Controls["End Date"].Set(FormatDate(DateTime.Now.AddDays(GetRandomNumber(60, 1))));
                addPersonDialog.Save();

                peopleItems = peopleList.GetAllPeopleListItems();
                uniquePeopleCount = peopleItems.GroupBy(x => x.PersonName).Count();
            }

            peopleListPage.PeopleSortDialog.Sort("Person", SortOrder.Descending);

            peopleItems = peopleList.GetAllPeopleListItems();
            Assert.That(peopleItems, Is.Ordered.Descending.By(nameof(PeopleListItem.PersonName)));

            peopleListPage.PeopleSortDialog.RestoreSortDefaults();

            peopleItems = peopleList.GetAllPeopleListItems();
            Assert.That(peopleItems, Is.Ordered.Ascending.By(nameof(PeopleListItem.PersonName)));
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(DataDependentTestCategory)]
        public void TasksEventsSort()
        {
            const string dateTimeFormat = "M-dd-yyyy h-mm-ss tt";
            const string typeEvent = "Event";
            const string typeTask = "Task";
            var eventSubjectText = Guid.NewGuid().ToString();
            var taskNameText = Guid.NewGuid().ToString();
            var taskDescription = $"Task generated at {DateTime.Now.ToString(dateTimeFormat)}";
            var eventDescription = $"Event generated at {DateTime.Now.ToString(dateTimeFormat)}";

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var tasksEventsListPage = _outlook.Oc.TasksEventsListPage;
            var tasksEventsList = tasksEventsListPage.ItemList;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetails.Tabs.Open("Tasks/Events");

            // Add Events with event dates
            tasksEventsList.OpenAddDialog();
            tasksEventsListPage.AddEventDialog.Controls["Type"].Set(typeEvent);
            tasksEventsListPage.AddEventDialog.Controls["Subject"].Set(eventSubjectText);
            tasksEventsListPage.AddEventDialog.Controls["Start Date/Time"].Set(FormatDateTime(DateTime.Now));
            tasksEventsListPage.AddEventDialog.Controls["End Date/Time"].Set(FormatDateTime(DateTime.Now));
            tasksEventsListPage.AddEventDialog.Controls["Description"].Set(eventDescription);
            tasksEventsListPage.AddEventDialog.Save();

            //Add Tasks with task dates
            tasksEventsList.OpenAddDialog();
            tasksEventsListPage.AddTaskDialog.Controls["Type"].Set(typeTask);
            tasksEventsListPage.AddTaskDialog.Controls["Name"].Set(taskNameText);
            tasksEventsListPage.AddTaskDialog.Controls["Due Date"].Set(FormatDate(DateTime.Now));
            tasksEventsListPage.AddTaskDialog.Controls["Description"].Set(taskDescription);
            tasksEventsListPage.AddTaskDialog.Controls["Invitees/Assigned To"].Set("Sally Brown");
            tasksEventsListPage.AddTaskDialog.Save();

            tasksEventsListPage.TasksEventsSortDialog.Sort("Name", SortOrder.Descending);

            var tasksEventsItems = tasksEventsList.GetAllTasksEventsListItems();
            Assert.That(tasksEventsItems.Count, Is.GreaterThanOrEqualTo(2), "No items added in the list to check sorting. Need two items at least.");
            Assert.That(tasksEventsItems, Is.Ordered.Descending.By(nameof(TasksEventsListItem.Name)));

            tasksEventsListPage.TasksEventsSortDialog.RestoreSortDefaults();
            tasksEventsItems = tasksEventsList.GetAllTasksEventsListItems();
            Assert.That(tasksEventsItems, Is.Ordered.Ascending.By(nameof(TasksEventsListItem.Type)));

            // Cleanup
            matterDetails.ItemList.GetTasksEventsListItemFromText(eventSubjectText).Delete().Confirm();
            matterDetails.ItemList.GetTasksEventsListItemFromText(taskNameText).Delete().Confirm();
        }

        [Test]
        [Category(SmokeTestCategory)]
        public void EmailsSort()
        {
            var folderNames = new[] { GetRandomText(5), GetRandomText(10) };

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var emailsListPage = _outlook.Oc.EmailListPage;
            var emailsList = emailsListPage.ItemList;
            var addFolderDialog = emailsListPage.AddFolderDialog;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetails.Tabs.Open("Emails");

            // Upload emails and add folders
            var emailsToUpload = _outlook.AddTestEmailsToFolder(2, useDifferentTemplates: true);
            _outlook.OpenTestEmailFolder();
            _outlook.Oc.WaitForLoadComplete();
            _outlook.SelectAllItems();
            matterDetails.QuickFile();

            foreach (var folderName in folderNames)
            {
                emailsList.OpenAddFolderDialog();
                addFolderDialog.Controls["Name"].Set(folderName);
                addFolderDialog.Save();
            }

            // Verify scenario
            emailsListPage.EmailsSortDialog.Sort("From", SortOrder.Descending);

            var emails = emailsList.GetAllEmailListItems();
            Assert.That(emails.Where(x => !x.IsFolder).GroupBy(x => x.From).Count(), Is.GreaterThanOrEqualTo(2),
                "There are no emails or all emails from the same sender on the list. Need emails from different senders to check sorting.");
            Assert.That(emails.Count(x => x.IsFolder), Is.GreaterThanOrEqualTo(2),
                "There are not enough folders to verify folders sorting. Need two folders at least.");
            Assert.That(emails,
                Is.Ordered.Ascending.By(nameof(EmailListItem.IsFolder)).Using(new BooleanInverterComparer())
                    .Then.Descending.By(nameof(EmailListItem.From)));

            emailsListPage.EmailsSortDialog.RestoreSortDefaults();

            emails = emailsList.GetAllEmailListItems();
            Assert.That(emails,
                Is.Ordered.Ascending.By(nameof(EmailListItem.IsFolder)).Using(new BooleanInverterComparer())
                    .Then.Descending.By(nameof(EmailListItem.ReceivedTime)));
            Assert.That(emails.Where(x => x.IsFolder).ToList(),
                Is.Ordered.Ascending.By(nameof(EmailListItem.FolderName)));

            // Cleanup
            foreach (var uploadedEmail in emailsToUpload)
            {
                matterDetails.ItemList.GetEmailListItemFromText(uploadedEmail.Key).Delete().Confirm();
            }

            foreach (var folderName in folderNames)
            {
                matterDetails.ItemList.GetEmailListItemFromText(folderName).Delete().Confirm();
            }
        }

        [Test]
        [Category(SmokeTestCategory)]
        public void NarrativesSort()
        {
            const string type1 = "Note";
            const string type2 = "Status";
            var description1 = GetRandomText(20);
            var description2 = GetRandomText(30);
            var date1 = FormatDateTime(DateTime.Now);
            var date2 = FormatDateTime(DateTime.Now.AddMinutes(1));

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var narrativesListPage = _outlook.Oc.NarrativesListPage;
            var narrativesList = narrativesListPage.ItemList;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetails.Tabs.Open("Narratives");

            // Add narratives with narrative dates
            narrativesList.OpenAddDialog();
            narrativesListPage.AddNarrativeDialog.Controls["Narrative Type"].Set(type1);
            narrativesListPage.AddNarrativeDialog.Controls["Description"].Set(description1);
            narrativesListPage.AddNarrativeDialog.Controls["Narrative Date"].Set(date1);
            narrativesListPage.AddNarrativeDialog.Save();

            narrativesList.OpenAddDialog();
            narrativesListPage.AddNarrativeDialog.Controls["Narrative Type"].Set(type2);
            narrativesListPage.AddNarrativeDialog.Controls["Description"].Set(description2);
            narrativesListPage.AddNarrativeDialog.Controls["Narrative Date"].Set(date2);
            narrativesListPage.AddNarrativeDialog.Save();

            // Verify scenario
            narrativesListPage.NarrativeSortDialog.Sort("Description", SortOrder.Descending);

            var narrativeItems = narrativesList.GetAllNarrativeListItems();
            Assert.That(narrativeItems.Count, Is.GreaterThanOrEqualTo(2), "No narratives in the list to check sorting. Need two narratives at least.");
            Assert.That(narrativeItems.Select(x => x.Description).All(x => !string.IsNullOrEmpty(x)));
            Assert.That(narrativeItems, Is.Ordered.Descending.By(nameof(NarrativeListItem.Description)));

            narrativesListPage.NarrativeSortDialog.RestoreSortDefaults();

            narrativeItems = narrativesList.GetAllNarrativeListItems();
            Assert.That(narrativeItems.Select(x => x.NarrativeDate).Any(x => x.HasValue));
            Assert.That(narrativeItems, Is.Ordered.Descending.By(nameof(NarrativeListItem.NarrativeDate)));

            // Cleanup
            var narrative1 = narrativesList.GetNarrativeListItemFromText(description1);
            narrative1.Delete().Confirm();

            var narrative2 = narrativesList.GetNarrativeListItemFromText(description2);
            narrative2.Delete().Confirm();
        }

        [Test]
        [Category(SmokeTestCategory)]
        public void MatterEmailsWithBulkOperation()
        {
            var subject = _outlook.AddTestEmailsToFolder(1, FileSize.VerySmall, true, OfficeApp.Word).First().Key;
            _outlook.OpenTestEmailFolder();

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var emailsListPage = _outlook.Oc.EmailListPage;
            var emailsList = emailsListPage.ItemList;

            _outlook.SelectNthItem(0);

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetails.Tabs.Open("Emails");

            // quick file email
            matterDetails.QuickFile();

            var filedEmail = emailsList.GetEmailListItemFromText(subject);
            Assert.IsNotNull(filedEmail);

            filedEmail.Delete().Confirm();

            filedEmail = emailsList.GetEmailListItemFromText(subject, false);
            Assert.IsNull(filedEmail);

            var subject1 = _outlook.AddTestEmailsToFolder(1, FileSize.VerySmall, true).First().Key;

            // drag-n-drop email
            _outlook.DragAndDropEmailToOc(_outlook.GetNthItem(0), matterDetails.DropPoint.GetElement());
            filedEmail = emailsList.GetEmailListItemFromText(subject1);
            Assert.IsNotNull(filedEmail);

            // quick file additional email for bulk operation
            var subject2 = _outlook.AddTestEmailsToFolder(1).First().Key;
            _outlook.SelectNthItem(0);
            matterDetails.QuickFile();
            filedEmail = emailsList.GetEmailListItemFromText(subject2);
            Assert.IsNotNull(filedEmail);

            // bulk operation
            var firstEmail = emailsList.GetEmailListItemFromText(subject1);
            firstEmail.Select();

            var secondEmail = emailsList.GetEmailListItemFromText(subject2);
            secondEmail.Select();

            emailsListPage.DeleteEmails();

            firstEmail = emailsList.GetEmailListItemFromText(subject1, false);
            Assert.IsNull(firstEmail);

            secondEmail = emailsList.GetEmailListItemFromText(subject2, false);
            Assert.IsNull(secondEmail);
        }

        [Test]
        [Category(SmokeTestCategory)]
        public void NarrativesList()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var narrativesDetails = _outlook.Oc.NarrativesListPage;

            var description = GetRandomText(100);
            var narrative = GetRandomText(100);

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetails.Tabs.Open("Narratives");

            // Add narrative
            narrativesDetails.ItemList.OpenAddDialog();

            narrativesDetails.AddNarrativeDialog.Controls["Narrative Type"].Set("Note");
            narrativesDetails.AddNarrativeDialog.Controls["Description"].Set(description);
            narrativesDetails.AddNarrativeDialog.Controls["Narrative"].Set(narrative);

            narrativesDetails.AddNarrativeDialog.Save();

            var createdNote = narrativesDetails.ItemList.GetNarrativeListItemFromText(description);
            Assert.IsNotNull(createdNote);

            // Edit narrative
            description = GetRandomText(50);
            narrative = GetRandomText(50);

            createdNote.Edit();

            narrativesDetails.EditNarrativeDialog.Controls["Description"].Set(description);
            narrativesDetails.EditNarrativeDialog.Controls["Narrative"].Set(narrative);

            narrativesDetails.EditNarrativeDialog.Save();

            createdNote = narrativesDetails.ItemList.GetNarrativeListItemFromText(description);
            Assert.IsNotNull(createdNote);

            // Delete narrative
            createdNote.Delete().Confirm();

            createdNote = narrativesDetails.ItemList.GetNarrativeListItemFromText(description, false);
            Assert.IsNull(createdNote);
        }

        [Test]
        [Category(SmokeTestCategory)]
        public void DocumentsList()
        {
            var testEmail = _outlook.AddTestEmailsToFolder(1, FileSize.Medium, true).First();
            _outlook.OpenTestEmailFolder();
            _outlook.TurnOnReadingPane();

            var filename = new FileInfo(testEmail.Value).Name;

            var attachment = _outlook.GetAttachmentFromReadingPane(filename);

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetails.Tabs.Open("Documents");

            DragAndDrop.FromElementToElement(attachment, matterDetails.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();

            var uploadedDocument = documentsListPage.ItemList.GetMatterDocumentListItemFromText(filename);
            Assert.IsNotNull(uploadedDocument);

            uploadedDocument.Delete().Confirm();

            uploadedDocument = documentsListPage.ItemList.GetMatterDocumentListItemFromText(filename, false);
            Assert.IsNull(uploadedDocument);
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category("Documents")]
        public void DocumentCheckInCheckOut()
        {
            var checkedIn = CheckInStatus.CheckedIn.ToLower();
            var checkedOut = CheckInStatus.CheckedOut.ToLower();

            const string editedFileContent = "Content is edited by automated test.";
            const string dateTimeFormat = "M-dd-yyyy h-mm-ss tt";

            var testEmail = _outlook.AddTestEmailsToFolder(1, FileSize.VerySmall, true).First();
            _outlook.OpenTestEmailFolder();
            _outlook.TurnOnReadingPane();

            var filename = new FileInfo(testEmail.Value).Name;
            var attachment = _outlook.GetAttachmentFromReadingPane(filename);

            var newFolderName = DateTime.Now.ToString(dateTimeFormat);
            var renamedFolderName = DateTime.Now.AddYears(1).ToString(dateTimeFormat);

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var documentsList = documentsListPage.ItemList;
            var documentSummary = _outlook.Oc.DocumentSummaryPage;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetails.Tabs.Open("Documents");
            mattersListPage.ItemList.OpenAddFolderDialog();
            documentsListPage.AddFolderDialog.Controls["Name"].Set(newFolderName);
            documentsListPage.AddFolderDialog.Save();

            var testFolder = documentsList.GetMatterDocumentListItemFromText(newFolderName);
            Assert.IsNotNull(testFolder);

            testFolder.Open();

            var breadcrumbsPath = documentsListPage.BreadcrumbsControl.GetCurrentPath();
            Assert.IsTrue(breadcrumbsPath.EndsWith(newFolderName));

            documentsListPage.BreadcrumbsControl.NavigateToTheRoot();

            testFolder = documentsList.GetMatterDocumentListItemFromText(newFolderName);
            testFolder.Open();

            DragAndDrop.FromElementToElement(attachment, documentsListPage.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();

            var uploadedDocument = documentsList.GetMatterDocumentListItemFromText(filename);
            Assert.IsNotNull(uploadedDocument);

            var documentStatus = uploadedDocument.Status.ToLower();
            Assert.AreEqual(checkedIn, documentStatus);

            uploadedDocument.FileOptions.CheckOut();
            uploadedDocument = documentsList.GetMatterDocumentListItemFromText(filename);
            documentStatus = uploadedDocument.Status.ToLower();
            Assert.AreEqual(checkedOut, documentStatus);

            var notepad = new Notepad(filename);
            notepad.ReplaceTextWith(editedFileContent);
            notepad.Close();

            uploadedDocument.FileOptions.CheckIn();

            var checkInDocumentDialog = documentSummary.CheckInDocumentDialog;
            checkInDocumentDialog.Controls["Comments"].Set(AutomatedComment);
            checkInDocumentDialog.UploadDocument();

            uploadedDocument = documentsList.GetMatterDocumentListItemFromText(filename);
            documentStatus = uploadedDocument.Status.ToLower();
            Assert.AreEqual(checkedIn, documentStatus);

            uploadedDocument = documentsList.GetMatterDocumentListItemFromText(filename);
            var fileInfo = uploadedDocument.Download(filename);

            var fileContent = File.ReadAllText(fileInfo.FullName);
            Assert.AreEqual(editedFileContent, fileContent);

            uploadedDocument.Delete().Confirm();

            documentsListPage.BreadcrumbsControl.NavigateToTheRoot();

            testFolder = documentsList.GetMatterDocumentListItemFromText(newFolderName);
            testFolder.Rename();
            documentsListPage.AddFolderDialog.Controls["Name"].Set(renamedFolderName);
            documentsListPage.AddFolderDialog.Update();

            var renamedFolder = documentsList.GetMatterDocumentListItemFromText(renamedFolderName);
            Assert.IsNotNull(renamedFolder);

            renamedFolder.Delete().Confirm();

            var deletedFolder = documentsList.GetMatterDocumentListItemFromText(renamedFolderName, wait: false);
            Assert.IsNull(deletedFolder);
        }

        [Test]
        [Category(SmokeTestCategory)]
        public void DocumentSummary()
        {
            const string content = "Automated content";

            var testEmail = _outlook.AddTestEmailsToFolder(1, FileSize.Small, true).First();
            _outlook.OpenTestEmailFolder();
            _outlook.TurnOnReadingPane();

            var filename = new FileInfo(testEmail.Value).Name;

            var attachment = _outlook.GetAttachmentFromReadingPane(filename);

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var documentSummary = _outlook.Oc.DocumentSummaryPage;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetails.Tabs.Open("Documents");

            DragAndDrop.FromElementToElement(attachment, matterDetails.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();
            var uploadedDocument = documentsListPage.ItemList.GetMatterDocumentListItemFromText(filename);
            Assert.IsNotNull(uploadedDocument);

            //   documentSummary
            uploadedDocument.NavigateToSummary();

            var countEarlier = documentSummary.ItemList.GetCount();
            Assert.AreEqual(1, countEarlier);

            var summaryInfo = documentSummary.GetDocumentSummaryInfo();
            Assert.That(summaryInfo, Is.Not.Empty, "Document Summary fields are not retrieved or empty.");

            foreach (var webElement in summaryInfo)
            {
                Assert.IsNotEmpty(webElement.Text);
            }

            var documentNewVersion = CreateDocument(OfficeApp.Notepad, content);
            DragAndDrop.FromFileSystem(documentNewVersion, documentSummary.DropPoint.GetElement());

            var expectedDialogText = UploadDocumentMessage(documentNewVersion.Name, filename);
            var actualDialogText = documentSummary.AddDocumentDialog.Text;

            Assert.That(expectedDialogText, Is.EqualTo(actualDialogText), "Document Summary Upload message is not correct!");

            documentSummary.AddDocumentDialog.Cancel();

            DragAndDrop.FromFileSystem(documentNewVersion, documentSummary.DropPoint.GetElement());
            documentSummary.AddDocumentDialog.UploadDocument();

            var checkInDocumentDialog = documentSummary.CheckInDocumentDialog;

            checkInDocumentDialog.Controls["Comments"].Set(AutomatedComment);
            checkInDocumentDialog.UploadDocument();

            var versionsList = documentSummary.ItemList.GetAllVersionHistoryListItems().Select(x => x.Version).ToList();
            var countLater = versionsList.Count;
            var descendingVersionsList = versionsList.OrderByDescending(x => x).ToList();

            Assert.AreEqual(versionsList, descendingVersionsList);
            Assert.AreEqual(2, countLater);

            var newVersionDocument = documentsListPage.ItemList.GetVersionHistoryListItemByIndex(0);
            Assert.IsNotNull(newVersionDocument);

            Assert.IsNotNull(newVersionDocument.Version);
            Assert.IsNotEmpty(newVersionDocument.CreatedBy);
            Assert.IsNotEmpty(newVersionDocument.Comments);
            Assert.IsNotEmpty(newVersionDocument.Size);
            Assert.IsNotNull(newVersionDocument.UploadedAt);
            Assert.True(newVersionDocument.IsDownloadIconVisible());

            //download specific version
            var fileInfo = newVersionDocument.Download(filename);
            var fileContent = File.ReadAllText(fileInfo.FullName);
            Assert.AreEqual(fileContent, content);

            // Cleanup
            _outlook.Oc.Header.NavigateBack();
            uploadedDocument = documentsListPage.ItemList.GetMatterDocumentListItemFromText(filename);
            uploadedDocument.Delete().Confirm();
        }

        [Test]
        [Category(SmokeTestCategory)]
        public void RefreshAndBackButton()
        {
            const string dateTimeFormat = "M-dd-yyyy h-mm-ss tt";
            const string expectedTab = "documents";

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var documentsList = documentsListPage.ItemList;
            var ocHeader = _outlook.Oc.Header;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetails.Tabs.Open("Documents");

            var newFolderName = DateTime.Now.ToString(dateTimeFormat);
            mattersListPage.ItemList.OpenAddFolderDialog();
            documentsListPage.AddFolderDialog.Controls["Name"].Set(newFolderName);
            documentsListPage.AddFolderDialog.Save();

            var testFolder = documentsList.GetMatterDocumentListItemFromText(newFolderName);
            testFolder.Open();
            var breadcrumbsPath = documentsListPage.BreadcrumbsControl.GetCurrentPath();
            Assert.IsTrue(breadcrumbsPath.EndsWith(newFolderName));

            _outlook.Oc.ReloadOc();
            breadcrumbsPath = documentsListPage.BreadcrumbsControl.GetCurrentPath();
            Assert.IsTrue(breadcrumbsPath.EndsWith(newFolderName));

            matterDetails.Tabs.Open("Documents");
            testFolder = documentsList.GetMatterDocumentListItemFromText(newFolderName);
            testFolder.Open();
            breadcrumbsPath = documentsListPage.BreadcrumbsControl.GetCurrentPath();
            Assert.IsTrue(breadcrumbsPath.EndsWith(newFolderName));

            ocHeader.NavigateBack();
            var selectedTab = matterDetails.Tabs.GetActiveTab().ToLower();
            Assert.AreEqual(expectedTab, selectedTab);

            testFolder = documentsList.GetMatterDocumentListItemFromText(newFolderName);
            testFolder.Delete().Confirm();

            var deletedFolder = documentsList.GetMatterDocumentListItemFromText(newFolderName, false);
            Assert.IsNull(deletedFolder);
        }

        [Test]
        [Category(SmokeTestCategory)]
        public void UploadHistory()
        {
            const string expectedTab = "emails";

            var emails = _outlook.AddTestEmailsToFolder(3);
            _outlook.OpenTestEmailFolder();
            _outlook.SelectAllItems();

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var emailsListPage = _outlook.Oc.EmailListPage;
            var emailsList = emailsListPage.ItemList;
            var ocHeader = _outlook.Oc.Header;
            var ocUploadHistory = _outlook.Oc.UploadHistoryPage;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetails.Tabs.Open("Emails");

            matterDetails.QuickFile();

            ocHeader.OpenUploadQueue();
            ocHeader.OpenUploadHistory();

            var uploadHistoryItemsCount = ocUploadHistory.ItemList.GetCount();
            Assert.AreEqual(3, uploadHistoryItemsCount);

            ocUploadHistory.ClearUploadHistory();
            uploadHistoryItemsCount = ocUploadHistory.ItemList.GetCount();
            Assert.AreEqual(0, uploadHistoryItemsCount);

            ocUploadHistory.CloseUploadHistory();
            var selectedTab = matterDetails.Tabs.GetActiveTab().ToLower();
            Assert.AreEqual(expectedTab, selectedTab);

            // Cleanup
            foreach (var email in emails)
            {
                emailsList.GetEmailListItemFromText(email.Key).Select();
            }
            emailsListPage.DeleteEmails();
        }

        [Test]
        [Category(SmokeTestCategory)]
        public void WordDocumentCheckInCheckOut()
        {
            var checkedIn = CheckInStatus.CheckedIn.ToLower();
            var checkedOut = CheckInStatus.CheckedOut.ToLower();

            const string editedFileContent = "Content is edited by automated test.";

            var testEmail = _outlook.AddTestEmailsToFolder(1, FileSize.VerySmall, true, OfficeApp.Word).First();
            _outlook.OpenTestEmailFolder();
            _outlook.SelectNthItem(0);
            _outlook.TurnOnReadingPane();

            var filename = new FileInfo(testEmail.Value).Name;
            var attachment = _outlook.GetAttachmentFromReadingPane(filename);

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var documentsList = documentsListPage.ItemList;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetails.Tabs.Open("Documents");

            DragAndDrop.FromElementToElement(attachment, matterDetails.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();
            var uploadedDocument = documentsList.GetMatterDocumentListItemFromText(filename);
            Assert.IsNotNull(uploadedDocument);

            uploadedDocument.Open();
            _word = new Word(TestEnvironment);
            _word.Attach(filename);
            Assert.IsNotNull(_word.GetReadOnlyLabel());

            _word.CheckOut();
            _outlook.Oc.ReloadOc();
            var checkedOutDocument = documentsList.GetMatterDocumentListItemFromText(filename);
            var documentStatus = checkedOutDocument.Status.ToLower();
            Assert.AreEqual(checkedOut, documentStatus);

            _word.Attach(filename);
            _word.ReplaceTextWith(editedFileContent);
            _word.SaveDocument();
            _word.Close();

            checkedOutDocument.FileOptions.CheckIn();
            var checkInDocumentDialog = documentsListPage.CheckInDocumentDialog;
            checkInDocumentDialog.Controls["Comments"].Set(AutomatedComment);
            checkInDocumentDialog.UploadDocument();

            uploadedDocument = documentsList.GetMatterDocumentListItemFromText(filename);
            documentStatus = uploadedDocument.Status.ToLower();
            Assert.AreEqual(checkedIn, documentStatus);

            var fileInfo = uploadedDocument.Download(filename);
            Assert.IsTrue(fileInfo.Exists);

            _word = new Word(TestEnvironment);
            _word.OpenDocumentFromExplorer(fileInfo.FullName);
            Assert.IsNull(_word.GetReadOnlyLabel(false));
            _word.Close();

            var fileContent = _word.ReadWordContent(fileInfo.FullName);
            Assert.AreEqual(editedFileContent, fileContent);

            uploadedDocument.Delete().Confirm();
            uploadedDocument = documentsList.GetMatterDocumentListItemFromText(filename);
            Assert.IsNull(uploadedDocument);
        }

        [Test]
        [Category(SmokeTestCategory)]
        public void PeopleList()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var peopleListPage = _outlook.Oc.PeopleListPage;
            var peopleList = peopleListPage.ItemList;

            mattersListPage.ItemList.OpenFirst();
            matterDetails.Tabs.Open("People");
            peopleListPage.ItemList.OpenAddDialog();

            var addPersonDialog = _outlook.Oc.PeopleListPage.AddPersonDialog;

            addPersonDialog.Controls["Person Type"].SetByIndex(3);
            addPersonDialog.Controls["Comments"].Set("comments_test");
            addPersonDialog.Controls["Role/Involvement Type"].SetByIndex(2);
            var selectedPerson = addPersonDialog.Controls["Person"].SetByIndex(6);
            addPersonDialog.Controls["Start Date"].Set(FormatDate(DateTime.Now));
            addPersonDialog.Controls["End Date"].Set(FormatDate(DateTime.Now.AddDays(60)));
            addPersonDialog.Save();

            var createdPerson = peopleList.GetPeopleListItemFromText(selectedPerson);
            Assert.IsNotNull(createdPerson);

            createdPerson.Remove().Confirm();

            createdPerson = peopleList.GetPeopleListItemFromText(selectedPerson, wait: false);
            Assert.IsNull(createdPerson);
        }

        [Test]
        [Category(SmokeTestCategory)]
        public void EventsList()
        {
            var eventSubject = Guid.NewGuid().ToString();
            var eventDescription = $"Event generated at {FormatDateTime(DateTime.Now)}";

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var eventsListPage = _outlook.Oc.TasksEventsListPage;
            var eventsList = eventsListPage.ItemList;

            mattersListPage.Open();

            var matter = mattersListPage.ItemList.GetMatterListItemByIndex(0);
            var matterName = matter.Name;

            matter.Open();
            matterDetails.Tabs.Open("Tasks/Events");
            eventsList.OpenAddDialog();

            var addEventDialog = eventsListPage.AddEventDialog;
            addEventDialog.Controls["Type"].Set("Event");
            addEventDialog.Controls["Subject"].Set(eventSubject);
            addEventDialog.Controls["Start Date/Time"].Set(FormatDateTime(DateTime.Now));
            addEventDialog.Controls["End Date/Time"].Set(FormatDateTime(DateTime.Now));
            addEventDialog.Controls["Description"].Set(eventDescription);
            addEventDialog.Save();

            var addedEventItem = eventsList.GetTasksEventsListItemFromText(eventDescription);
            Assert.That(addedEventItem, Is.Not.Null, "Newly added event item is not listed after saving");

            _outlook.Calendars.RemovePassportCalendar(matterName);

            eventsListPage.ImportCalendar();

            var calendars = _outlook.Calendars.GetPassportCalendarsList().GetValue(0).ToString();
            Assert.IsTrue(calendars.Contains(matterName), "Matter calendar is not found in the list of passport calendars in Outlook");
            var deleteDialog = addedEventItem.Delete();
            deleteDialog.Confirm();

            var removedEventItem = eventsList.GetTasksEventsListItemFromText(eventDescription, false);
            Assert.That(removedEventItem, Is.Null, "Event item is listed after deleting it");

            // Cleanup
            _outlook.Calendars.RemovePassportCalendar(matterName);
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(DataDependentTestCategory)]
        public void TasksList()
        {
            const string dateTimeFormat = "M-dd-yyyy h-mm-ss tt";
            var taskName = Guid.NewGuid().ToString();
            var taskDescription = $"Task generated at {DateTime.Now.ToString(dateTimeFormat)}";

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var tasksListPage = _outlook.Oc.TasksEventsListPage;
            var tasksList = tasksListPage.ItemList;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();

            matterDetails.Tabs.Open("Tasks/Events");
            tasksListPage.ItemList.OpenAddDialog();

            var addTaskDialog = tasksListPage.AddTaskDialog;
            addTaskDialog.Controls["Type"].Set("Task");
            addTaskDialog.Controls["Name"].Set(taskName);
            addTaskDialog.Controls["Due Date"].Set(FormatDate(DateTime.Now));
            addTaskDialog.Controls["Description"].Set(taskDescription);
            addTaskDialog.Controls["Invitees/Assigned To"].Set("Sally Brown");
            addTaskDialog.Save();

            var addedTaskItem = tasksList.GetTasksEventsListItemFromText(taskName);
            Assert.That(addedTaskItem, Is.Not.Null, "Newly added task item is not listed after saving");

            var deleteDialog = addedTaskItem.Delete();
            deleteDialog.Confirm();

            var removedTaskItem = tasksList.GetTasksEventsListItemFromText(taskName, false);
            Assert.That(removedTaskItem, Is.Null, "Task item is listed after deleting it");
        }

        [Test]
        [Category(SmokeTestCategory)]
        public void DocumentsSort()
        {
            const int uniqueDocumentsAndFoldersCount = 2;
            var folderNames = new string[] { };
            IDictionary<string, string> emails = new Dictionary<string, string>();

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var documentsList = documentsListPage.ItemList;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetails.Tabs.Open("Documents");

            var docs = documentsList.GetAllMatterDocumentListItems();

            if (docs.Where(x => x.IsFolder).GroupBy(x => x.FolderName).Count() < uniqueDocumentsAndFoldersCount)
            {
                // Create folders
                folderNames = new[] { GetRandomText(5), GetRandomText(10) };

                foreach (var folderName in folderNames)
                {
                    mattersListPage.ItemList.OpenAddFolderDialog();
                    documentsListPage.AddFolderDialog.Controls["Name"].Set(folderName);
                    documentsListPage.AddFolderDialog.Save();
                }
            }

            if (docs.Where(x => !x.IsFolder).GroupBy(x => x.Name).Count() < uniqueDocumentsAndFoldersCount)
            {
                // Upload documents
                emails = _outlook.AddTestEmailsToFolder(uniqueDocumentsAndFoldersCount, FileSize.VerySmall, true);
                _outlook.OpenTestEmailFolder();
                _outlook.TurnOnReadingPane();

                for (var i = 0; i < emails.Count; i++)
                {
                    _outlook.SelectNthItem(i);
                    var filename = new FileInfo(emails.ElementAt(i).Value).Name;
                    var attachment = _outlook.GetAttachmentFromReadingPane(filename);
                    DragAndDrop.FromElementToElement(attachment, documentsListPage.DropPoint.GetElement());
                    documentsListPage.AddDocumentDialog.UploadDocument();

                    var uploadedDocument = documentsList.GetMatterDocumentListItemFromText(filename);
                    Assert.IsNotNull(uploadedDocument);
                }
            }

            // Verify scenario
            documentsListPage.DocumentSortDialog.Sort("Document File Name", SortOrder.Descending);

            docs = documentsList.GetAllMatterDocumentListItems();
            Assert.That(docs.Count(x => x.IsFolder), Is.GreaterThanOrEqualTo(uniqueDocumentsAndFoldersCount),
                "There are not enough folders to verify folders sorting. Need two folders at least.");
            Assert.That(docs.Where(x => !x.IsFolder).GroupBy(x => x.Name).Count(), Is.GreaterThanOrEqualTo(uniqueDocumentsAndFoldersCount),
                "There are no documents on the list. Need two documents to check sorting.");
            Assert.That(docs, Is.Ordered.Ascending.By(nameof(MatterDocumentListItem.IsFolder)).Using(new BooleanInverterComparer())
                .Then.Descending.By(nameof(MatterDocumentListItem.DocumentFileName)));

            documentsListPage.DocumentSortDialog.RestoreSortDefaults();

            docs = documentsList.GetAllMatterDocumentListItems();
            Assert.That(docs, Is.Ordered.Ascending.By(nameof(MatterDocumentListItem.IsFolder)).Using(new BooleanInverterComparer())
                .Then.Ascending.By(nameof(MatterDocumentListItem.Name)));
            Assert.That(docs.Where(x => x.IsFolder).ToList(), Is.Ordered.Ascending.By(nameof(MatterDocumentListItem.FolderName)));

            // Cleanup
            for (var i = 0; i < emails.Count; i++)
            {
                var filename = new FileInfo(emails.ElementAt(i).Value).Name;
                var document = documentsList.GetMatterDocumentListItemFromText(filename);
                document.Delete().Confirm();
            }

            foreach (var folderName in folderNames)
            {
                matterDetails.ItemList.GetEmailListItemFromText(folderName).Delete().Confirm();
            }
        }

        [Test]
        [Category(SmokeTestCategory)]
        public void Help()
        {
            var ocHeader = _outlook.Oc.Header;
            ocHeader.OpenHelp();

            var helpWindow = _outlook.Oc.HelpPage;
            var helpContents = helpWindow.GetAllLinksInFrame();

            Assert.NotZero(helpContents.Count);
        }

        [TearDown]
        public void TearDown()
        {
            SaveScreenShotsAndLogs(_outlook);
            SaveScreenShotsAndLogs(_word);
            _word?.Destroy();
            _outlook?.Destroy();
        }
    }
}
