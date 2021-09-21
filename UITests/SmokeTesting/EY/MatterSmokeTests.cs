using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using IntegratedDriver;
using NUnit.Framework;
using UITests.PageModel;
using UITests.PageModel.Shared;
using UITests.PageModel.Shared.Comparators;
using static UITests.Constants;
using static UITests.TestHelpers;

namespace UITests.SmokeTesting.EY
{
    [TestFixture]
    public class MatterSmokeTests : UITestBase
    {
        private Outlook _outlook;
        protected readonly IAppInstance App;

        public MatterSmokeTests()
        {
            Configuration = EnvironmentConfiguration.EY;
        }

        [SetUp]
        public void SetUp()
        {
            _outlook = new Outlook(TestEnvironment);
            _outlook.Launch();
            _outlook.Oc.BasicSettingsPage.LogIn();
        }

        [Test]
        [Category(SmokeTestCategory)]
        public void NarrativesList()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterList = mattersListPage.ItemList;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var narrativesListPage = _outlook.Oc.NarrativesListPage;
            var narrativeList = narrativesListPage.ItemList;

            var description = GetRandomText(100);

            mattersListPage.Open();
            matterList.OpenFirst();
            matterDetailsPage.Tabs.Open("Narratives");

            // Add narrative
            narrativeList.OpenAddDialog();

            narrativesListPage.AddNarrativeDialog.Controls["Narrative Type"].Set("Note");
            narrativesListPage.AddNarrativeDialog.Controls["Narrative Description"].Set(description);

            narrativesListPage.AddNarrativeDialog.Save();

            var createdNote = narrativeList.GetNarrativeListItemFromText(description);
            Assert.IsNotNull(createdNote);

            // Edit narrative
            description = GetRandomText(50);

            createdNote.Edit();

            narrativesListPage.EditNarrativeDialog.Controls["Narrative Description"].Set(description);

            narrativesListPage.EditNarrativeDialog.Save();

            createdNote = narrativeList.GetNarrativeListItemFromText(description);
            Assert.IsNotNull(createdNote);

            // Delete narrative
            createdNote.Delete().Confirm();

            createdNote = narrativeList.GetNarrativeListItemFromText(description, false);
            Assert.IsNull(createdNote);
        }

        [Test]
        [Category(SmokeTestCategory)]
        public void DocumentsList()
        {
            var testEmail = _outlook.AddTestEmailsToFolder(1, FileSize.Medium, true).First();
            _outlook.OpenTestEmailFolder();
            _outlook.TurnOnReadingPane();
            _outlook.SelectNthItem(0);

            var filename = new FileInfo(testEmail.Value).Name;

            var attachment = _outlook.GetAttachmentFromReadingPane(filename);

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var documentList = documentsListPage.ItemList;

            mattersListPage.Open();
            matterList.OpenFirst();
            matterDetails.Tabs.Open("Documents");

            DragAndDrop.FromElementToElement(attachment, matterDetails.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();

            var uploadedDocument = documentList.GetMatterDocumentListItemFromText(filename);
            Assert.IsNotNull(uploadedDocument);

            uploadedDocument.Delete().Confirm();

            uploadedDocument = documentList.GetMatterDocumentListItemFromText(filename, false);
            Assert.IsNull(uploadedDocument);
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(DataDependentTestCategory)]
        public void AllMattersSort()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterList = mattersListPage.ItemList;
            mattersListPage.Open();

            mattersListPage.MatterSortDialog.Sort("Matter Name", SortOrder.Descending);

            var matterItems = matterList.GetAllMatterListItems();
            Assert.That(matterItems.Count, Is.GreaterThanOrEqualTo(2), "No matters in the list to check sorting. Need two matters at least.");
            Assert.That(matterItems.Select(x => x.PrimaryInternalContact).All(x => !string.IsNullOrEmpty(x)));
            Assert.That(matterItems, Is.Ordered.Descending.By(nameof(MatterListItem.Name)));

            mattersListPage.MatterSortDialog.RestoreSortDefaults();

            matterItems = matterList.GetAllMatterListItems();
            Assert.That(matterItems.Select(x => x.Name).All(x => !string.IsNullOrEmpty(x)));
            Assert.That(matterItems, Is.Ordered.Ascending.By(nameof(MatterListItem.Name)));
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(DataDependentTestCategory)]
        public void FavoritesMattersList()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterList = mattersListPage.ItemList;

            mattersListPage.SetNthMatterAsFavorite(0);
            mattersListPage.SetNthMatterAsFavorite(1);

            mattersListPage.OpenFavoritesList();
            var favoritesMatterCount = matterList.GetCount();
            Assert.AreEqual(2, favoritesMatterCount);

            mattersListPage.ClearFavorites(2);
            Assert.That(matterList.GetCount(), Is.Zero, "Favorite list has matters after removing");
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(DataDependentTestCategory)]
        public void MyMattersSort()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterList = mattersListPage.ItemList;
            mattersListPage.OpenMyMattersList();

            Assert.Warn("Sort options not shown http://mingle/projects/growth/cards/20624");

            mattersListPage.MyMattersSortDialog.Sort("Matter Name", SortOrder.Descending);

            var matterItems = matterList.GetAllMatterListItems();
            Assert.That(matterItems.Count, Is.GreaterThanOrEqualTo(2), "No matters in the list to check sorting. Need two matters at least.");
            Assert.That(matterItems.Select(x => x.Name).All(x => !string.IsNullOrEmpty(x)));
            Assert.That(matterItems, Is.Ordered.Descending.By(nameof(MatterListItem.Name)));

            mattersListPage.MyMattersSortDialog.RestoreSortDefaults();

            matterItems = matterList.GetAllMatterListItems();
            Assert.That(matterItems.Select(x => x.StatusDate).All(x => x.HasValue));
            Assert.That(matterItems, Is.Ordered.Descending.By(nameof(MatterListItem.SpendToDate)));
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
            var matterList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var narrativesListPage = _outlook.Oc.NarrativesListPage;
            var narrativeList = narrativesListPage.ItemList;

            mattersListPage.Open();
            matterList.OpenFirst();
            matterDetails.Tabs.Open("Narratives");

            // Add narratives with narrative dates
            narrativeList.OpenAddDialog();
            narrativesListPage.AddNarrativeDialog.Controls["Narrative Type"].Set(type1);
            narrativesListPage.AddNarrativeDialog.Controls["Narrative Description"].Set(description1);
            narrativesListPage.AddNarrativeDialog.Controls["Narrative Date"].Set(date1);
            narrativesListPage.AddNarrativeDialog.Save();

            narrativeList.OpenAddDialog();
            narrativesListPage.AddNarrativeDialog.Controls["Narrative Type"].Set(type2);
            narrativesListPage.AddNarrativeDialog.Controls["Narrative Description"].Set(description2);
            narrativesListPage.AddNarrativeDialog.Controls["Narrative Date"].Set(date2);
            narrativesListPage.AddNarrativeDialog.Save();

            // Verify scenario
            narrativesListPage.NarrativeSortDialog.Sort("Narrative Type", SortOrder.Descending);

            var narrativeItems = narrativeList.GetAllNarrativeListItems();
            Assert.That(narrativeItems.Count, Is.GreaterThanOrEqualTo(2), "No narratives in the list to check sorting. Need two narratives at least.");
            Assert.That(narrativeItems.Select(x => x.Description).All(x => !string.IsNullOrEmpty(x)));
            Assert.That(narrativeItems, Is.Ordered.Descending.By(nameof(NarrativeListItem.Type)));

            narrativesListPage.NarrativeSortDialog.RestoreSortDefaults();

            narrativeItems = narrativeList.GetAllNarrativeListItems();
            Assert.That(narrativeItems.Select(x => x.NarrativeDate).Any(x => x.HasValue));
            Assert.That(narrativeItems, Is.Ordered.Descending.By(nameof(NarrativeListItem.NarrativeDate)));

            // Cleanup
            var narrative1 = narrativeList.GetNarrativeListItemFromText(description1);
            narrative1.Delete().Confirm();

            var narrative2 = narrativeList.GetNarrativeListItemFromText(description2);
            narrative2.Delete().Confirm();
        }

        [Test]
        [Category(SmokeTestCategory)]
        public void PeopleList()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var peopleListPage = _outlook.Oc.PeopleListPage;
            var peopleList = peopleListPage.ItemList;

            matterList.OpenFirst();
            matterDetails.Tabs.Open("People");
            peopleList.OpenAddDialog();

            var addPersonDialog = _outlook.Oc.PeopleListPage.AddPersonDialog;

            addPersonDialog.Controls["Person Type"].SetByIndex(3);
            addPersonDialog.Controls["RoleInvolvement Type"].SetByIndex(4);
            var selectedPerson = addPersonDialog.Controls["Person"].SetByIndex(3);
            addPersonDialog.Controls["Start Date"].Set(FormatDate(DateTime.Now));
            addPersonDialog.Controls["End Date"].Set(FormatDate(DateTime.Now.AddDays(60)));
            addPersonDialog.Controls["Comments"].Set("comments_test");
            addPersonDialog.Save();

            var selectedPersonName = string.Join(" ", selectedPerson.Split().Take(3));
            var createdPerson = peopleList.GetPeopleListItemFromText(selectedPersonName);
            Assert.IsNotNull(createdPerson);

            createdPerson.Remove().Confirm();

            createdPerson = peopleList.GetPeopleListItemFromText(selectedPerson, wait: false);
            Assert.IsNull(createdPerson);
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(DataDependentTestCategory)]
        public void PeopleSort()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var peopleListPage = _outlook.Oc.PeopleListPage;
            var peopleList = peopleListPage.ItemList;
            var selectedPersonList = new List<string>();

            matterList.OpenRandom();
            matterDetails.Tabs.Open("People");

            var peopleItems = peopleList.GetAllPeopleListItems();
            var uniquePeopleCount = peopleItems.GroupBy(x => x.PersonName).Count();
            while (uniquePeopleCount < 2)
            {
                var addPersonDialog = peopleListPage.AddPersonDialog;
                peopleList.OpenAddDialog();

                addPersonDialog.Controls["Person Type"].SetByIndex(3);
                addPersonDialog.Controls["Comments"].Set($"comments_test + {GetRandomNumber(3)}");
                addPersonDialog.Controls["RoleInvolvement Type"].SetByIndex(GetRandomNumber(3, 1));
                var selectedPerson = addPersonDialog.Controls["Person"].SetByIndex(GetRandomNumber(4, 1));
                addPersonDialog.Controls["Start Date"].Set(FormatDate(DateTime.Now));
                addPersonDialog.Controls["End Date"].Set(FormatDate(DateTime.Now.AddDays(GetRandomNumber(60, 1))));
                addPersonDialog.Save();

                peopleItems = peopleList.GetAllPeopleListItems();
                uniquePeopleCount = peopleItems.GroupBy(x => x.PersonName).Count();
                selectedPersonList.Add(selectedPerson);
            }

            peopleListPage.PeopleSortDialog.Sort("Person", SortOrder.Descending);

            peopleItems = peopleList.GetAllPeopleListItems();
            Assert.That(peopleItems, Is.Ordered.Descending.By(nameof(PeopleListItem.PersonName)));

            peopleListPage.PeopleSortDialog.RestoreSortDefaults();

            peopleItems = peopleList.GetAllPeopleListItems();

            //Cleanup
            foreach (var personList in selectedPersonList)
            {
                var selectedPersonName = string.Join(" ", personList.Split().Take(3));
                var createdPerson = peopleList.GetPeopleListItemFromText(selectedPersonName);
                createdPerson.Remove().Confirm();
            }
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
            var matterList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var tasksEventsListPage = _outlook.Oc.TasksEventsListPage;
            var tasksEventsList = tasksEventsListPage.ItemList;

            mattersListPage.Open();
            matterList.OpenFirst();
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
            tasksEventsListPage.AddTaskDialog.Controls["Invitees/Assigned To"].Set("sally");
            tasksEventsListPage.AddTaskDialog.Save();

            tasksEventsListPage.TasksEventsSortDialog.Sort("Name", SortOrder.Descending);

            var tasksEventsItems = tasksEventsList.GetAllTasksEventsListItems();
            Assert.That(tasksEventsItems.Count, Is.GreaterThanOrEqualTo(2), "No items added in the list to check sorting. Need two items at least.");
            Assert.That(tasksEventsItems, Is.Ordered.Descending.By(nameof(TasksEventsListItem.Name)));

            tasksEventsListPage.TasksEventsSortDialog.RestoreSortDefaults();
            tasksEventsItems = tasksEventsList.GetAllTasksEventsListItems();
            Assert.That(tasksEventsItems, Is.Ordered.Ascending.By(nameof(TasksEventsListItem.Type)));

            //Code Updated Start --Sumit
            tasksEventsListPage.TasksEventsSortDialog.Sort("End Date Time", SortOrder.Descending);
            //Code Updated End --Sumit

            // Cleanup
            matterDetails.ItemList.GetTasksEventsListItemFromText(eventSubjectText).Delete().Confirm();
            matterDetails.ItemList.GetTasksEventsListItemFromText(taskNameText).Delete().Confirm();
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
            var matterList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var tasksListPage = _outlook.Oc.TasksEventsListPage;
            var taskList = tasksListPage.ItemList;

            mattersListPage.Open();
            matterList.OpenFirst();

            matterDetails.Tabs.Open("Tasks/Events");
            taskList.OpenAddDialog();

            var addTaskDialog = tasksListPage.AddTaskDialog;
            addTaskDialog.Controls["Type"].Set("Task");
            addTaskDialog.Controls["Name"].Set(taskName);
            addTaskDialog.Controls["Due Date"].Set(FormatDate(DateTime.Now));
            addTaskDialog.Controls["Description"].Set(taskDescription);
            addTaskDialog.Controls["Invitees/Assigned To"].Set("sally");
            addTaskDialog.Save();

            var addedTaskItem = taskList.GetTasksEventsListItemFromText(taskName);
            Assert.That(addedTaskItem, Is.Not.Null, "Newly added task item is not listed after saving");

            var deleteDialog = addedTaskItem.Delete();
            deleteDialog.Confirm();

            var removedTaskItem = taskList.GetTasksEventsListItemFromText(taskName, false);
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
            var matterList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var documentsList = documentsListPage.ItemList;

            mattersListPage.Open();
            matterList.OpenFirst();
            matterDetails.Tabs.Open("Documents");

            var docs = documentsList.GetAllMatterDocumentListItems();

            if (docs.Where(x => x.IsFolder).GroupBy(x => x.FolderName).Count() < uniqueDocumentsAndFoldersCount)
            {
                // Create folders
                folderNames = new[] { GetRandomText(5), GetRandomText(10) };

                foreach (var folderName in folderNames)
                {
                    matterList.OpenAddFolderDialog();
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
            documentsListPage.DocumentSortDialog.Sort("Name", SortOrder.Descending);

            docs = documentsList.GetAllMatterDocumentListItems();
            Assert.That(docs.Count(x => x.IsFolder), Is.GreaterThanOrEqualTo(uniqueDocumentsAndFoldersCount),
                "There are not enough folders to verify folders sorting. Need two folders at least.");
            Assert.That(docs.Where(x => !x.IsFolder).GroupBy(x => x.Name).Count(), Is.GreaterThanOrEqualTo(uniqueDocumentsAndFoldersCount),
                "There are no documents on the list. Need two documents to check sorting.");
            Assert.That(docs, Is.Ordered.Ascending.By(nameof(MatterDocumentListItem.IsFolder)).Using(new BooleanInverterComparer())
                .Then.Descending.By(nameof(MatterDocumentListItem.Name)));

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
        public void EmailsSort()
        {
            var folderNames = new[] { GetRandomText(5), GetRandomText(10) };

            var emailsToUpload = _outlook.AddTestEmailsToFolder(2, useDifferentTemplates: true);
            _outlook.OpenTestEmailFolder();
            _outlook.SelectAllItems();
            _outlook.SelectNthItem(1);

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var emailsListPage = _outlook.Oc.EmailListPage;
            var emailList = emailsListPage.ItemList;
            var addFolderDialog = emailsListPage.AddFolderDialog;

            mattersListPage.Open();
            matterList.OpenFirst();
            matterDetails.Tabs.Open("Emails");
            _outlook.SelectAllItems();

            // Upload emails and add folders
            matterDetails.QuickFile();

            foreach (var folderName in folderNames)
            {
                emailList.OpenAddFolderDialog();
                addFolderDialog.Controls["Name"].Set(folderName);
                addFolderDialog.Save();
            }

            // Verify scenario
            emailsListPage.EmailsSortDialog.Sort("From", SortOrder.Descending);

            var emails = emailList.GetAllEmailListItems();
            Assert.That(emails.Where(x => !x.IsFolder).GroupBy(x => x.From).Count(), Is.GreaterThanOrEqualTo(2),
                "There are no emails or all emails from the same sender on the list. Need emails from different senders to check sorting.");
            Assert.That(emails.Count(x => x.IsFolder), Is.GreaterThanOrEqualTo(2),
                "There are not enough folders to verify folders sorting. Need two folders at least.");
            Assert.That(emails,
                Is.Ordered.Ascending.By(nameof(EmailListItem.IsFolder)).Using(new BooleanInverterComparer())
                    .Then.Descending.By(nameof(EmailListItem.From)));

            emailsListPage.EmailsSortDialog.RestoreSortDefaults();

            emails = emailList.GetAllEmailListItems();
            Assert.That(emails,
                Is.Ordered.Ascending.By(nameof(EmailListItem.IsFolder)).Using(new BooleanInverterComparer())
                    .Then.Descending.By(nameof(EmailListItem.ReceivedTime)));
            Assert.That(emails.Where(x => x.IsFolder).ToList(),
                Is.Ordered.Ascending.By(nameof(EmailListItem.ReceivedTime)));

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

        [TearDown]
        public void TearDown()
        {
            SaveScreenShotsAndLogs(_outlook);
            _outlook?.Destroy();
        }
    }
}
