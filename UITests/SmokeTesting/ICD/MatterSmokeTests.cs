using System;
using System.IO;
using System.Linq;
using IntegratedDriver;
using NUnit.Framework;
using UITests.PageModel;
using UITests.PageModel.Shared;
using static UITests.Constants;
using static UITests.TestHelpers;

namespace UITests.SmokeTesting.ICD
{
    public class MatterSmokeTests : UITestBase
    {
        private Outlook _outlook;

        public MatterSmokeTests()
        {
            Configuration = EnvironmentConfiguration.ICD;
        }

        [SetUp]
        public void SetUp()
        {
            _outlook = new Outlook(TestEnvironment);
            _outlook.Launch();
            _outlook.Oc.BasicSettingsPage.LogInAsStandardUser();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("Test case reference: 16705 CRUD Operation on Event")]
        public void EventsAddEditViewDelete()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var eventsListPage = _outlook.Oc.TasksEventsListPage;
            var eventsList = eventsListPage.ItemList;
            var eventDialog = eventsListPage.AddEventDialog;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetails.Tabs.Open("Events");

            // TODO: Fields triggering live update are excluded from the test because of dialog refresh

            // Verify to Add the new event
            eventsList.OpenAddDialog();
            var subject = eventDialog.Controls["Subject"].Set(Guid.NewGuid().ToString());
            var category = eventDialog.Controls["Event Category"].SetByIndex(1);
            var subCategory = eventDialog.Controls["Event Sub-Category"].SetByIndex(1);
            eventDialog.Controls["Start Date/Time"].Set(FormatDateTime(DateTime.Now));
            var startDateTime = eventDialog.Controls["Start Date/Time"].GetValue();
            eventDialog.Controls["End Date/Time"].Set(FormatDateTime(DateTime.Now.AddDays(1)));
            var endDateTime = eventDialog.Controls["End Date/Time"].GetValue();
            var keyDate = eventDialog.Controls["Key Date"].Set("Yes");
            var allDayEvent = "No";
            eventDialog.Controls["Location"].Set(Guid.NewGuid().ToString());
            var assigneesToMatter = "Yes";
            eventDialog.Controls["Description"].Set(Guid.NewGuid().ToString());
            eventDialog.Controls["Invitees/Assigned To"].SetByIndex(0);
            var invitees = eventDialog.Controls["Invitees/Assigned To"].GetValue();
            eventDialog.Save();

            // Verify to Edit an event with one or all fields
            var eventItem = eventsList.GetTasksEventsListItemFromText(subject);
            eventItem.Edit();
            subject = eventDialog.Controls["Subject"].Set(Guid.NewGuid().ToString());
            var location = eventDialog.Controls["Location"].Set(Guid.NewGuid().ToString());
            var description = eventDialog.Controls["Description"].Set(Guid.NewGuid().ToString());
            eventDialog.Save();

            // View Event
            eventItem = eventsList.GetTasksEventsListItemFromText(subject);
            eventItem.Open();
            Assert.That(eventDialog.Controls["Subject"].GetReadOnlyValue(), Is.EqualTo(subject));
            Assert.That(eventDialog.Controls["Event Category"].GetReadOnlyValue(), Is.EqualTo(category));
            Assert.That(eventDialog.Controls["Event Sub-Category"].GetReadOnlyValue(), Is.EqualTo(subCategory));
            Assert.That(eventDialog.Controls["Start Date/Time"].GetReadOnlyValue(), Is.EqualTo(startDateTime));
            Assert.That(eventDialog.Controls["End Date/Time"].GetReadOnlyValue(), Is.EqualTo(endDateTime));
            Assert.That(eventDialog.Controls["Key Date"].GetReadOnlyValue(), Is.EqualTo(keyDate));
            Assert.That(eventDialog.Controls["All Day Event"].GetReadOnlyValue(), Is.EqualTo(allDayEvent));
            Assert.That(eventDialog.Controls["Location"].GetReadOnlyValue(), Is.EqualTo(location));
            Assert.That(eventDialog.Controls["Assignee's to Matter"].GetReadOnlyValue(), Is.EqualTo(assigneesToMatter));
            Assert.That(eventDialog.Controls["Description"].GetReadOnlyValue(), Is.EqualTo(description));
            Assert.That(eventDialog.Controls["Invitees/Assigned To"].GetReadOnlyValue(), Is.EqualTo(invitees));
            eventDialog.Cancel();

            // Verify to Delete an event
            eventItem = eventsList.GetTasksEventsListItemFromText(subject);
            eventItem.Delete().Confirm();
            eventItem = eventsList.GetTasksEventsListItemFromText(subject, false);
            Assert.That(eventItem, Is.Null);
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("Test case reference: 16706 CRUD Operation on Task")]
        public void TasksAddViewEditSearchDelete()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var tasksListPage = _outlook.Oc.TasksEventsListPage;
            var tasksList = tasksListPage.ItemList;
            var taskDialog = tasksListPage.AddTaskDialog;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetails.Tabs.Open("Tasks");

            // Verify to Add new task
            tasksList.OpenAddDialog();
            Assert.That(taskDialog.HeaderText, Is.EqualTo("Add Task"));
            Assert.That(taskDialog.Controls["Name"].IsRequired, Is.True, "Name field is not marked as required.");
            Assert.That(taskDialog.Controls["Due Date"].IsRequired, Is.True, "Due Date field is not marked as required.");
            Assert.That(taskDialog.Controls["Invitees/Assigned To"].IsRequired, Is.True, "Invitees/Assigned To field is not marked as required.");
            Assert.That(taskDialog.Controls["Key Date"].GetValue(), Is.EqualTo("No"));
            Assert.That(taskDialog.Controls["Assignee's to Matter"].GetValue(), Is.EqualTo("Yes"));

            taskDialog.Save(false);
            Assert.That(taskDialog.Controls["Name"].GetRequiredWarning(), Is.EqualTo(FieldIsRequiredWarning));
            Assert.That(taskDialog.Controls["Due Date"].GetRequiredWarning(), Is.EqualTo(FieldIsRequiredWarning));
            Assert.That(taskDialog.Controls["Invitees/Assigned To"].GetRequiredWarning(), Is.EqualTo(FieldIsRequiredWarning));

            var taskName = taskDialog.Controls["Name"].Set(Guid.NewGuid().ToString());
            var dueDate = taskDialog.Controls["Due Date"].Set(FormatDate(DateTime.Now));
            taskDialog.Controls["Invitees/Assigned To"].SetByIndex(0);
            var invitees = taskDialog.Controls["Invitees/Assigned To"].GetValue();
            taskDialog.Save();

            // Verify to View the task with row click
            var task = tasksList.GetTasksEventsListItemFromText(taskName);
            task.Open();
            taskDialog.Edit();
            var taskCategory = taskDialog.Controls["Task Category"].SetByIndex(1);
            var taskSubCategory = taskDialog.Controls["Task Sub-Category"].SetByIndex(1);
            var priority = taskDialog.Controls["Priority"].SetByIndex(1);
            var completedDate = taskDialog.Controls["Completed Date"].Set(FormatDate(DateTime.Now));
            var description = taskDialog.Controls["Description"].Set(Guid.NewGuid().ToString());
            taskDialog.Cancel(false);
            Assert.That(taskDialog.HeaderText, Is.EqualTo(ConfirmationMessageHeader));
            Assert.That(taskDialog.Text, Is.EqualTo(CancelMessage));
            Assert.That(taskDialog.GetDialogButtons(), Is.EqualTo(new[] { "Discard Changes", "Don't Discard" }));

            taskDialog.DoNotDiscard();
            Assert.That(taskDialog.HeaderText, Is.EqualTo("Edit Task"));
            Assert.That(taskDialog.Controls["Task Category"].GetValue(), Is.EqualTo(taskCategory));
            Assert.That(taskDialog.Controls["Task Sub-Category"].GetValue(), Is.EqualTo(taskSubCategory));
            Assert.That(taskDialog.Controls["Priority"].GetValue(), Is.EqualTo(priority));
            Assert.That(taskDialog.Controls["Due Date"].GetValue(), Is.EqualTo(dueDate));
            Assert.That(taskDialog.Controls["Completed Date"].GetValue(), Is.EqualTo(completedDate));
            Assert.That(taskDialog.Controls["Key Date"].GetValue(), Is.EqualTo("No"));
            Assert.That(taskDialog.Controls["Assignee's to Matter"].GetValue(), Is.EqualTo("Yes"));
            Assert.That(taskDialog.Controls["Description"].GetValue(), Is.EqualTo(description));

            taskDialog.Cancel(false);
            taskDialog.DiscardChanges();

            // Verify to Edit from the created Task by pencil icon
            task = tasksList.GetTasksEventsListItemFromText(taskName);
            task.Edit();
            Assert.That(taskDialog.Controls["Name"].GetValue(), Is.EqualTo(taskName));
            Assert.That(taskDialog.Controls["Task Category"].GetValue(), Is.Empty);
            Assert.That(taskDialog.Controls["Priority"].GetValue(), Is.Empty);
            Assert.That(taskDialog.Controls["Due Date"].GetValue(), Is.EqualTo(dueDate));
            Assert.That(taskDialog.Controls["Completed Date"].GetValue(), Is.Empty);
            Assert.That(taskDialog.Controls["Key Date"].GetValue(), Is.EqualTo("No"));
            Assert.That(taskDialog.Controls["Assignee's to Matter"].GetValue(), Is.EqualTo("Yes"));
            Assert.That(taskDialog.Controls["Description"].GetValue(), Is.Empty);
            Assert.That(taskDialog.Controls["Invitees/Assigned To"].GetValue(), Is.EqualTo(invitees));

            taskDialog.Controls["Task Category"].SetByIndex(1);
            taskDialog.Controls["Task Sub-Category"].SetByIndex(1);
            taskDialog.Controls["Priority"].SetByIndex(1);
            taskDialog.Controls["Completed Date"].Set(FormatDate(DateTime.Now));
            taskDialog.Controls["Assignee's to Matter"].Set("No");
            taskDialog.Controls["Key Date"].Set("Yes");
            taskDialog.Controls["Description"].Set(Guid.NewGuid().ToString());

            taskDialog.Reset();
            Assert.That(taskDialog.Controls["Name"].GetValue(), Is.EqualTo(taskName));
            Assert.That(taskDialog.Controls["Task Category"].GetValue(), Is.Empty);
            Assert.That(taskDialog.Controls["Priority"].GetValue(), Is.Empty);
            Assert.That(taskDialog.Controls["Due Date"].GetValue(), Is.EqualTo(dueDate));
            Assert.That(taskDialog.Controls["Completed Date"].GetValue(), Is.Empty);
            Assert.That(taskDialog.Controls["Key Date"].GetValue(), Is.EqualTo("No"));
            Assert.That(taskDialog.Controls["Assignee's to Matter"].GetValue(), Is.EqualTo("Yes"));
            Assert.That(taskDialog.Controls["Description"].GetValue(), Is.Empty);
            Assert.That(taskDialog.Controls["Invitees/Assigned To"].GetValue(), Is.EqualTo(invitees));

            taskDialog.Controls["Task Category"].SetByIndex(1);
            taskDialog.Controls["Task Sub-Category"].SetByIndex(1);
            taskDialog.Controls["Priority"].SetByIndex(1);
            taskDialog.Controls["Completed Date"].Set(FormatDate(DateTime.Now));
            taskDialog.Controls["Description"].Set(Guid.NewGuid().ToString());
            taskDialog.Save();

            // Perform Search operations
            tasksListPage.QuickSearch.SearchBy(taskName);
            Assert.That(tasksList.GetCount(), Is.EqualTo(1));
            Assert.That(tasksList.GetTasksEventsListItemFromText(taskName), Is.Not.Null);

            // Perform Delete operation for newly created Task
            task = tasksList.GetTasksEventsListItemFromText(taskName);
            var deleteDialog = task.Delete();
            Assert.That(deleteDialog.HeaderText, Is.EqualTo("Delete Task"));
            Assert.That(deleteDialog.Text, Is.EqualTo($"Are you sure you want to delete '{taskName}' from the matter?{Environment.NewLine}This action cannot be undone."));
            Assert.That(deleteDialog.GetDialogButtons(), Is.EqualTo(new[] { "Ok", "Cancel" }));

            deleteDialog.Cancel();
            task = tasksList.GetTasksEventsListItemFromText(taskName);
            task.Delete().Confirm();
            task = tasksList.GetTasksEventsListItemFromText(taskName, false);
            Assert.That(task, Is.Null);
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("Test case reference: 16703 Verify to Create 'Pending Assignment' & 'Pending Acceptance' Matters from OC")]
        public void CreatePendingAssignmentMatterFromOc()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterPassportPage = _outlook.Oc.MatterPassportPage;

            // Verify to create "Pending Assignment" & "Pending Acceptance" Matters from OC
            mattersListPage.Open();
            mattersListPage.ItemList.AddMatter();

            var matterName = "z" + Guid.NewGuid();
            matterPassportPage.AddMatter(matterName, true);

            mattersListPage.QuickSearch.SearchBy(matterName);
            var matter = mattersListPage.ItemList.GetMatterListItemFromText(matterName);
            Assert.That(matter.Status, Is.EqualTo("Pending Assignment"));
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(DataDependentTestCategory)]
        [Description("Test case reference: 16704 Verify to navigate to Pending Assignment / Acceptance Matter from OC")]
        public void VerifyNavigateToPendingAssignmentAcceptanceMatterFromOc()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterPassportPage = _outlook.Oc.MatterPassportPage;
            var matterDetailPage = _outlook.Oc.MatterDetailsPage;
            var matterStatuses = new[] { "Pending Assignment", "Pending Acceptance", "Open", "Closed" };

            foreach (var matterStatus in matterStatuses)
            {
                mattersListPage.Open();

                //filter matter from list page
                mattersListPage.ItemList.OpenListOptionsMenu().OpenCreateListFilterDialog();
                mattersListPage.MatterListFilterDialog.Controls["Matter Status"].Set(matterStatus);
                mattersListPage.MatterListFilterDialog.Apply();

                var matter = mattersListPage.ItemList.GetMatterListItemByIndex(0);
                matter.AccessMatter();

                var matterStatusPassport = matterPassportPage.GetMatterStatus();
                Assert.AreEqual(matterStatusPassport, matterStatus);

                mattersListPage.ItemList.OpenFirst();
                matterDetailPage.AccessMatter();

                matterStatusPassport = matterPassportPage.GetMatterStatus();
                Assert.AreEqual(matterStatusPassport, matterStatus);
            }
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(DataDependentTestCategory)]
        [Description("TC_16707 Security : Verify the Matter status with 'Closed', 'Pending Assignment' & 'Pending Acceptance' are not editable from OC")]
        public void DisabledOperationsOnMatterListAndSummary()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var emailsListPage = _outlook.Oc.EmailListPage;

            var testEmail = _outlook.AddTestEmailsToFolder(1).First().Key;
            _outlook.OpenTestEmailFolder();

            // filter matter from list page
            var matterStatuses = new[] { "Pending Assignment", "Pending Acceptance", "Closed" };

            foreach (var matterStatus in matterStatuses)
            {
                mattersListPage.Open();
                mattersList.OpenListOptionsMenu().OpenCreateListFilterDialog();
                mattersListPage.MatterListFilterDialog.Controls["Matter Status"].Set(matterStatus);
                mattersListPage.MatterListFilterDialog.Apply();

                // verify disabled operations on matter list - quick file
                var filteredMatters = mattersList.GetAllMatterListItems();
                Assert.That(filteredMatters.Count, Is.GreaterThan(0), "Filtered list has no items");
                Assert.That(filteredMatters, Has.All.Property(nameof(MatterListItem.HasQuickFileIcon)).EqualTo(false));

                // drag-n-drop email
                var selectedMatter = mattersList.GetMatterListItemByIndex(GetRandomNumber(filteredMatters.Count - 1));
                _outlook.DragAndDropEmailToOc(_outlook.GetNthItem(0), selectedMatter.DropPoint, false);
                var messages = _outlook.Oc.GetAllToastMessages();
                Assert.AreEqual(1, messages.Length);
                var expected = $"Drag and drop not supported for `{matterStatus}` matter";
                StringAssert.Contains(expected, messages[0]);
                _outlook.Oc.CloseAllToastMessages();

                // verify disabled operations on matter summary - quick file
                selectedMatter.Open();
                Assert.IsFalse(matterDetailsPage.HasQuickFileIcon);

                // drag-n-drop email
                _outlook.DragAndDropEmailToOc(_outlook.GetNthItem(0), matterDetailsPage.DropPoint.GetElement());
                matterDetailsPage.Tabs.Open("Emails");
                var filedEmail = emailsListPage.ItemList.GetEmailListItemFromText(testEmail, false);
                Assert.IsNull(filedEmail);
            }
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(DataDependentTestCategory)]
        [TestCase("Pending Assignment")]
        [TestCase("Pending Acceptance")]
        [TestCase("Closed")]
        [Description("TC_16707 Security : Verify the Matter status with 'Closed', 'Pending Assignment' & 'Pending Acceptance' are not editable from OC")]
        public void DisabledOperationsOnMatterSummaryTabs(string matterStatus)
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var peopleListPage = _outlook.Oc.PeopleListPage;
            var emailsListPage = _outlook.Oc.EmailListPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var narrativesListPage = _outlook.Oc.NarrativesListPage;
            var tasksEventsListPage = _outlook.Oc.TasksEventsListPage;

            // filter matter from list page
            mattersListPage.Open();
            mattersList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            mattersListPage.MatterListFilterDialog.Controls["Matter Status"].Set(matterStatus);
            mattersListPage.MatterListFilterDialog.Apply();

            mattersList.OpenRandom();

            // verify disabled operations on People tab
            Assert.IsFalse(peopleListPage.ItemList.IsAddButtonVisible);
            var peopleItems = peopleListPage.ItemList.GetAllPeopleListItems();
            var selectedPerson = mattersList.GetPeopleListItemByIndex(GetRandomNumber(peopleItems.Count - 1));
            Assert.That(peopleItems, Has.All.Property(nameof(PeopleListItem.HasActionItems)).EqualTo(false));
            selectedPerson.Open();
            Assert.IsTrue(peopleListPage.ViewPersonDialog.IsDisplayed());
            peopleListPage.ViewPersonDialog.Cancel();

            // verify disabled operations on Emails tab
            matterDetailsPage.Tabs.Open("Emails");
            Assert.IsFalse(matterDetailsPage.HasQuickFileIcon);
            Assert.IsFalse(emailsListPage.ItemList.IsAddFolderButtonVisible);

            var testEmail = _outlook.AddTestEmailsToFolder(1, asAttachment: true).First();
            _outlook.OpenTestEmailFolder();
            _outlook.DragAndDropEmailToOc(_outlook.GetNthItem(0), matterDetailsPage.DropPoint.GetElement());
            var filedEmail = emailsListPage.ItemList.GetEmailListItemFromText(testEmail.Key, false);
            Assert.IsNull(filedEmail);

            // verify disabled operations on Documents tab
            matterDetailsPage.Tabs.Open("Documents");
            Assert.IsFalse(matterDetailsPage.HasQuickFileIcon);
            Assert.IsFalse(documentsListPage.ItemList.IsAddFolderButtonVisible);

            var fileName = new FileInfo(testEmail.Value).Name;
            var attachment = _outlook.GetAttachmentFromReadingPane(fileName);
            DragAndDrop.FromElementToElement(attachment, matterDetailsPage.DropPoint.GetElement());
            var uploadedDocument = documentsListPage.ItemList.GetMatterDocumentListItemFromText(fileName, false);
            Assert.IsNull(uploadedDocument);

            // verify disabled operations on Narratives tab
            matterDetailsPage.Tabs.Open("Narratives");
            Assert.IsFalse(narrativesListPage.ItemList.IsAddButtonVisible);
            _outlook.DragAndDropEmailToOc(_outlook.GetNthItem(0), matterDetailsPage.DropPoint.GetElement());
            var filedNarrative = narrativesListPage.ItemList.GetNarrativeListItemFromText(testEmail.Key, false);
            Assert.IsNull(filedNarrative);

            // verify disabled operations on Events tab
            matterDetailsPage.Tabs.Open("Events");
            Assert.IsFalse(tasksEventsListPage.ItemList.IsAddButtonVisible);

            // verify disabled operations on Tasks tab
            matterDetailsPage.Tabs.Open("Tasks");
            Assert.IsFalse(tasksEventsListPage.ItemList.IsAddButtonVisible);
        }

        [TearDown]
        public void TearDown()
        {
            SaveScreenShotsAndLogs(_outlook);
            _outlook?.Destroy();
        }
    }
}
