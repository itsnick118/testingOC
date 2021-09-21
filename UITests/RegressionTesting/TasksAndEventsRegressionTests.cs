using System;
using System.Collections.Generic;
using NUnit.Framework;
using UITests.PageModel;
using static UITests.Constants;
using static UITests.TestHelpers;

namespace UITests.RegressionTesting
{
    public class TasksAndEventsRegressionTests : UITestBase
    {
        private Outlook _outlook;

        public TasksAndEventsRegressionTests()
        {
            Configuration = EnvironmentConfiguration.GA;
        }

        [SetUp]
        public void SetUp()
        {
            _outlook = new Outlook(TestEnvironment);
            _outlook.Launch();
            _outlook.Oc.BasicSettingsPage.LogInAsStandardUser();
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: 16498 Verify to get add events/tasks dialog by selecting type (Live update)")]
        public void TaskAndEventsLiveUpdate()
        {
            var defaultFieldsOfAddTaskAndEventDialog = new List<string> { "Type", "Description" };

            var addEventDialogFields = new List<string>{"Type", "Subject", "Category","Start Date/Time",
                                     "End Date/Time", "Location","Description", "Invitees/Assigned To" };

            var addTaskDialogFields = new List<string> { "Type", "Name", "Due Date","Completed Date",
                                                                   "Description", "Invitees/Assigned To"};
            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;

            var tasksEventsListPage = _outlook.Oc.TasksEventsListPage;
            var tasksEventsList = tasksEventsListPage.ItemList;
            var eventDialog = tasksEventsListPage.AddEventDialog;
            var taskDialog = tasksEventsListPage.AddTaskDialog;

            mattersListPage.Open();
            mattersList.OpenRandom();
            matterDetailsPage.Tabs.Open("Tasks/Events");

            tasksEventsList.OpenAddDialog();
            var labels = taskDialog.GetAllLabelTexts();
            Assert.AreEqual(labels, defaultFieldsOfAddTaskAndEventDialog);

            //Choose Type as Event and verify fields visibility after live update
            eventDialog.Controls["Type"].Set("Event");
            eventDialog.Controls["Invitees/Assigned To"].GetValue();
            labels = eventDialog.GetAllLabelTexts();
            eventDialog.Cancel(false);
            eventDialog.DiscardChanges();
            Assert.AreEqual(labels, addEventDialogFields);

            //Choose Type as Task and verify fields visibility after live update
            tasksEventsList.OpenAddDialog();
            taskDialog.Controls["Type"].Set("Task");
            taskDialog.Controls["Invitees/Assigned To"].GetValue();
            labels = taskDialog.GetAllLabelTexts();
            Assert.AreEqual(labels, addTaskDialogFields);
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: 16634 CRUD Verify Edit Tasks and Events")]
        public void EditTasksAndEvents()
        {
            var eventSubject = Guid.NewGuid().ToString();
            var eventDescription = $"Event generated at {FormatDateTime(DateTime.Now)}";
            var taskName = Guid.NewGuid().ToString();
            var taskDescription = $"Task generated at {FormatDateTime(DateTime.Now)}";

            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var tasksEventsListPage = _outlook.Oc.TasksEventsListPage;

            var tasksEventsList = tasksEventsListPage.ItemList;
            var eventDialog = tasksEventsListPage.AddEventDialog;
            var taskDialog = tasksEventsListPage.AddTaskDialog;

            mattersListPage.Open();
            mattersList.OpenRandom();
            matterDetailsPage.Tabs.Open("Tasks/Events");
            tasksEventsList.OpenAddDialog();

            // Create new Event
            var type = eventDialog.Controls["Type"].Set("Event");
            var subject = eventDialog.Controls["Subject"].Set(eventSubject);
            var startDate = eventDialog.Controls["Start Date/Time"].Set(FormatDateTime(DateTime.Now));
            var endDate = eventDialog.Controls["End Date/Time"].Set(FormatDateTime(DateTime.Now));
            var description = eventDialog.Controls["Description"].Set(eventDescription);
            eventDialog.Save();

            var addedEventItem = tasksEventsList.GetTasksEventsListItemFromText(subject);
            Assert.IsNotNull(addedEventItem, "Newly added event item is not listed after saving");

            // Get event in edit mode using pencil icon
            addedEventItem.Edit();

            // Verify buttons in edit dialog
            Assert.AreEqual(eventDialog.HeaderText, "Edit Event/Task");
            Assert.AreEqual(eventDialog.GetDialogButtons(), editPopupButtons);

            // Modify any of the fields
            var subjectModified = eventDialog.Controls["Subject"].Set("Modified Subject");
            Assert.AreEqual(eventDialog.Controls["Subject"].GetValue(), subjectModified);

            // Verify modified fields changes to original value on click of reset
            eventDialog.Reset();
            Assert.AreEqual(eventDialog.Controls["Subject"].GetValue(), subject);

            // Edit any of the fields and save
            var location = eventDialog.Controls["Location"].Set("New Location");
            eventDialog.Save();

            // Click on edit and modify any fields then click on close
            addedEventItem = tasksEventsList.GetTasksEventsListItemFromText(description);
            addedEventItem.Edit();
            Assert.AreEqual(eventDialog.Controls["Location"].GetValue(), location);

            eventDialog.Controls["Subject"].Set("Modified Subject");
            eventDialog.Cancel(false);

            Assert.AreEqual(eventDialog.HeaderText, ConfirmationMessageHeader);
            Assert.AreEqual(eventDialog.Text, CancelMessage);
            Assert.AreEqual(eventDialog.GetDialogButtons(), new[] { "Discard Changes", "Don't Discard" });
            eventDialog.DoNotDiscard();

            // Verify modified fields are same as edited after chosen do not discard
            Assert.AreEqual(eventDialog.Controls["Subject"].GetValue(), "Modified Subject");

            // click on close button and choose discard changes
            eventDialog.Cancel(false);
            eventDialog.DiscardChanges();

            // Perform row click on event and verify event fields are up to date as modified
            addedEventItem.Open();
            Assert.AreEqual(eventDialog.Controls["Type"].GetReadOnlyValue(), type);
            Assert.AreEqual(eventDialog.Controls["Subject"].GetReadOnlyValue(), subject);
            Assert.AreEqual(eventDialog.Controls["Start Date/Time"].GetReadOnlyValue(), startDate);
            Assert.AreEqual(eventDialog.Controls["End Date/Time"].GetReadOnlyValue(), endDate);
            Assert.AreEqual(eventDialog.Controls["Description"].GetReadOnlyValue(), description);
            Assert.AreEqual(eventDialog.Controls["Location"].GetReadOnlyValue(), location);

            // Bring event in edit mode by clicking on edit button from view mode
            eventDialog.Edit();

            // Verify buttons in edit dialog
            Assert.AreEqual(eventDialog.HeaderText, "Edit Event/Task");
            Assert.AreEqual(eventDialog.GetDialogButtons(), editPopupButtons);

            // Modify any of the fields
            var category = eventDialog.Controls["Category"].Set("Other");
            Assert.AreEqual(eventDialog.Controls["Category"].GetValue(), category);

            // Verify modified fields changes to original value on click of reset
            eventDialog.Reset();
            Assert.AreEqual(eventDialog.Controls["Category"].GetValue(), string.Empty);

            // Edit any of the fields and save
            category = eventDialog.Controls["Category"].Set(category);
            eventDialog.Save();

            // Click on edit from view mode and modify any fields then click on close
            addedEventItem = tasksEventsList.GetTasksEventsListItemFromText(description);
            addedEventItem.Open();
            eventDialog.Edit();
            Assert.AreEqual(eventDialog.Controls["Category"].GetValue(), category);

            eventDialog.Controls["Subject"].Set("Modified Subject");
            eventDialog.Cancel(false);
            Assert.AreEqual(eventDialog.HeaderText, ConfirmationMessageHeader);
            Assert.AreEqual(eventDialog.Text, CancelMessage);
            Assert.AreEqual(eventDialog.GetDialogButtons(), new[] { "Discard Changes", "Don't Discard" });

            eventDialog.DoNotDiscard();

            // Verify modified fields are same as edited after choosing do not discard
            Assert.AreEqual(eventDialog.Controls["Subject"].GetValue(), "Modified Subject");

            // click on close button and choose discard changes
            eventDialog.Cancel(false);
            eventDialog.DiscardChanges();

            // Perform row click on event and verify event fields are up to date as modified
            addedEventItem.Open();
            Assert.AreEqual(eventDialog.Controls["Type"].GetReadOnlyValue(), type);
            Assert.AreEqual(eventDialog.Controls["Subject"].GetReadOnlyValue(), subject);
            Assert.AreEqual(eventDialog.Controls["Category"].GetReadOnlyValue(), category);
            Assert.AreEqual(eventDialog.Controls["Start Date/Time"].GetReadOnlyValue(), startDate);
            Assert.AreEqual(eventDialog.Controls["End Date/Time"].GetReadOnlyValue(), endDate);
            Assert.AreEqual(eventDialog.Controls["Description"].GetReadOnlyValue(), description);
            Assert.AreEqual(eventDialog.Controls["Location"].GetReadOnlyValue(), location);

            // Delete an event
            eventDialog.Cancel();
            addedEventItem.Delete().Confirm();
            var removedEventItem = tasksEventsList.GetTasksEventsListItemFromText(subject, false);
            Assert.IsNull(removedEventItem, "Event item is listed after deleting it");

            // Create new task
            tasksEventsList.OpenAddDialog();
            var taskType = taskDialog.Controls["Type"].Set("Task");
            var name = taskDialog.Controls["Name"].Set(taskName);
            var dueDate = taskDialog.Controls["Due Date"].Set(FormatDate(DateTime.Now));
            var descriptionTask = taskDialog.Controls["Description"].Set(taskDescription);
            var assignedTo = taskDialog.Controls["Invitees/Assigned To"].Set("Sally Brown");
            taskDialog.Save();

            var addedTaskItem = tasksEventsList.GetTasksEventsListItemFromText(taskName);
            Assert.IsNotNull(addedTaskItem, "Newly added task item is not listed after saving");

            // Get task in edit mode using pencil icon
            addedTaskItem.Edit();

            // Verify buttons in edit dialog
            Assert.AreEqual(taskDialog.HeaderText, "Edit Event/Task");
            Assert.AreEqual(taskDialog.GetDialogButtons(), editPopupButtons);

            // Modify any of the fields
            var nameModified = taskDialog.Controls["Name"].Set("Modified Name");
            Assert.AreEqual(taskDialog.Controls["Name"].GetValue(), nameModified);

            // Verify modified fields changes to original value on click of reset
            taskDialog.Reset();
            Assert.AreEqual(taskDialog.Controls["Name"].GetValue(), name);

            // Edit any of the fields and save
            var completedDate = taskDialog.Controls["Completed Date"].Set(FormatDate(DateTime.Now));
            taskDialog.Save();

            // Click on edit and modify any fields then click on close
            addedTaskItem = tasksEventsList.GetTasksEventsListItemFromText(taskDescription);
            addedTaskItem.Edit();
            Assert.AreEqual(taskDialog.Controls["Completed Date"].GetValue(), completedDate);

            taskDialog.Controls["Name"].Set("Modified Name");
            taskDialog.Cancel(false);

            Assert.AreEqual(taskDialog.HeaderText, ConfirmationMessageHeader);
            Assert.AreEqual(taskDialog.Text, CancelMessage);
            Assert.AreEqual(taskDialog.GetDialogButtons(), new[] { "Discard Changes", "Don't Discard" });
            taskDialog.DoNotDiscard();

            // Verify modified fields are same as edited after choosing do not discard
            Assert.AreEqual(taskDialog.Controls["Name"].GetValue(), "Modified Name");

            // click on close button and choose discard changes
            taskDialog.Cancel(false);
            taskDialog.DiscardChanges();

            // Perform row click on task and verify event fields are up to date as modified
            addedTaskItem.Open();
            Assert.AreEqual(taskDialog.Controls["Type"].GetReadOnlyValue(), taskType);
            Assert.AreEqual(taskDialog.Controls["Name"].GetReadOnlyValue(), name);
            Assert.AreEqual(taskDialog.Controls["Due Date"].GetReadOnlyValue(), dueDate);
            Assert.AreEqual(taskDialog.Controls["Completed Date"].GetReadOnlyValue(), completedDate);
            Assert.AreEqual(taskDialog.Controls["Description"].GetReadOnlyValue(), descriptionTask);
            Assert.AreEqual(taskDialog.Controls["Invitees/Assigned To"].GetReadOnlyValue(), assignedTo);

            // Bring task in edit mode by clicking on edit button from view mode
            taskDialog.Edit();

            // Verify buttons in edit dialog
            Assert.AreEqual(taskDialog.HeaderText, "Edit Event/Task");
            Assert.AreEqual(taskDialog.GetDialogButtons(), editPopupButtons);

            // Modify any of the fields
            nameModified = taskDialog.Controls["Name"].Set("Modified Name");
            Assert.AreEqual(taskDialog.Controls["Name"].GetValue(), nameModified);

            // Verify modified fields changes to original value on click of reset
            taskDialog.Reset();
            Assert.AreEqual(taskDialog.Controls["Name"].GetValue(), name);

            // Edit any of the fields and save
            var updatedDescription = taskDialog.Controls["Description"].Set(taskDescription + " Updated");
            taskDialog.Save();

            // Click on edit from view mode and modify any fields then click on close
            addedTaskItem = tasksEventsList.GetTasksEventsListItemFromText(taskDescription);
            addedTaskItem.Open();
            taskDialog.Edit();
            Assert.AreEqual(taskDialog.Controls["Description"].GetValue(), updatedDescription);

            taskDialog.Controls["Name"].Set("Modified Name");
            taskDialog.Cancel(false);

            Assert.AreEqual(taskDialog.HeaderText, ConfirmationMessageHeader);
            Assert.AreEqual(taskDialog.Text, CancelMessage);
            Assert.AreEqual(taskDialog.GetDialogButtons(), new[] { "Discard Changes", "Don't Discard" });
            taskDialog.DoNotDiscard();

            // Verify modified fields are same as edited after chosen do not discard
            Assert.AreEqual(taskDialog.Controls["Name"].GetValue(), "Modified Name");

            // click on close button and choose discard changes
            taskDialog.Cancel(false);
            taskDialog.DiscardChanges();

            // Perform row click on task and verify event fields are up to date as modified
            addedTaskItem.Open();
            Assert.AreEqual(taskDialog.Controls["Type"].GetReadOnlyValue(), taskType);
            Assert.AreEqual(taskDialog.Controls["Name"].GetReadOnlyValue(), name);
            Assert.AreEqual(taskDialog.Controls["Due Date"].GetReadOnlyValue(), dueDate);
            Assert.AreEqual(taskDialog.Controls["Completed Date"].GetReadOnlyValue(), completedDate);
            Assert.AreEqual(taskDialog.Controls["Description"].GetReadOnlyValue(), updatedDescription);
            Assert.AreEqual(taskDialog.Controls["Invitees/Assigned To"].GetReadOnlyValue(), assignedTo);

            // clean up
            taskDialog.Cancel();
            addedTaskItem.Delete().Confirm();
            var removedTaskItem = tasksEventsList.GetTasksEventsListItemFromText(taskName, false);
            Assert.IsNull(removedTaskItem, "Task item is listed after deleting it");
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: 16497 Verify to view Tasks and Events")]
        public void ViewTasksAndEvents()
        {
            var eventSubject = Guid.NewGuid().ToString();
            var dateTime = FormatDateTime(DateTime.Now);
            var eventDescription = $"Event generated at {dateTime}";

            var taskName = Guid.NewGuid().ToString();
            var date = FormatDate(DateTime.Now);
            var taskDescription = $"Task generated at {dateTime}";

            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var tasksEventsDetailsPage = _outlook.Oc.TasksEventsListPage;

            var tasksEventsList = tasksEventsDetailsPage.ItemList;
            var eventDialog = tasksEventsDetailsPage.AddEventDialog;
            var taskDialog = tasksEventsDetailsPage.AddTaskDialog;

            mattersListPage.Open();
            mattersList.OpenRandom();
            matterDetailsPage.Tabs.Open("Tasks/Events");
            tasksEventsList.OpenAddDialog();

            // Create new Event
            var type = eventDialog.Controls["Type"].Set("Event");
            var subject = eventDialog.Controls["Subject"].Set(eventSubject);
            var startDate = eventDialog.Controls["Start Date/Time"].Set(dateTime);
            var endDate = eventDialog.Controls["End Date/Time"].Set(dateTime);
            var category = eventDialog.Controls["Category"].Set("Other");
            var location = eventDialog.Controls["Location"].Set("New Location");
            var description = eventDialog.Controls["Description"].Set(eventDescription);
            var assignedTo = eventDialog.Controls["Invitees/Assigned To"].Set("Sally Brown");

            eventDialog.Save();

            var addedEventItem = tasksEventsList.GetTasksEventsListItemFromText(subject);
            Assert.IsNotNull(addedEventItem, "Newly added event item is not listed after saving");

            // Perform row click on event and verify event fields are up to date as modified
            addedEventItem.Open();
            Assert.AreEqual(eventDialog.Controls["Type"].GetReadOnlyValue(), type);
            Assert.AreEqual(eventDialog.Controls["Subject"].GetReadOnlyValue(), subject);
            Assert.AreEqual(eventDialog.Controls["Category"].GetReadOnlyValue(), category);
            Assert.AreEqual(eventDialog.Controls["Start Date/Time"].GetReadOnlyValue(), startDate);
            Assert.AreEqual(eventDialog.Controls["End Date/Time"].GetReadOnlyValue(), endDate);
            Assert.AreEqual(eventDialog.Controls["Location"].GetReadOnlyValue(), location);
            Assert.AreEqual(eventDialog.Controls["Description"].GetReadOnlyValue(), description);
            Assert.AreEqual(eventDialog.Controls["Invitees/Assigned To"].GetReadOnlyValue(), assignedTo);

            // Delete an event
            eventDialog.Cancel();
            addedEventItem.Delete().Confirm();
            var removedEventItem = tasksEventsList.GetTasksEventsListItemFromText(subject, false);
            Assert.IsNull(removedEventItem, "Event item is listed after deleting it");

            // Create new task
            tasksEventsList.OpenAddDialog();
            var taskType = taskDialog.Controls["Type"].Set("Task");
            var name = taskDialog.Controls["Name"].Set(taskName);
            var dueDate = taskDialog.Controls["Due Date"].Set(date);
            var completedDate = taskDialog.Controls["Completed Date"].Set(date);
            description = taskDialog.Controls["Description"].Set(taskDescription);
            assignedTo = taskDialog.Controls["Invitees/Assigned To"].Set("Sally Brown");
            taskDialog.Save();

            var addedTaskItem = tasksEventsList.GetTasksEventsListItemFromText(taskName);
            Assert.IsNotNull(addedTaskItem, "Newly added task item is not listed after saving");

            // Perform row click on task and verify event fields are up to date as modified
            addedTaskItem.Open();
            Assert.AreEqual(taskDialog.Controls["Type"].GetReadOnlyValue(), taskType);
            Assert.AreEqual(taskDialog.Controls["Name"].GetReadOnlyValue(), name);
            Assert.AreEqual(taskDialog.Controls["Due Date"].GetReadOnlyValue(), dueDate);
            Assert.AreEqual(taskDialog.Controls["Completed Date"].GetReadOnlyValue(), completedDate);
            Assert.AreEqual(taskDialog.Controls["Description"].GetReadOnlyValue(), description);
            Assert.AreEqual(taskDialog.Controls["Invitees/Assigned To"].GetReadOnlyValue(), assignedTo);

            // clean up
            taskDialog.Cancel();
            addedTaskItem.Delete().Confirm();
            var removedTaskItem = tasksEventsList.GetTasksEventsListItemFromText(taskName, false);
            Assert.IsNull(removedTaskItem, "Task item is listed after deleting it");
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: 16541,16493 Verify to delete task and events from Tasks/Events list")]
        public void DeleteTasksAndEvents()
        {
            var eventSubject = Guid.NewGuid().ToString();
            var dateTime = FormatDateTime(DateTime.Now);
            var eventDescription = $"Event generated at {dateTime}";

            var taskName = Guid.NewGuid().ToString();
            var date = FormatDate(DateTime.Now);
            var taskDescription = $"Task generated at {dateTime}";

            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var tasksEventsListPage = _outlook.Oc.TasksEventsListPage;

            var tasksEventsList = tasksEventsListPage.ItemList;
            var eventDialog = tasksEventsListPage.AddEventDialog;
            var taskDialog = tasksEventsListPage.AddTaskDialog;

            mattersListPage.Open();
            mattersList.OpenRandom();
            matterDetailsPage.Tabs.Open("Tasks/Events");
            tasksEventsList.OpenAddDialog();

            // Create new Event
            eventDialog.Controls["Type"].Set("Event");
            eventDialog.Controls["Subject"].Set(eventSubject);
            eventDialog.Controls["Start Date/Time"].Set(FormatDateTime(DateTime.Now));
            eventDialog.Controls["End Date/Time"].Set(FormatDateTime(DateTime.Now));
            eventDialog.Controls["Description"].Set(eventDescription);
            eventDialog.Save();

            var addedEventItem = tasksEventsList.GetTasksEventsListItemFromText(eventSubject);
            Assert.IsNotNull(addedEventItem, "Newly added event item is not listed after saving");

            //Cancel Delete events
            addedEventItem.Delete().Cancel();
            var cancelEventItem = tasksEventsList.GetTasksEventsListItemFromText(eventSubject, false);
            Assert.IsNotNull(cancelEventItem, "Event item is listed before deleting it");

            // Delete an event
            addedEventItem.Delete().Confirm();
            var removedEventItem = tasksEventsList.GetTasksEventsListItemFromText(eventSubject, false);
            Assert.IsNull(removedEventItem, "Event item is listed after deleting it");

            // Create new task
            tasksEventsList.OpenAddDialog();
            taskDialog.Controls["Type"].Set("Task");
            taskDialog.Controls["Name"].Set(taskName);
            taskDialog.Controls["Due Date"].Set(date);
            taskDialog.Controls["Completed Date"].Set(date);
            taskDialog.Controls["Description"].Set(taskDescription);
            taskDialog.Controls["Invitees/Assigned To"].Set("Sally Brown");
            taskDialog.Save();

            var addedTaskItem = tasksEventsList.GetTasksEventsListItemFromText(taskName);
            Assert.IsNotNull(addedTaskItem, "Newly added task item is not listed after saving");

            //Cancel Delete task
            addedTaskItem.Delete().Cancel();
            var cancelTaskItem = tasksEventsList.GetTasksEventsListItemFromText(taskName, false);
            Assert.IsNotNull(cancelTaskItem, "Task list item is deleted after cancel delete operation");

            // Delete an task.
            addedTaskItem.Delete().Confirm();
            var removedTaskItem = tasksEventsList.GetTasksEventsListItemFromText(taskName, false);
            Assert.IsNull(removedTaskItem, "Task item is listed after deleting it");
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: 16495 Verify to add new Task with multiple person selection")]
        public void AddTasksWithMultiplePersons()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;

            var tasksEventsListPage = _outlook.Oc.TasksEventsListPage;
            var tasksEventsList = tasksEventsListPage.ItemList;
            var personMultiSelectList = _outlook.Oc.SelectPersonDialog;
            var personMultiselectItemList = personMultiSelectList.ItemList;

            var taskDialog = tasksEventsListPage.AddTaskDialog;
            var taskName = Guid.NewGuid().ToString();
            var taskDescription = $"Task generated at {FormatDateTime(DateTime.Now)}";

            var alice = "Alice Lee";
            var david = "David Maxwell";
            var eric = "Eric Stone";
            var multiplePersons = $"{alice},{eric},{david}";
            var aliceDavid = $"{alice},{david}";

            mattersListPage.Open();
            mattersList.OpenRandom();
            matterDetails.Tabs.Open("Tasks/Events");
            tasksEventsList.OpenAddDialog();

            taskDialog.Controls["Type"].Set("Task");
            taskDialog.Controls["Name"].Set(taskName);
            taskDialog.Controls["Due Date"].Set(FormatDate(DateTime.Now));
            taskDialog.Controls["Description"].Set(taskDescription);

            // Add Multiple persons to the task
            taskDialog.Controls["Invitees/Assigned To"].Set(alice);
            taskDialog.Controls["Invitees/Assigned To"].Set(eric);
            taskDialog.Controls["Invitees/Assigned To"].Set(david);

            var addedPerson = taskDialog.Controls["Invitees/Assigned To"].GetValue();
            Assert.AreEqual(addedPerson, multiplePersons);

            // Remove added persons from autocomplete then add new person and validate.
            taskDialog.Controls["Invitees/Assigned To"].Clear();
            taskDialog.Controls["Invitees/Assigned To"].Set(alice);
            addedPerson = taskDialog.Controls["Invitees/Assigned To"].GetValue();
            Assert.AreEqual(addedPerson, alice);

            // Navigate to multiselect window and verify added person displayed in multiselect window
            taskDialog.Controls["Invitees/Assigned To"].SelectPersonDialog();
            var addedPersonInMultiSelectDialog = personMultiSelectList.GetValue();
            Assert.AreEqual(addedPersonInMultiSelectDialog, alice);

            // Add a person from a list
            var assignedTo = personMultiselectItemList.GetMultiSelectPersonListItemFromText(david);
            assignedTo.Select();

            // Verify added person displayed in invitees list at multiselect window
            addedPersonInMultiSelectDialog = personMultiSelectList.GetValue();
            Assert.AreEqual(addedPersonInMultiSelectDialog, aliceDavid);

            // Verify removed person is not displayed in invitees list at multiselect window
            personMultiSelectList.Remove(david);
            addedPersonInMultiSelectDialog = personMultiSelectList.GetValue();
            Assert.AreEqual(addedPersonInMultiSelectDialog, alice);

            // Verify person added in multiselect dialog does not added in add task dialog
            personMultiSelectList.Close();
            addedPerson = taskDialog.Controls["Invitees/Assigned To"].GetValue();
            Assert.AreEqual(addedPerson, alice);

            // navigate to multiselect dialog and add person
            taskDialog.Controls["Invitees/Assigned To"].SelectPersonDialog();
            assignedTo = personMultiselectItemList.GetMultiSelectPersonListItemFromText(david);
            assignedTo.Select();

            addedPersonInMultiSelectDialog = personMultiSelectList.GetValue();
            Assert.AreEqual(addedPersonInMultiSelectDialog, aliceDavid);

            // Remove all persons from multiselect dialog
            personMultiSelectList.RemoveAll();
            addedPersonInMultiSelectDialog = personMultiSelectList.GetValue();
            Assert.IsEmpty(addedPersonInMultiSelectDialog);

            // Add a person from multi select dialog and click done
            assignedTo = personMultiselectItemList.GetMultiSelectPersonListItemFromText(alice);
            assignedTo.Select();
            personMultiSelectList.Done();

            // Use autocomplete to add another person
            taskDialog.Controls["Invitees/Assigned To"].Set(david);
            addedPerson = taskDialog.Controls["Invitees/Assigned To"].GetValue();
            Assert.AreEqual(addedPerson, aliceDavid);
            taskDialog.Save();

            // Verify task added to the list
            var addedTaskItem = tasksEventsList.GetTasksEventsListItemFromText(taskName);
            Assert.IsNotNull(addedTaskItem, "Newly added task item is not listed after saving");
            addedTaskItem.Delete().Confirm();

            // Clean up
            var removedTaskItem = tasksEventsList.GetTasksEventsListItemFromText(taskName, false);
            Assert.IsNull(removedTaskItem, "Task item is listed after deleting it");
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: 16495 Verify to add new event with multiple person selection")]
        public void AddEventsWithMultiplePersons()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;

            var tasksEventsListPage = _outlook.Oc.TasksEventsListPage;
            var tasksEventsList = tasksEventsListPage.ItemList;
            var personMultiSelectList = _outlook.Oc.SelectPersonDialog;
            var personMultiselectItemList = personMultiSelectList.ItemList;

            var eventDialog = tasksEventsListPage.AddEventDialog;
            var eventSubject = Guid.NewGuid().ToString();
            var eventDescription = $"Task generated at {FormatDateTime(DateTime.Now)}";

            var alice = "Alice Lee";
            var david = "David Maxwell";
            var eric = "Eric Stone";
            var multiplePersons = $"{alice},{eric},{david}";
            var aliceDavid = $"{alice},{david}";

            mattersListPage.Open();
            mattersList.OpenRandom();
            matterDetails.Tabs.Open("Tasks/Events");
            tasksEventsList.OpenAddDialog();

            eventDialog.Controls["Type"].Set("Event");
            eventDialog.Controls["Subject"].Set(eventSubject);
            eventDialog.Controls["Category"].Set("Other");
            eventDialog.Controls["Start Date/Time"].Set(FormatDateTime(DateTime.Now));
            eventDialog.Controls["End Date/Time"].Set(FormatDateTime(DateTime.Now));
            eventDialog.Controls["Description"].Set(eventDescription);

            // Add Multiple persons to the events
            eventDialog.Controls["Invitees/Assigned To"].Set(alice);
            eventDialog.Controls["Invitees/Assigned To"].Set(eric);
            eventDialog.Controls["Invitees/Assigned To"].Set(david);

            var addedPerson = eventDialog.Controls["Invitees/Assigned To"].GetValue();
            Assert.AreEqual(addedPerson, multiplePersons);

            // Remove added persons from autocomplete then add new person and validate.
            eventDialog.Controls["Invitees/Assigned To"].Clear();
            eventDialog.Controls["Invitees/Assigned To"].Set(alice);
            addedPerson = eventDialog.Controls["Invitees/Assigned To"].GetValue();
            Assert.AreEqual(addedPerson, alice);

            // Navigate to multiselect window and verify added person displayed in multiselect window
            eventDialog.Controls["Invitees/Assigned To"].SelectPersonDialog();
            var addedPersonInMultiSelectDialog = personMultiSelectList.GetValue();
            Assert.AreEqual(addedPersonInMultiSelectDialog, alice);

            // Add a person from a list
            var assignedTo = personMultiselectItemList.GetMultiSelectPersonListItemFromText(david);
            assignedTo.Select();

            // Verify added person displayed in invitees list at multiselect window
            addedPersonInMultiSelectDialog = personMultiSelectList.GetValue();
            Assert.AreEqual(addedPersonInMultiSelectDialog, aliceDavid);

            // Verify removed person is not displayed in invitees list at multiselect window
            personMultiSelectList.Remove(david);
            addedPersonInMultiSelectDialog = personMultiSelectList.GetValue();
            Assert.AreEqual(addedPersonInMultiSelectDialog, alice);

            // Verify person added in multiselect dialog does not added in add event dialog
            personMultiSelectList.Close();
            addedPerson = eventDialog.Controls["Invitees/Assigned To"].GetValue();
            Assert.AreEqual(addedPerson, alice);

            // navigate to multiselect dialog and add person
            eventDialog.Controls["Invitees/Assigned To"].SelectPersonDialog();
            assignedTo = personMultiselectItemList.GetMultiSelectPersonListItemFromText(david);
            assignedTo.Select();

            addedPersonInMultiSelectDialog = personMultiSelectList.GetValue();
            Assert.AreEqual(addedPersonInMultiSelectDialog, aliceDavid);

            // Remove all persons from multiselect dialog
            personMultiSelectList.RemoveAll();
            addedPersonInMultiSelectDialog = personMultiSelectList.GetValue();
            Assert.IsEmpty(addedPersonInMultiSelectDialog);

            // Add a person from multi select dialog and click done
            assignedTo = personMultiselectItemList.GetMultiSelectPersonListItemFromText(alice);
            assignedTo.Select();
            personMultiSelectList.Done();

            // Use autocomplete to add another person
            eventDialog.Controls["Invitees/Assigned To"].Set(david);
            addedPerson = eventDialog.Controls["Invitees/Assigned To"].GetValue();
            Assert.AreEqual(addedPerson, aliceDavid);
            eventDialog.Save();

            // Verify event added to the list
            var addedEventItem = tasksEventsList.GetTasksEventsListItemFromText(eventSubject);
            Assert.IsNotNull(addedEventItem, "Newly added event item is not listed after saving");
            addedEventItem.Delete().Confirm();

            // Clean up
            var removedEventItem = tasksEventsList.GetTasksEventsListItemFromText(eventSubject, false);
            Assert.IsNull(removedEventItem, "event item is listed after deleting it");
        }

        [TearDown]
        public void TearDown()
        {
            SaveScreenShotsAndLogs(_outlook);
            _outlook?.Destroy();
        }
    }
}
