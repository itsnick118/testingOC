using UITests.PageModel.Shared.InputControls;

namespace UITests.PageModel.Configurations.EY
{
    public class Dialogs
    {
        public static InputControlList AddNarrativeDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new Dropdown(app, "Narrative Type", "narrativeType"),
                new TextArea(app, "Narrative Description", "narrativeDescription"),
                new DateField(app, "Narrative Date", "narrativeDate"),
            };
        }

        public static InputControlList AddPersonDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new Dropdown(app, "Person Type", "matterPersonType"),
                new AutoComplete(app, "Person", "personInternalGco"),
                new Dropdown(app, "RoleInvolvement Type", "roleInvolvementInternalGco "),
                new TextArea(app, "Comments", "comments"),
                new DateField(app, "Start Date", "startDate"),
                new DateField(app, "End Date", "endDate"),
                new CheckBox(app, "Active", "active")
            };
        }

        public static InputControlList AddTaskDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new Dropdown(app, "Type", "eventType"),
                new InputField(app, "Name", "name"),
                new Dropdown(app, "Task Category", "taskCategory"),
                new Dropdown(app, "Task Sub-Category", "taskSubCategoryForAdministration"),
                new Dropdown(app, "Priority", "priority"),
                new DateField(app, "Due Date", "dueDate"),
                new DateField(app, "Completed Date", "completedDate"),
                new CheckBox(app, "Key Date", "keyDate"),
                new CheckBox(app, "Assignee's to Matter", "assigneesToMatter"),
                new TextArea(app, "Description", "description"),
                new AutoComplete(app, "Invitees/Assigned To", "people1", panelNumber: 1)
            };
        }

        public static InputControlList AddEventDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new Dropdown(app, "Type", "eventType"),
                new InputField(app, "Subject", "name1"),
                new Dropdown(app, "Event Category", "categoryType"),
                new Dropdown(app, "Event Sub-Category", "eventSubCategoryForAdministration"),
                new DateField(app, "Start Date/Time", "startDateTime"),
                new DateField(app, "End Date/Time", "endDateTime"),
                new CheckBox(app, "All Day Event", "allDayEvent"),
                new CheckBox(app, "Key Date", "keyDate"),
                new InputField(app, "Location", "location"),
                new CheckBox(app, "Assignee's to Matter", "assigneesToMatter"),
                new TextArea(app, "Description", "description"),
                new AutoComplete(app, "Invitees/Assigned To", "people", panelNumber: 1)
            };
        }

        public static InputControlList AddFolderDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new InputField(app, "Name", "name")
            };
        }

        public static InputControlList AddDocumentDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new InputField(app, "Document Name", "name"),
            };
        }

        public static InputControlList CheckInDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new Dropdown(app, "Select Check-in Version", "selectCheckinVersion"),
                new InputField(app, "Comments", "newCheckinComment")
            };
        }

        public static InputControlList SaveCurrentViewDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new InputField(app, "Create New", "SavedSearchName"),
                new Dropdown(app, "Update Existing", "SavedSearch")
            };
        }

        public static InputControlList EmailsListFilterDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new InputField(app, "Sender Name", "senderName"),
                new InputField(app, "Sender Email Address", "senderEmailAddress"),
                new InputField(app, "Subject", "subject"),
                new InputField(app, "Email Body", "mailBody"),
                new DateField(app, "Received Date", "receivedTime"),
                new Dropdown(app, "Has Attachment", "attachmentPresent")
            };
        }
    }
}
