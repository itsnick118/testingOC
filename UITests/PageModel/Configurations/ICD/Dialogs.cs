using UITests.PageModel.Shared.InputControls;

namespace UITests.PageModel.Configurations.ICD
{
    public class Dialogs
    {
        public static InputControlList AddTaskDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new InputField(app, "Name", "name"),
                new Dropdown(app, "Task Category", "taskCategory"),
                new Dropdown(app, "Task Sub-Category", "taskSubCategoryForAdministration"),
                new Dropdown(app, "Priority", "priority"),
                new DateField(app, "Due Date", "dueDate"),
                new DateField(app, "Completed Date", "completedDate"),
                new CheckBox(app, "Key Date", "keyDate"),
                new CheckBox(app, "Assignee's to Matter", "assigneesToMatter"),
                new TextArea(app, "Description", "description"),
                new AutoComplete(app, "Invitees/Assigned To", "matterPerson", panelNumber: 1)
            };
        }

        public static InputControlList AddEventDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new InputField(app, "Subject", "subject"),
                new Dropdown(app, "Event Category", "categoryType"),
                new Dropdown(app, "Event Sub-Category", "eventSubCategoryForAdministration"),
                new DateField(app, "Start Date/Time", "startDateTime"),
                new DateField(app, "End Date/Time", "endDateTime"),
                new CheckBox(app, "All Day Event", "allDayEvent"),
                new CheckBox(app, "Key Date", "keyDate"),
                new InputField(app, "Location", "location"),
                new CheckBox(app, "Assignee's to Matter", "assigneesToMatter"),
                new TextArea(app, "Description", "description"),
                new AutoComplete(app, "Invitees/Assigned To", "matterPerson", panelNumber: 1)
            };
        }
    }
}
