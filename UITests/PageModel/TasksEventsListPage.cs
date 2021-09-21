using UITests.PageModel.Shared;

namespace UITests.PageModel
{
    public class TasksEventsListPage
    {
        private readonly IAppInstance _app;

        public IDialog AddEventDialog { get; }
        public IDialog AddTaskDialog { get; }
        public ISortDialog TasksEventsSortDialog { get; }
        public ItemList ItemList { get; }
        public QuickSearch QuickSearch { get; }

        public TasksEventsListPage(IAppInstance app)
        {
            _app = app;
            TasksEventsSortDialog = new SortDialog(_app);
            ItemList = new ItemList(_app);
            QuickSearch = new QuickSearch(_app);
            AddEventDialog = new Dialog(_app, null, Configurations.GA.Dialogs.AddEventDialogControls(_app));

            switch (_app.Environment.Configuration)
            {
                case EnvironmentConfiguration.GA:
                    AddEventDialog = new Dialog(_app, null, Configurations.GA.Dialogs.AddEventDialogControls(_app));
                    AddTaskDialog = new Dialog(_app, null, Configurations.GA.Dialogs.AddTaskDialogControls(_app));
                    break;

                case EnvironmentConfiguration.ICD:
                    AddEventDialog = new Dialog(_app, null, Configurations.ICD.Dialogs.AddEventDialogControls(_app));
                    AddTaskDialog = new Dialog(_app, null, Configurations.ICD.Dialogs.AddTaskDialogControls(_app));
                    break;

                case EnvironmentConfiguration.EY:
                    AddEventDialog = new Dialog(_app, null, Configurations.EY.Dialogs.AddEventDialogControls(_app));
                    AddTaskDialog = new Dialog(_app, null, Configurations.EY.Dialogs.AddTaskDialogControls(_app));
                    break;
            }
        }

        public void ImportCalendar() => _app.JustClick(Selectors.Oc.ImportMatterCalendarButton);
    }
}
