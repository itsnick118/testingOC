using UITests.PageModel.Selectors;
using UITests.PageModel.Shared;

namespace UITests.PageModel
{
    public class EmailsListPage
    {
        private readonly IAppInstance _app;

        public IDialog AddFolderDialog { get; }
        public IDialog SaveCurrentViewDialog { get; }
        public IDialog EmailsListFilterDialog { get; }
        public ISortDialog EmailsSortDialog { get; }
        public IDialog DeleteEmailDialog { get; }
        public ItemList ItemList { get; }
        public QuickSearch QuickSearch { get; }
        public BreadcrumbsControl BreadcrumbsControl { get; }
        public bool IsDeleteEmailsButtonDisplayed => _app.IsElementDisplayed(Oc.DeleteEmailsButton);

        public EmailsListPage(IAppInstance app)
        {
            _app = app;
            ItemList = new ItemList(app);
            switch (app.Environment.Configuration)
            {
                case EnvironmentConfiguration.GA:
                    AddFolderDialog = new Dialog(_app, null, Configurations.GA.Dialogs.AddFolderDialogControls(_app));
                    SaveCurrentViewDialog = new Dialog(_app, null, Configurations.GA.Dialogs.SaveCurrentViewDialogControls(_app));
                    EmailsListFilterDialog = new Dialog(_app, null, Configurations.GA.Dialogs.EmailsListFilterDialogControls(_app));
                    break;

                case EnvironmentConfiguration.EY:
                    AddFolderDialog = new Dialog(_app, null, Configurations.EY.Dialogs.AddFolderDialogControls(_app));
                    SaveCurrentViewDialog = new Dialog(_app, null, Configurations.EY.Dialogs.SaveCurrentViewDialogControls(_app));
                    EmailsListFilterDialog = new Dialog(_app, null, Configurations.EY.Dialogs.EmailsListFilterDialogControls(_app));
                    break;
            }
            EmailsSortDialog = new SortDialog(_app);
            DeleteEmailDialog = new Dialog(_app);
            QuickSearch = new QuickSearch(app);
            BreadcrumbsControl = new BreadcrumbsControl(app);
        }

        public void DeleteEmails()
        {
            _app.WaitUntilElementAppears(Oc.DeleteEmailsButton);
            _app.JustClick(Oc.DeleteEmailsButton);
            _app.JustClick(Oc.DialogOkButton);
            _app.WaitUntilElementDisappears(Oc.DeleteEmailsButton);
            _app.WaitForLoadComplete();
        }
    }
}
