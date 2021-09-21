using UITests.PageModel.Shared;

namespace UITests.PageModel
{
    public class NarrativesListPage
    {
        public IDialog AddNarrativeDialog { get; }
        public IDialog EditNarrativeDialog => AddNarrativeDialog;
        public ISortDialog NarrativeSortDialog { get; }
        public ItemList ItemList { get; }
        public QuickSearch QuickSearch { get; }

        public NarrativesListPage(IAppInstance app)
        {
            ItemList = new ItemList(app);
            NarrativeSortDialog = new SortDialog(app);
            QuickSearch = new QuickSearch(app);

            switch (app.Environment.Configuration)
            {
                default:
                    AddNarrativeDialog = new Dialog(app, null, Configurations.GA.Dialogs.AddNarrativeDialogControls(app));
                    break;
                case EnvironmentConfiguration.EY:
                    AddNarrativeDialog = new Dialog(app, null, Configurations.EY.Dialogs.AddNarrativeDialogControls(app));
                    break;
            }
        }
    }
}