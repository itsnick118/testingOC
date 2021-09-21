using UITests.PageModel.Shared;

namespace UITests.PageModel
{
    public class GlobalDocumentsPage
    {
        private readonly IAppInstance _app;

        public IDialog SaveCurrentViewDialog { get; }
        public IDialog CheckInDocumentDialog { get; }
        public IDialog GlobalDocumentsListFilterDialog { get; }
        public IDialog CheckOutDocumentDialog { get; }
        public IDialog DiscardCheckOutDocumentDialog { get; }
        public ISortDialog RecentDocumentsSortDialog { get; }
        public ISortDialog CheckedOutDocumentsSortDialog { get; }
        public ISortDialog AllDocumentsSortDialog { get; }
        public ItemList ItemList { get; }
        public QuickSearch QuickSearch { get; }
        public EntityTabs Tabs { get; }
        public IDialog CheckInErrorDialog { get; }

        public GlobalDocumentsPage(IAppInstance app)
        {
            _app = app;
            SaveCurrentViewDialog = new Dialog(_app, null, Configurations.GA.Dialogs.SaveCurrentViewDialogControls(_app));
            CheckInDocumentDialog = new Dialog(_app, null, Configurations.GA.Dialogs.CheckInDialogControls(_app));
            GlobalDocumentsListFilterDialog = new Dialog(_app, null, Configurations.GA.Dialogs.GlobalDocumentsListFilterDialogControls(_app));
            CheckOutDocumentDialog = new Dialog(_app);
            CheckInErrorDialog = new Dialog(_app);
            DiscardCheckOutDocumentDialog = new Dialog(_app);
            RecentDocumentsSortDialog = new SortDialog(_app);
            CheckedOutDocumentsSortDialog = new SortDialog(_app);
            AllDocumentsSortDialog = new SortDialog(_app);
            ItemList = new ItemList(_app);
            QuickSearch = new QuickSearch(_app);
            Tabs = new EntityTabs(_app);
        }

        public void Open() => _app.WaitAndClickThenWait(Selectors.Oc.DocumentsTab);

        public void OpenAllDocumentsList()
        {
            _app.JustClick(Selectors.Oc.EntityAllDocumentsTab);
            _app.WaitForListLoadComplete();
        }

        public void ShowResultAllDocumentsList()
        {
            _app.JustClick(Selectors.Oc.AllDocumentShowResult);
            _app.WaitForListLoadComplete();
        }

        public void OpenCheckedOutDocumentsList()
        {
            _app.JustClick(Selectors.Oc.EntityCheckedOutTab);
            _app.WaitForListLoadComplete();
        }

        public void OpenRecentDocumentsList()
        {
            _app.JustClick(Selectors.Oc.EntityRecentDocumentsTab);
            _app.WaitForListLoadComplete();
        }
    }
}
