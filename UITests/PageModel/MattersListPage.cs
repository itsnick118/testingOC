using UITests.PageModel.Shared;

namespace UITests.PageModel
{
    public class MattersListPage
    {
        private readonly IAppInstance _app;

        public IDialog SaveCurrentViewDialog { get; }
        public IDialog MatterListFilterDialog { get; }
        public ISortDialog MatterSortDialog { get; }
        public ISortDialog MyMattersSortDialog { get; }
        public ItemList ItemList { get; }
        public QuickSearch QuickSearch { get; }

        public MattersListPage(IAppInstance app)
        {
            _app = app;
            SaveCurrentViewDialog = new Dialog(app, null, Configurations.GA.Dialogs.SaveCurrentViewDialogControls(app));
            MatterListFilterDialog = new Dialog(_app, null, Configurations.GA.Dialogs.MatterListFilterDialogControls(_app));
            MatterSortDialog = new SortDialog(_app);
            MyMattersSortDialog = new SortDialog(_app);
            ItemList = new ItemList(_app);
            QuickSearch = new QuickSearch(_app);
        }

        public void Open() => _app.WaitAndClickThenWait(Selectors.Oc.MattersTab);

        public void SetNthMatterAsFavorite(int n)
        {
            _app.WaitForLoadComplete();

            var nthMatter = ItemList.GetMatterListItemByIndex(n);
            nthMatter.SetAsFavorite();
        }

        public void ClearFavorites(int numFavorites)
        {
            for (var i = numFavorites - 1; i >= 0; i--)
            {
                var ithMatter = ItemList.GetMatterListItemByIndex(i);
                ithMatter.ClearFavorite();
            }
        }

        public void ClearAllFavorites() => ClearFavorites(ItemList.GetCount());

        public void OpenFavoritesList() => _app.WaitAndClickThenWait(Selectors.Oc.FavoritesTab);

        public void OpenAllMattersList() => _app.WaitAndClickThenWait(Selectors.Oc.AllMattersTab);

        public void OpenMyMattersList() => _app.WaitAndClickThenWait(Selectors.Oc.MyMattersTab);
    }
}