using UITests.PageModel.Shared;

namespace UITests.PageModel
{
    public class UploadHistoryPage
    {
        private readonly IAppInstance _app;

        public ItemList ItemList { get; }

        public UploadHistoryPage(IAppInstance app)
        {
            _app = app;
            ItemList = new ItemList(app);
        }

        public void ClearUploadHistory() => _app.ClickAndWait(Selectors.Oc.ClearUploadHistory);

        public void CloseUploadHistory() => _app.ClickAndWait(Selectors.Oc.CloseUploadHistory);
    }
}