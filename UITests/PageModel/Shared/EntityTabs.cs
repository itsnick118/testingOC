using UITests.PageModel.Selectors;

namespace UITests.PageModel.Shared
{
    public class EntityTabs
    {
        private readonly IAppInstance _app;

        public EntityTabs(IAppInstance app)
        {
            _app = app;
        }

        public void Open(string tabName) {
            _app.WaitAndClickThenWait(Oc.EntityTabByName(tabName));
            _app.WaitForListLoadComplete();
        }

        public string GetActiveTab()
        {
            _app.WaitForLoadComplete();
            var activeTab = _app.Driver.FindElement(Oc.EntityActiveTab);
            return activeTab.Text;
        }
    }
}