using OpenQA.Selenium;
using UITests.PageModel.Selectors;
using UITests.PageModel.Shared;

namespace UITests.PageModel
{
    public class MatterDetailsPage
    {
        private readonly IAppInstance _app;
        public EntityTabs Tabs { get; }
        public SingleEntityDropPoint DropPoint { get; }
        public ItemList ItemList { get; }

        public string MatterName => new Group(_app).Title;
        public string MatterNumber => GetMatterProperty("matterNumber").Text;
        public string PrimaryInternalContact => GetMatterProperty("primaryInternalContact").Text;
        public string Status => GetMatterProperty("status").Text;
        public string MatterType => GetMatterProperty("matterType").Text;
        public string PracticeAreaBusinessUnit => GetMatterProperty("practiceAreaBusinessUnit").Text;
        public bool HasQuickFileIcon => _app.IsElementDisplayed(Oc.ItemDetailQuickFile);

        public MatterDetailsPage(IAppInstance app)
        {
            _app = app;
            Tabs = new EntityTabs(_app);
            DropPoint = new SingleEntityDropPoint(_app);
            ItemList = new ItemList(_app);
        }

        public void QuickFile(bool waitForQueueComplete = true)
        {
            _app.JustClick(Oc.ItemDetailQuickFile);

            if (waitForQueueComplete)
            {
                _app.WaitForQueueComplete();
            }
        }

        private IWebElement GetMatterProperty(string className) =>
            _app.Driver.FindElement(Oc.MatterPropertyByClass(className));

        public void AccessMatter()
        {
            _app.ClickAndWait(Oc.AccessButton);
        }
    }
}
