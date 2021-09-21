using UITests.PageModel.Selectors;

namespace UITests.PageModel.Shared
{
    public class Group
    {
        private readonly IAppInstance _app;
        private readonly int _index;

        public Group(IAppInstance app, int groupIndex = 0)
        {
            _app = app;
            _index = groupIndex;
        }

        public string Title => _app.Driver.FindElement(Oc.Group).Text;
        public string HeaderValue => _app.Driver.FindElement(Oc.GroupHeaderValue).Text;

        public void Toggle() => _app.ClickAndWait(Oc.GroupByIndex(_index));
    }
}