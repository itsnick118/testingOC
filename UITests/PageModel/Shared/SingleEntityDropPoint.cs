using OpenQA.Selenium;

namespace UITests.PageModel.Shared
{
    public class SingleEntityDropPoint
    {
        private readonly IAppInstance _app;

        public SingleEntityDropPoint(IAppInstance app)
        {
            _app = app;
        }

        public IWebElement GetElement()
        {
            var ppoDrop = _app.Driver.FindElements(Selectors.Oc.DropTarget);
            return ppoDrop[0];
        }
    }
}
