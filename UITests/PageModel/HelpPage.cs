using OpenQA.Selenium;
using System.Collections.Generic;
using PassportSelector = UITests.PageModel.Selectors.Passport;

namespace UITests.PageModel
{
    public class HelpPage
    {
        private readonly IAppInstance _app;

        public HelpPage(IAppInstance app)
        {
            _app = app;
        }

        public IReadOnlyCollection<IWebElement> GetAllLinksInFrame()
        {
            _app.WaitForLoadComplete();

            _app.SwitchToLastDriverWindow();

            SwitchToFrame(PassportSelector.MiniNavFrame);

            SwitchToFrame(PassportSelector.MiniBarFrame);

            SwitchToFrame(PassportSelector.NavPaneFrame);

            SwitchToFrame(PassportSelector.HelpItemsFrame);

            return _app.Driver.FindElements(PassportSelector.LinkItems);
        }

        private void SwitchToFrame(By frameSelector)
        {
            var frame = _app.Driver.FindElement(frameSelector);
            _app.Driver.SwitchTo().Frame(frame);
        }
    }
}