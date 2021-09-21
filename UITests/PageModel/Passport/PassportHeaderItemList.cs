using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using PassportSelector = UITests.PageModel.Selectors.Passport;

namespace UITests.PageModel.Passport
{
    public class PassportHeaderItemList
    {
        private readonly IAppInstance _app;

        public PassportHeaderItemList(IAppInstance app)
        {
            _app = app;
        }

        public PassportHeaderAdjustmentListItem GetPassportListItemFromText(string content, bool wait = true) => GetPassportListItemFromText<PassportHeaderAdjustmentListItem>(content, wait);

        public IReadOnlyCollection<IWebElement> GetAllWebElementListItemsPassport()
        {
            _app.WaitUntilElementAppears(PassportSelector.ListItems);

            return _app.Driver.FindElements(PassportSelector.ListItems);
        }

        public int GetCountPassport() => GetAllWebElementListItemsPassport().Count;

        private T GetPassportListItemFromText<T>(string content, bool wait = true) where T : PassportHeaderAdjustmentListItem
        {
            var element = GetPassportListItem(content, wait);
            return element == null ? null : (T)Activator.CreateInstance(typeof(T), _app);
        }

        private IWebElement GetPassportListItem(string content, bool wait = true)
        {
            IWebElement result = null;

            if (!wait) _app.SetShortImplicitWait();
            _app.WaitUntilElementAppears(PassportSelector.ListItems);

            do
            {
                try
                {
                    result = _app.Driver.FindElement(PassportSelector.ListItemByContent(content));
                }
                catch
                {
                    // swallowing
                }
            } while (result == null);
            return result;
        }
    }
}