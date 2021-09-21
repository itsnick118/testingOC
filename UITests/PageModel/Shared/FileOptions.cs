using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using OpenQA.Selenium;
using UITests.PageModel.Selectors;

namespace UITests.PageModel.Shared
{
    public class FileOptions
    {
        private readonly IAppInstance _app;
        private readonly IWebElement _element;
        private readonly IDialog _discardCheckOutDocumentDialog;

        public FileOptions(IAppInstance app, IWebElement element)
        {
            _app = app;
            _element = element;
            _discardCheckOutDocumentDialog = new Dialog(app);
        }

        public void CheckOut()
        {
            OpenFileOptionsAndClickOnButton(Oc.ItemOptionsCheckOutButton);
            WaitForDialogOrStatusChange();
        }

        public void CheckIn()
        {
            OpenFileOptionsAndClickOnButton(Oc.ItemOptionsCheckInButton);
            _app.WaitUntilElementAppears(Oc.DialogHeader);
        }

        public void DiscardCheckOut()
        {
            OpenFileOptionsAndClickOnButton(Oc.ItemOptionsDiscardCheckOutButton);
            _app.WaitForLoadComplete();
        }

        public void DiscardCheckOutAndRemoveLocalCopy()
        {
            OpenFileOptionsAndClickOnButton(Oc.ItemOptionsDiscardCheckOutButton);
            _discardCheckOutDocumentDialog.Remove();
            _app.WaitUntilElementAppears(Oc.MatProgressBar);
            _app.WaitForLoadComplete();
        }

        public List<IWebElement> OpenFileOptionsAndGetOptions()
        {
            _app.JustClick(Oc.ItemOptions, _element);
            List<IWebElement> status = _element.FindElements(Oc.AllOptions).ToList();
            _app.JustClick(Oc.AllOptionsOverlay, _element);
            return status;
        }

        private void OpenFileOptionsAndClickOnButton(By selector)
        {
            _app.JustClick(Oc.ItemOptions, _element);
            _app.WaitAndClick(selector);
        }

        private void WaitForDialogOrStatusChange()
        {
            _app.WaitFor(condition =>
            {
                if (_app.IsElementDisplayed(Oc.DialogHeader) || IsStaleElement())
                {
                    return true;
                }

                try
                {
                    var options = _element.FindElement(Oc.ItemOptions).Text.ToLower();
                    return options.Contains("out");
                }
                catch (NoSuchElementException)
                {
                    return false;
                }
            });
        }

        [SuppressMessage("ReSharper", "UnusedVariable")]
        private bool IsStaleElement()
        {
            try
            {
                var staleElement = _element.Displayed;
                return false;
            }
            catch (StaleElementReferenceException)
            {
                return true;
            }
        }
    }
}
