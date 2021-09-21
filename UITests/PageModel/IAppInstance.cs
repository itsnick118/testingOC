using System;
using System.Drawing;
using OpenQA.Selenium;

namespace UITests.PageModel
{
    public interface IAppInstance
    {
        IWebDriver Driver { get; }
        TestEnvironment Environment { get; }

        void WaitForLoadComplete(int timeoutSeconds = Constants.NormalTimeoutSeconds);
        void WaitForListLoadComplete();
        void WaitForQueueComplete();
        void WaitFor(Func<IWebDriver, bool> condition, int timeoutSeconds = Constants.NormalTimeoutSeconds);
        void WaitUntilElementAppears(By elementSelector, int timeoutSeconds = Constants.NormalTimeoutSeconds);
        void WaitUntilElementDisappears(By elementSelector, IWebElement context = null, int timeoutSeconds = Constants.NormalTimeoutSeconds);
        void WaitForAnimatedTransitionComplete(IWebElement webElement);
        void SetShortImplicitWait();
        void SetLongImplicitWait();
        void ReloadOc();
        IWebElement ClickAndWait(By elementSelector, IWebElement element = null);
        IWebElement JustClick(By elementSelector, IWebElement element = null);
        IWebElement WaitAndClick(By elementSelector, IWebElement element = null);
        IWebElement WaitAndClickThenWait(By elementSelector, IWebElement element = null);
        void SwitchToLastDriverWindow(WindowHandles handles = WindowHandles.Multiple);
        bool IsElementDisplayed(By elementSelector, IWebElement context = null);
        bool IsElementSelected(By elementSelector);
        void SwitchToFirstDriverWindow();
        IWebElement GetToolTip(IWebElement element);
        void HoverMouseOverElement(By elementSelector);
        void HoverMouseOverElement(IWebElement webElement);
        Color GetColor(IWebElement webElement, string colorProperty = "color");
        Color GetCheckboxColor(IWebElement element);
        void OpenSettings();
    }
}
