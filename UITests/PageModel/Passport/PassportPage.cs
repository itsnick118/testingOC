using OpenQA.Selenium;
using System;
using System.Threading;
using PassportSelector = UITests.PageModel.Selectors.Passport;

namespace UITests.PageModel.Passport
{
    public class PassportPage
    {
        protected readonly IAppInstance App;

        public PassportPage(IAppInstance app)
        {
            App = app;
        }

        public void SwitchToOc()
        {
            App.SwitchToFirstDriverWindow();
        }

        public void CloseWindowHandleSwitchToOc()
        {
            App.Driver.Close();
            SwitchToOc();
        }

        protected void SwitchToPassportFrame()
        {
            App.WaitUntilElementAppears(PassportSelector.PassportFrame);
            var iFrame = App.Driver.FindElement(PassportSelector.PassportFrame);
            App.WaitUntilElementDisappears(PassportSelector.PageLoading, iFrame);
            App.Driver.SwitchTo().Frame(iFrame);
        }

        protected void OpenMyPreferences()
        {
            App.SwitchToLastDriverWindow();
            App.WaitUntilElementAppears(PassportSelector.MySettingsIcon);
            App.JustClick(PassportSelector.MySettingsIcon);
            App.JustClick(PassportSelector.MyPreferencesButton);
        }

        protected void WaitAndClick(By selector)
        {
            WebDriverException ex = null;
            var cancellationToken = new CancellationTokenSource(TimeSpan.FromSeconds(Constants.NormalTimeoutSeconds)).Token;
            while (!cancellationToken.IsCancellationRequested)
            {
                try
                {
                    App.WaitUntilElementDisappears(PassportSelector.PageLoading);
                    App.JustClick(selector);
                    return;
                }
                catch (WebDriverException e)
                {
                    ex = e;
                }
            }

            throw new WebDriverException(ex?.Message, ex);
        }

    }
}
