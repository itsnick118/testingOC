using OpenQA.Selenium;

namespace UITests.PageModel
{
    public class BasicSettingsPage
    {
        private readonly IAppInstance _app;

        public BasicSettingsPage(IAppInstance app)
        {
            _app = app;
        }

        public bool CheckForLoading()
        {
            try
            {
                _app.WaitUntilElementAppears(Selectors.Oc.SpinnerLarge, timeoutSeconds: Constants.ShortTimeOutSeconds);
                return _app.IsElementDisplayed(Selectors.Oc.FormApplyButton);
            }
            catch (WebDriverTimeoutException)
            {
                return true;
            }
        }

        public void LogIn() => LogIn(_app.Environment.ElevatedUser, _app.Environment.ElevatedUserPassword);

        public void LogInAsAttorneyUser() => LogIn(_app.Environment.AttorneyUser, _app.Environment.AttorneyUserPassword);

        public void LogInAsStandardUser() => LogIn(_app.Environment.StandardUser, _app.Environment.StandardUserPassword);

        private void EnterBaseUrl() => EnterIn(Selectors.Oc.HostServerUrlInputBox, _app.Environment.BaseUrl);

        private void EnterPassword(string pwd) => EnterIn(Selectors.Oc.PasswordInputBox, pwd);

        private void EnterUserName(string user) => EnterIn(Selectors.Oc.UserNameInputBox, user);

        private void EnterIn(By inputBox, string value)
        {
            var userName = _app.Driver.FindElement(inputBox);
            userName.SendKeys(Keys.Control + "a");
            userName.Clear();
            userName.SendKeys(value);
        }

        private void LogIn(string userName, string password, int retryLimit = Constants.LoginRetryLimit)
        {
            EnterBaseUrl();
            EnterUserName(userName);
            EnterPassword(password);
            _app.JustClick(Selectors.Oc.FormApplyButton);

            var tries = 0;
            while (CheckForLoading() && tries < retryLimit)
            {
                EnterPassword(password);
                _app.JustClick(Selectors.Oc.FormApplyButton);
                tries++;

                if (tries >= retryLimit)
                {
                    throw new NoSuchWindowException("Retry limit for login exceeded.");
                }
            }
        }
    }
}
