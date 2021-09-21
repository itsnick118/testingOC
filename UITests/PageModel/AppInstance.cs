using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading;
using UITests.PageModel.Selectors;
using static IntegratedDriver.Constants;

namespace UITests.PageModel
{
    public class AppInstance : IAppInstance
    {
        public TestEnvironment Environment { get; }
        public Process OfficeAppProcess { get; }
        public IWebDriver Driver { get; }

        private const string GetHeightScript = @"return window.innerHeight";

        private const string DocumentReadyStateScript = @"var interval = setInterval(function() { 
            if(document.readyState === 'complete') { 
                clearInterval(interval);
            }}, 100);";

        public AppInstance(TestEnvironment environment, Process officeAppProcess, IWebDriver driver)
        {
            Environment = environment;
            OfficeAppProcess = officeAppProcess;
            Driver = driver;
        }

        public void WaitForLoadComplete(int timeoutSeconds = Constants.NormalTimeoutSeconds)
        {
            timeoutSeconds = timeoutSeconds > 0 ? timeoutSeconds : Constants.LongTimeoutSeconds;
            var initialState = Driver.PageSource.GetHashCode();

            var officeApp = OfficeAppProcess;
            officeApp.WaitForInputIdle();

            // is the DOM ready?
            ((IJavaScriptExecutor)Driver).ExecuteScript(DocumentReadyStateScript);

            // are there any visible spinners?
            SetShortImplicitWait();

            var viewportHeight = Convert.ToInt32(((IJavaScriptExecutor)Driver).ExecuteScript(GetHeightScript));

            new WebDriverWait(Driver, TimeSpan.FromSeconds(timeoutSeconds))
                .Until(condition =>
                {
                    var actionSpinners = Driver.FindElements(Oc.SpinnerAction);
                    var initialLoginProgress = Driver.FindElements(Oc.ProgressIndicator);
                    var initialSpinners = Driver.FindElements(Oc.SpinnerInitial);
                    var loadingSpinners = Driver.FindElements(Oc.SpinnerLoading);
                    var matProgressBar = Driver.FindElements(Oc.MatProgressBar);

                    return !(initialLoginProgress.Any(s => IsVisible(s, viewportHeight))
                             || initialSpinners.Any(s => IsVisible(s, viewportHeight))
                             || actionSpinners.Any(s => IsVisible(s, viewportHeight))
                             || loadingSpinners.Any(s => IsVisible(s, viewportHeight))
                             || matProgressBar.Any(s => IsVisible(s, viewportHeight)));
                });

            SetLongImplicitWait();

            // has anything changed?
            var postState = Driver.PageSource.GetHashCode();

            if (postState == initialState)
            {
                // nothing changed so we might be calling everything too early; let's wait and check again
                Thread.Sleep(500);
                ((IJavaScriptExecutor)Driver).ExecuteScript(DocumentReadyStateScript);
            }
        }

        public void WaitForListLoadComplete()
        {
            try
            {
                WaitUntilElementAppears(Oc.MatProgressBar, Constants.ShortTimeOutSeconds);
            }
            catch (WebDriverTimeoutException)
            {
                // ignore
            }

            WaitUntilElementDisappears(Oc.MatProgressBar);
        }

        public void WaitForAnimatedTransitionComplete(IWebElement webElement)
        {
            try
            {
                var previousLocation = new Point();
                var currentLocation = webElement.Location;

                while (currentLocation != previousLocation)
                {
                    previousLocation = currentLocation;
                    Thread.Sleep(TimeSpan.FromMilliseconds(100));
                    currentLocation = webElement.Location;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($@"Error waiting for animated transition to complete: {ex.Message}");
            }
        }

        public void WaitForQueueComplete()
        {
            const int shortTimeout = 10 * 1000;
            const int longTimeout = Constants.LongTimeoutSeconds;
            var cancelToken = new CancellationTokenSource(TimeSpan.FromSeconds(longTimeout)).Token;
            while (!cancelToken.IsCancellationRequested)
            {
                WaitForLoadComplete();

                while (true)
                {
                    try
                    {
                        SetImplicitWait(shortTimeout);
                        Driver.FindElement(Oc.UploadIndicatorCounter);
                        WaitUntilElementDisappears(Oc.UploadIndicatorCounter);
                    }
                    catch (NoSuchElementException)
                    {
                        break;
                    }
                }

                SetLongImplicitWait();
                WaitForListLoadComplete();
                return;
            }

            SetLongImplicitWait();
            throw new TimeoutException($"Upload has not completed within {longTimeout} seconds.");
        }

        public void WaitUntilElementAppears(By elementSelector, int timeoutSeconds = Constants.NormalTimeoutSeconds)
        {
            WaitFor(condition => IsElementDisplayed(elementSelector), timeoutSeconds);
        }

        public void WaitUntilElementDisappears(By elementSelector, IWebElement context = null, int timeoutSeconds = Constants.NormalTimeoutSeconds)
        {
            WaitFor(condition => !IsElementDisplayed(elementSelector, context), timeoutSeconds);
        }

        public void WaitFor(Func<IWebDriver, bool> condition, int timeoutSeconds = Constants.NormalTimeoutSeconds)
        {
            var wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(timeoutSeconds));

            SetShortImplicitWait();

            wait.Until(condition);

            SetLongImplicitWait();
        }

        public void SetShortImplicitWait()
        {
            SetImplicitWait(100);
        }

        public void SetLongImplicitWait()
        {
            SetImplicitWait(FindElementTimeout * 1000);
        }

        public void ReloadOc()
        {
            JustClick(Oc.RefreshIcon);
            WaitUntilElementAppears(Oc.SpinnerLoading);
        }

        public void OpenSettings()
        {
            JustClick(Oc.SettingsButton);
            WaitUntilElementAppears(Oc.OcOptionsLabel);
        }

        public IWebElement JustClick(By elementSelector, IWebElement elementContext = null)
        {
            var element = elementContext == null
                ? Driver.FindElement(elementSelector)
                : elementContext.FindElement(elementSelector);
            element.Click();
            return element;
        }

        public IWebElement WaitAndClick(By elementSelector, IWebElement elementContext = null)
        {
            WaitForLoadComplete();
            return JustClick(elementSelector, elementContext);
        }

        public IWebElement ClickAndWait(By elementSelector, IWebElement elementContext = null)
        {
            var element = JustClick(elementSelector, elementContext);
            WaitForLoadComplete();
            return element;
        }

        public IWebElement WaitAndClickThenWait(By elementSelector, IWebElement elementContext = null)
        {
            WaitForLoadComplete();
            return ClickAndWait(elementSelector, elementContext);
        }

        public void SwitchToLastDriverWindow(WindowHandles handles = WindowHandles.Multiple)
        {
            WaitForDriverWindowAvailable(handles);
            Driver.SwitchTo().Window(Driver.WindowHandles.Last());
        }

        public void SwitchToFirstDriverWindow() => Driver.SwitchTo().Window(Driver.WindowHandles.First());

        public void HoverMouseOverElement(By elementSelector)
        {
            var webElement = Driver.FindElement(elementSelector);
            HoverMouseOverElement(webElement);
        }

        public void HoverMouseOverElement(IWebElement webElement) => new Actions(Driver).MoveToElement(webElement).Perform();

        public Color GetColor(IWebElement webElement, string colorProperty = "color")
        {
            var rgbaValues = webElement.GetCssValue(colorProperty).Split('(', ')')[1];
            var channels = rgbaValues.Split(',').Select(double.Parse).ToArray();
            return Color.FromArgb((int)channels[0], (int)channels[1], (int)channels[2]);
        }

        public Color GetCheckboxColor(IWebElement element) => GetColor(element, "background-color");

        public bool IsElementDisplayed(By elementSelector, IWebElement context = null)
        {
            try
            {
                return context?.FindElement(elementSelector).Displayed ?? Driver.FindElement(elementSelector).Displayed;
            }
            catch
            {
                return false;
            }
        }

        public bool IsElementSelected(By elementSelector) => Driver.FindElement(elementSelector).Selected;

        public IWebElement GetToolTip(IWebElement element)
        {
            CheckIfOtherToolTipIsDisplayed();
            HoverMouseOverElement(element);
            WaitUntilElementAppears(Oc.Tooltip);
            return Driver.FindElement(Oc.Tooltip);
        }

        private void SetImplicitWait(int milliseconds)
        {
            Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(milliseconds);
        }

        private bool IsVisible(IWebElement element, int viewportHeight)
        {
            try
            {
                return element.Displayed
                       && element.Location.Y <= viewportHeight
                       && element.Location.Y + element.Size.Height >= 0;
            }
            catch
            {
                return false;
            }
        }

        private void CheckIfOtherToolTipIsDisplayed()
        {
            if (IsElementDisplayed(Oc.Tooltip))
            {
                HoverMouseOverElement(Oc.MattersTab);
                WaitUntilElementDisappears(Oc.Tooltip);
            }
        }

        private void WaitForDriverWindowAvailable(WindowHandles handles)
        {
            var cancelToken = new CancellationTokenSource(TimeSpan.FromSeconds(Constants.NormalTimeoutSeconds)).Token;
            while (!cancelToken.IsCancellationRequested)
            {
                if (Driver.WindowHandles.Count >= (int)handles)
                {
                    return;
                }

                Thread.Sleep(200);
            }

            throw new Exception("The driver window has not become available within the given time.");
        }
    }
}
