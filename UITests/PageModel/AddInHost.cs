using IntegratedDriver;
using IntegratedDriver.ElementFinders;
using OpenQA.Selenium.Chrome;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Automation;
using UITests.PageModel.Selectors;
using static IntegratedDriver.Constants;

namespace UITests.PageModel
{
    public class AddInHost
    {
        public OfficeCompanion Oc { get; set; }

        public void Destroy(Process process)
        {
            try
            {
                process.Kill();
                process.WaitForExit();
            }
            catch (InvalidOperationException) { }
            catch (Win32Exception) { }
        }

        protected Process Launch(string executable)
        {
            var startInfo =
                new ProcessStartInfo(executable)
                {
                    WindowStyle = ProcessWindowStyle.Maximized
                };

            var process = Process.Start(startInfo);
            process?.WaitForInputIdle((int)TimeSpan.FromSeconds(Constants.NormalTimeoutSeconds).TotalMilliseconds);

            return process;
        }

        protected ChromeDriver StartChromeDriver()
        {
            var service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;

            var options = new ChromeOptions
            {
                DebuggerAddress = "localhost:20480"
            };

            var driver = new ChromeDriver(service, options);

            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(FindElementTimeout);

            return driver;
        }

        protected AutomationElement GetTaskPane(AutomationElement parent)
        {
            if (parent == null)
            {
                throw new NullReferenceException(
                "GetTaskPane() requires a parent to prevent misidentification of objects.");
            }

            return NativeFinder.Find(parent, Native.AddInTaskPane);
        }

        protected Point GetTaskPaneToggle(AutomationElement parent)
        {
            if (parent == null)
            {
                throw new NullReferenceException(
                "GetTaskPaneToggle() requires a parent to prevent misidentification of objects.");
            }

            const int inset = 12;
            var element = NativeFinder.Find(parent, Native.AddInTaskPane);
            var bounds = element.Current.BoundingRectangle;
            return new Point((int)bounds.Right - inset, (int)bounds.Top + inset);
        }

        protected void SetOcView(AutomationElement parent, OcView view)
        {
            const int offset = 100;

            var ocPane = GetTaskPane(parent).Current.BoundingRectangle;
            var boundaryPoint = new Point(ocPane.X + 2, ocPane.Height / 2);
            var currentWidth = ocPane.Width;
            var transition = OcTransitionWidth - currentWidth;

            switch (view)
            {
                case OcView.Wide:
                    if (currentWidth < OcTransitionWidth)
                    {
                        DragAndDrop.FromPointToPoint(boundaryPoint, new Point(ocPane.X - (offset + transition), ocPane.Height / 2));
                    }
                    break;

                case OcView.Narrow:
                    if (currentWidth >= OcTransitionWidth)
                    {
                        DragAndDrop.FromPointToPoint(boundaryPoint, new Point(ocPane.X + (offset - transition), ocPane.Height / 2));
                    }
                    break;
            }
        }
    }
}