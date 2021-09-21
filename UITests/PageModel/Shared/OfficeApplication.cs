using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Automation;
using IntegratedDriver;
using IntegratedDriver.ElementFinders;
using OpenQA.Selenium.Chrome;

namespace UITests.PageModel.Shared
{
    public class OfficeApplication : AddInHost
    {
        private const string ReadOnlyTitle = "Read-Only";
        private const string CheckOutButton = "Check Out for Editing";
        private const string ReadOnlyLabel = "This document has been opened read only from Passport.";
        private const string ReadOnlyLabelForCheckedOutDcouments = "This document has been opened read only from Passport. Document is already checked out by sbrown.";
        private const string ExpandBannerLabel = "Office Companion should be expanded to perform any document related operations";
        private const string ExpandButton = "Expand";
        private ChromeDriver _driver;
        private Process _process;
        private string _lastOpenDocumentName;

        protected readonly TestEnvironment Environment;
        protected AutomationElement AppWindow;

        public bool IsReadOnly => GetActualTitle().Contains(ReadOnlyTitle.ToLower());
        public bool IsDocumentOpened => GetActualTitle().Contains(_lastOpenDocumentName.ToLower());
        public string CurrentUserDisplayName => Environment.ElevatedUserDisplayName;

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        public OfficeApplication(TestEnvironment testEnvironment)
        {
            Environment = testEnvironment;
        }

        public void Attach(string documentName)
        {
            documentName = Path.GetFileNameWithoutExtension(documentName);
            AppWindow = Windows.GetWindowWithName(documentName, false);
            _process = Process.GetProcessById(AppWindow.Current.ProcessId);
            _lastOpenDocumentName = documentName;
        }

        public void OpenDocumentFromExplorer(string fileFullPath)
        {
            _process = Launch(fileFullPath);
            var documentName = Path.GetFileNameWithoutExtension(fileFullPath);
            AppWindow = Windows.GetWindowWithName(documentName, false);
            _lastOpenDocumentName = documentName;
        }

        public void AttachToOc()
        {
            if (Oc != null)
            {
                throw new NotSupportedException("Multiple OC instances are not supported.");
            }

            _driver = StartChromeDriver();
            Oc = new OfficeCompanion(Environment, _driver, _process);

            WaitForInputIdle();
        }

        public void CheckOut()
        {
            Click(CheckOutButton, ControlType.Button);
            WaitForSwitchToEditMode();
            Oc?.SwitchToLastDriverWindow(WindowHandles.Single);
        }

        public void SaveDocument()
        {
            AppWindow.SetFocus();
            Wait();

            UserInput.Type("^s");
            Wait();
        }

        public void ClickOnExpandBanner(bool wait = true)
        {
            Wait();
            try
            {
                GetExpandBanner(wait);
                Click(ExpandButton, ControlType.Button);
                Wait();
            }
            catch (ElementNotAvailableException)
            {
                // ignore
            }
        }

        public AutomationElement GetExpandBanner(bool wait = true)
        {
            try
            {
                return wait ? NativeFinder.Find(AppWindow, ExpandBannerLabel) : NativeFinder.Find(AppWindow, ExpandBannerLabel, 50);
            }
            catch (ElementNotAvailableException)
            {
                return null;
            }
        }

        public AutomationElement GetReadOnlyLabel(bool wait = true)
        {
            try
            {
                return wait ? NativeFinder.Find(AppWindow, ReadOnlyLabel) : NativeFinder.Find(AppWindow, ReadOnlyLabel, 50);
            }
            catch (ElementNotAvailableException)
            {
                // ignore
            }

            return null;
        }

        public AutomationElement GetReadOnlyLabelForCheckedOutDocument(bool wait = true)
        {
            try
            {
                return wait ? NativeFinder.Find(AppWindow, ReadOnlyLabelForCheckedOutDcouments) : NativeFinder.Find(AppWindow, ReadOnlyLabelForCheckedOutDcouments, 50);
            }
            catch (ElementNotAvailableException)
            {
                // ignore
            }

            return null;
        }

        public void Close()
        {
            _process.CloseMainWindow();
            _process.WaitForExit((int)TimeSpan.FromSeconds(Constants.NormalTimeoutSeconds).TotalMilliseconds);

            WaitForProcessExit();

            _process.Dispose();
        }

        public void SetForegroundWindow()
        {
            SetForegroundWindow(_process.MainWindowHandle);
        }

        public virtual void CloseDocument()
        {
            AppWindow.SetFocus();
            Wait();

            UserInput.Type("^w");
            Wait();

            WaitForDocumentClose();
        }

        public void Destroy() => Destroy(_process);

        public static void Wait() => Thread.Sleep(500);

        public void WaitForDocumentClose()
        {
            var isClosed = WaitForTrueOrTimeout(() => !GetActualTitle().Contains(_lastOpenDocumentName));
            if (!isClosed)
            {
                throw new Exception("Application document has not closed within the given time.");
            }
        }

        private void Click(string strNameProperty, ControlType controlType)
        {
            var element = NativeFinder.Find(AppWindow, strNameProperty, controlType);
            UserInput.LeftClick(element);
        }

        private string GetActualTitle()
        {
            RefreshHandlers();
            return _process.MainWindowTitle.ToLower();
        }

        private void RefreshHandlers()
        {
            _process.Refresh();
            WaitForInputIdle();
            AppWindow = AutomationElement.FromHandle(_process.MainWindowHandle);
        }

        private void WaitForInputIdle()
        {
            while (_process.MainWindowHandle == IntPtr.Zero)
            {
                _process.WaitForInputIdle();
            }
        }

        private void WaitForProcessExit()
        {
            var exited = WaitForTrueOrTimeout(() => _process.HasExited);
            if (!exited)
            {
                _process.Kill();
                Console.WriteLine(@"Application has not exited within the given time.");
            }
        }

        private void WaitForSwitchToEditMode()
        {
            var isEditMode = WaitForTrueOrTimeout(() => !IsReadOnly && IsDocumentOpened);
            if (!isEditMode)
            {
                throw new Exception("Application has not switched to Edit mode within the given time.");
            }
        }

        private static bool WaitForTrueOrTimeout(Func<bool> function)
        {
            var cancelToken = new CancellationTokenSource(TimeSpan.FromSeconds(Constants.NormalTimeoutSeconds)).Token;
            while (!cancelToken.IsCancellationRequested)
            {
                if (function())
                {
                    return true;
                }

                Wait();
            }

            return false;
        }
    }
}
