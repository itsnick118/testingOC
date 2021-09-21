using System;
using System.Diagnostics;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using UITests.PageModel.Passport;
using UITests.PageModel.Selectors;
using UITests.PageModel.Shared;

namespace UITests.PageModel
{
    public class OfficeCompanion
    {
        private readonly IAppInstance _app;

        public Header Header { get; }

        public BasicSettingsPage BasicSettingsPage { get; }
        public SettingsPage SettingsPage { get; }
        public MattersListPage MattersListPage { get; }
        public InvoicesListPage InvoicesListPage { get; }
        public InvoiceSummaryPage InvoiceSummaryPage { get; }
        public HelpPage HelpPage { get; }
        public GlobalDocumentsPage GlobalDocumentsPage { get; }
        public MatterDetailsPage MatterDetailsPage { get; }
        public PeopleListPage PeopleListPage { get; }
        public EmailsListPage EmailListPage { get; }
        public DocumentsListPage DocumentsListPage { get; }
        public NarrativesListPage NarrativesListPage { get; }
        public TasksEventsListPage TasksEventsListPage { get; }
        public DocumentSummaryPage DocumentSummaryPage { get; }
        public UploadHistoryPage UploadHistoryPage { get; }
        public MatterPassportPage MatterPassportPage { get; }
        public PassportPreferencesPage PassportPreferencesPage { get; }
        public InvoicePassportPage InvoicePassportPage { get; }
        public SelectPersonDialog SelectPersonDialog { get; }

        public OfficeCompanion(TestEnvironment environment, IWebDriver driver, Process officeAppProcess)
        {
            _app = new AppInstance(environment, officeAppProcess, driver);
            Header = new Header(_app);

            BasicSettingsPage = new BasicSettingsPage(_app);
            SettingsPage = new SettingsPage(_app);
            DocumentsListPage = new DocumentsListPage(_app);
            DocumentSummaryPage = new DocumentSummaryPage(_app);
            InvoicesListPage = new InvoicesListPage(_app);
            InvoiceSummaryPage = new InvoiceSummaryPage(_app);
            MattersListPage = new MattersListPage(_app);
            MatterDetailsPage = new MatterDetailsPage(_app);
            NarrativesListPage = new NarrativesListPage(_app);
            PeopleListPage = new PeopleListPage(_app);
            HelpPage = new HelpPage(_app);
            GlobalDocumentsPage = new GlobalDocumentsPage(_app);
            UploadHistoryPage = new UploadHistoryPage(_app);
            TasksEventsListPage = new TasksEventsListPage(_app);
            MatterPassportPage = new MatterPassportPage(_app);
            PassportPreferencesPage = new PassportPreferencesPage(_app);
            InvoicePassportPage = new InvoicePassportPage(_app);
            EmailListPage = new EmailsListPage(_app);
            SelectPersonDialog = new SelectPersonDialog(_app);
        }

        public void SwitchToLastOcInstance() => _app.SwitchToLastDriverWindow();
        public void SwitchToFirstOcInstance() => _app.SwitchToFirstDriverWindow();

        public void WaitForLoadComplete()
        {
            _app.WaitForLoadComplete();
        }

        public void WaitForQueueComplete()
        {
            _app.WaitForQueueComplete();
        }

        public void CaptureScreen(string testName)
        {
            var screenShot = ((ChromeDriver)_app.Driver).GetScreenshot();
            _app.Environment.SaveToTestOutputDirectory(screenShot, testName);
        }

        public void SaveLogs(string testName)
        {
            var logDir = Path.Combine(_app.Environment.UserDataPath, "Profile", "logs");
            var files = Directory.GetFiles(logDir);

            foreach (var file in files)
            {
                _app.Environment.CopyToTestOutputDirectory(file, testName);
            }
        }

        public void ReloadOc()
        {
            _app.ReloadOc();

            WaitForLoadComplete();
        }

        public void OpenSettings()
        {
            _app.OpenSettings();
        }

        public void SwitchToLastDriverWindow(WindowHandles handles) => _app.SwitchToLastDriverWindow(handles);

        public bool IsErrorDisplayed() => _app.IsElementDisplayed(Oc.ToastMessage);

        public string[] GetAllToastMessages()
        {
            var toastMessages = _app.Driver.FindElements(Oc.ToastMessage);
            string[] messages = new string[toastMessages.Count];
            for (var i = 0; i < toastMessages.Count; i++)
            {
                messages[i] = toastMessages[i].Text;
            }
            return messages;
        }

        public int GetQueuedEmailCount()
        {
            _app.WaitUntilElementAppears(Oc.QueuedEmailCount, Constants.LongTimeoutSeconds);
            var count = Int32.Parse(_app.Driver.FindElement(Oc.QueuedEmailCount).Text);
            return count;
        }

        public void CloseAllToastMessages()
        {
            while (true)
            {
                try
                {
                    var closeButton = _app.Driver.FindElement(Oc.ToastMessageCloseButton);
                    closeButton.Click();
                }
                catch (NoSuchElementException)
                {
                    break;
                }
            }
        }
    }
}
