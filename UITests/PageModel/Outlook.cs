using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Windows.Automation;
using IntegratedDriver;
using IntegratedDriver.ElementFinders;
using OpenQA.Selenium;
using UITests.PageModel.Selectors;
using static IntegratedDriver.Constants;
using ControlType = System.Windows.Automation.ControlType;
using OUTLOOK = Microsoft.Office.Interop.Outlook;

namespace UITests.PageModel
{
    public class Outlook : AddInHost
    {
        private readonly TestEnvironment _environment;
        private AutomationElement _outlookWindow;
        private OutlookCalendars _calendars;
        private OutlookContacts _contacts;
        private readonly EmailGenerator _emailGenerator = new EmailGenerator();

        public Process Process;
        public OutlookCalendars Calendars => GetOutlookCalendarsInstance();
        public OutlookContacts Contacts => GetOutlookContactInstance();
        public string CurrentUserDisplayName => _environment.ElevatedUserDisplayName;
        public string CurrentUserPrimaryPABU => _environment.ElevatedUserPrimaryPABU;
        public string CurrentUserSecondaryPABU => _environment.ElevatedUserSecondaryPABU;

        public Outlook(TestEnvironment testEnvironment)
        {
            _environment = testEnvironment;
        }

        public void Launch()
        {
            Process = Launch(@"OUTLOOK.EXE");

            var driver = StartChromeDriver();
            Oc = new OfficeCompanion(_environment, driver, Process);

            while (Process.MainWindowHandle == IntPtr.Zero)
            {
                Thread.Sleep(TimeSpan.FromSeconds(1));
            }
            _outlookWindow = AutomationElement.FromHandle(Process.MainWindowHandle);
        }

        public void Destroy()
        {
            Destroy(Process);
        }

        public void DragAndDropEmailToOc(AutomationElement mailItem, IWebElement targetPoint, bool wait = true)
        {
            DragAndDrop.FromElementToElement(mailItem, targetPoint);
            if (wait)
            {
                Oc.WaitForQueueComplete();
            }
        }

        public void OpenTestEmailFolder()
        {
            UserInput.LeftClick(GetTempOutlookFolder());
            UserInput.LeftClick(GetCurrentMailList());
        }

        private void ClickViewMoreOnMsExchangeHyperLink()
        {
            var testFoldertableView = NativeFinder.Find(_outlookWindow, Native.OutlookItemList, ControlType.Table, 10);
            var loadEmailsFromServer = NativeFinder.FindByPartialMatch(testFoldertableView, ViewMoreOnMsExchange, ControlType.Button, 10);
            try
            {
                if (loadEmailsFromServer != null)
                {
                    UserInput.LeftClick(loadEmailsFromServer);
                }
            }
            catch (ElementNotAvailableException)
            {
                // Do nothing
            }
        }

        public IDictionary<string, string> AddTestEmailsToFolder(int numEmails, FileSize fileSize = FileSize.VerySmall, bool asAttachment = false,
            OfficeApp docType = OfficeApp.Notepad, string subject = null, string fileContent = null, bool useDifferentTemplates = false) =>
            _emailGenerator.AddTestEmailsToFolder(numEmails, fileSize, asAttachment, docType, subject, fileContent, useDifferentTemplates);

        public AutomationElement OpenNthTestEmail(int n)
        {
            var element = GetNthEmailInTestFolder(n);
            var mailName = element.Current.Name;
            var id = MailSubjectPattern.Match(mailName).Groups[1].Value;

            var retryTimes = 10;
            var gotClickablePoint = false;

            while (retryTimes-- > 0)
            {
                try
                {
                    Windows.BringParentWindowToFront(element);

                    gotClickablePoint = element.TryGetClickablePoint(out _);
                    if (gotClickablePoint)
                    {
                        UserInput.DoubleClick(element);
                        break;
                    }
                }
                finally
                {
                    Thread.Sleep(500);
                }
            }

            if (!gotClickablePoint)
            {
                throw new ElementNotAvailableException();
            }

            var openedElement = Windows.GetWindowWithName(id, false);
            return openedElement;
        }

        public void SelectNthItem(int n)
        {
            UserInput.LeftClick(GetNthItem(n));
        }

        public AutomationElement GetNthItem(int n, int retryLimit = RetryLimit)
        {
            var allElements = NativeFinder.FindAll(_outlookWindow, ControlType.DataItem, retryLimit);
            try
            {
                return allElements[n];
            }
            catch
            {
                return null;
            }
        }

        public void SelectAllItems()
        {
            UserInput.SelectAll();
        }

        public bool ToggleTaskPane()
        {
            return ToggleTaskPane(_outlookWindow);
        }

        public bool ToggleTaskPane(AutomationElement parent)
        {
            var widthBefore = GetTaskPane(parent).Current.BoundingRectangle.Width;
            var point = GetTaskPaneToggle(parent);
            UserInput.LeftClick(point);
            NativeFinder.WaitForElementReady(GetTaskPane(parent));

            var widthAfter = GetTaskPane(parent).Current.BoundingRectangle.Width;
            var isNowExpanded = widthAfter > widthBefore;

            return isNowExpanded;
        }

        public void OpenTaskPaneIfClosed()
        {
            var width = GetTaskPane(_outlookWindow).Current.BoundingRectangle.Width;
            if (width > 100) return;
            if (ToggleTaskPane()) return;
            ToggleTaskPane();
        }

        public void CreateMailItem()
        {
            var _newoutlook = new OUTLOOK.Application();
            OUTLOOK.MailItem mailItem = (OUTLOOK.MailItem)
            _newoutlook.CreateItem(OUTLOOK.OlItemType.olMailItem);
            mailItem.Subject = "This is the subject";
            mailItem.To = "someone@example.com";
            mailItem.Body = "This is the message.";
            mailItem.Importance = OUTLOOK.OlImportance.olImportanceLow;
            mailItem.Display(false);
        }

        public AutomationElement GetNewMailBodyElement(AutomationElement newMailwindow)
        {
            return NativeFinder.Find(newMailwindow, NewEmailPageContent, ControlType.Edit, 10);
        }

        public AutomationElement OpenNewMailWindow()
        {
            CreateMailItem();
            return GetWindowWithSubject("This is the subject");
        }

        public AutomationElement GetWindowWithSubject(string subject)
        {
            var openedElement = Windows.GetWindowWithName(subject, false);
            return openedElement;
        }

        public OutlookEmailForm OpenNewEmail()
        {
            var newEmailButton = NativeFinder.Find(_outlookWindow, Native.NewEmail, ControlType.Button);
            UserInput.LeftClick(newEmailButton);
            return new OutlookEmailForm(_environment);
        }

        public void SetOcView(OcView view)
        {
            SetOcView(_outlookWindow, view);
        }

        public void CloseInspector(AutomationElement parent)
        {
            var retries = 10;
            AutomationElement closeButton = null;

            while (retries-- > 0)
            {
                try
                {
                    closeButton = NativeFinder.Find(parent, Native.CloseButton, ControlType.Button);
                    break;
                }
                catch
                {
                    var pattern = (WindowPattern)parent.GetCurrentPattern(WindowPattern.Pattern);
                    pattern.SetWindowVisualState(WindowVisualState.Normal);
                }
            }

            if (closeButton != null)
            {
                UserInput.LeftClick(closeButton);
            }
        }

        public void CloseAllInspectors()
        {
            var inspectors = Windows.GetWindowsWithName(TestEmailPrefix);
            foreach (var inspector in inspectors)
            {
                Windows.BringParentWindowToFront(inspector);
                CloseInspector(inspector);
            }
        }

        public AutomationElement GetNthEmailInTestFolder(int n)
        {
            var element = NativeFinder.FindAll(_outlookWindow, ControlType.DataItem)[n];

            return element;
        }

        public void TurnOnReadingPane()
        {
            try
            {
                var viewTab = NativeFinder.Find(_outlookWindow, Native.ViewTab, ControlType.TabItem, 2);
                UserInput.LeftClick(viewTab);

                var readingPaneButton =
                    NativeFinder.Find(_outlookWindow, Native.ReadingPaneButton, ControlType.MenuItem, 2);
                UserInput.LeftClick(readingPaneButton);

                var readingPaneRightItem =
                    NativeFinder.Find(_outlookWindow, Native.ReadingPaneRight, ControlType.MenuItem, 2);
                UserInput.LeftClick(readingPaneRightItem);
            }
            catch
            {
                Console.WriteLine(@"Could not turn on reading pane.");
            }
        }

        public AutomationElement GetAttachmentFromReadingPane(string filename)
        {
            return NativeFinder.FindByPartialMatch(_outlookWindow, filename, ControlType.Button);
        }

        public AutomationElement GetCurrentMailList()
        {
            return NativeFinder.Find(_outlookWindow, Native.OutlookItemList, ControlType.Table);
        }

        private AutomationElement GetTempOutlookFolder()
        {
            Select(_outlookWindow);
            var folderList = NativeFinder.Find(_outlookWindow, Native.OutlookFolderList);
            Select(folderList);
            var outlookFolderElement = NativeFinder.FindByPartialMatch(folderList, TestEmailFolderName, ControlType.TreeItem, 10);
            if (outlookFolderElement != null) return outlookFolderElement;

            var inboxFolderElement = NativeFinder.Find(folderList, InboxFolderName);
            UserInput.RightClick(inboxFolderElement);

            var newFolderMenuItem = NativeFinder.Find(_outlookWindow, NewFolderMenuItemName, ControlType.MenuItem);
            UserInput.LeftClick(newFolderMenuItem);

            var newFolderNameBox = NativeFinder.Find(folderList, string.Empty, ControlType.Edit);
            newFolderNameBox.SetFocus();
            UserInput.Type(TestEmailFolderName + "{ENTER}");
            return NativeFinder.Find(folderList, TestEmailFolderName);
        }

        private OutlookCalendars GetOutlookCalendarsInstance()
        {
            if (_outlookWindow == null)
            {
                throw new InvalidOperationException("Outlook should be launched before accessing calendars");
            }

            return _calendars ?? (_calendars = new OutlookCalendars(_outlookWindow));
        }

        private OutlookContacts GetOutlookContactInstance()
        {
            if (_outlookWindow == null)
            {
                throw new InvalidOperationException("Outlook should be launched before accessing contacts");
            }

            return _contacts ?? (_contacts = new OutlookContacts(_outlookWindow));
        }

        private void Select(AutomationElement item)
        {
            try
            {
                var selectionItemPattern = item?.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                selectionItemPattern?.Select();
            }
            catch (InvalidOperationException)
            {
                Console.WriteLine($@"Unable to Select {item?.Current.Name}, proceeding anyway.");
            }
        }
    }
}
