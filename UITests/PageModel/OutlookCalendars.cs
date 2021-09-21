using System;
using IntegratedDriver;
using IntegratedDriver.ElementFinders;
using System.Threading;
using System.Windows.Automation;
using UITests.PageModel.Selectors;

namespace UITests.PageModel
{
    public class OutlookCalendars
    {
        private const int ShortRetryLimit = 20;

        private readonly AutomationElement _outlookWindow;

        public OutlookCalendars(AutomationElement outlookWindow)
        {
            _outlookWindow = outlookWindow;
        }

        public void Open()
        {
            var calendarButton = NativeFinder.Find(_outlookWindow, Native.CalendarsTab, ControlType.Button);
            UserInput.LeftClick(calendarButton);
            Wait();
        }

        public string[] GetPassportCalendarsList()
        {
            var matterCalendars = GetPassportCalendars();
            if (matterCalendars == null)
            {
                return new string[] { };
            }

            var result = new string[matterCalendars.Count];

            for (var i = 0; i < matterCalendars.Count; i++)
            {
                result[i] = matterCalendars[i].Current.Name;
            }

            return result;
        }

        public AutomationElement GetAppointment(string subject)
        {
            var appointment = NativeFinder.FindByPartialMatch(_outlookWindow, subject, ControlType.ListItem);
            return appointment;
        }

        public void RemovePassportCalendar(string calendarName)
        {
            Open();

            var calendars = GetPassportCalendars();
            if (calendars == null)
            {
                return;
            }

            ScrollCalendarFoldersToBottom();

            foreach (AutomationElement calendar in calendars)
            {
                if (calendar.Current.Name != calendarName)
                {
                    continue;
                }

                UserInput.RightClick(calendar);
                Wait();

                var deleteButton = NativeFinder.Find(_outlookWindow, Native.DeleteCalendarMenuItem, ControlType.MenuItem);
                UserInput.LeftClick(deleteButton);
                Wait();

                var confirmationDialog = NativeFinder.Find(_outlookWindow, Native.OutlookDialogTitle, ControlType.Window);
                var yesButton = NativeFinder.Find(confirmationDialog, Native.OutlookDialogYesButton, ControlType.Button);
                UserInput.LeftClick(yesButton);
            }
        }

        private AutomationElementCollection GetPassportCalendars()
        {
            try
            {
                var passportCalendarsGroup = NativeFinder.Find(_outlookWindow, Native.PassportCalendarsGroup, ControlType.TreeItem, ShortRetryLimit);
                ExpandCalendarsGroup(passportCalendarsGroup);
                var matterCalendars = NativeFinder.FindAll(passportCalendarsGroup, ControlType.TreeItem);
                return matterCalendars;
            }
            catch
            {
                return null;
            }
        }

        private void ScrollCalendarFoldersToBottom()
        {
            var tree = TreeWalker.ControlViewWalker;

            try
            {
                var calendarFolders = NativeFinder.Find(_outlookWindow, Native.CalendarFoldersPane, ControlType.Tree, ShortRetryLimit);
                var calendarFoldersPane = tree.GetParent(calendarFolders);

                UserInput.MoveMouseTo(calendarFolders);
                Wait();

                if (!(calendarFoldersPane.GetCurrentPattern(ScrollPattern.Pattern) is ScrollPattern scrollPattern))
                {
                    return;
                }

                while (scrollPattern.Current.VerticallyScrollable && scrollPattern.Current.VerticalScrollPercent < 100)
                {
                    scrollPattern.ScrollVertical(ScrollAmount.LargeIncrement);
                    Wait();
                }
            }
            catch
            {
                Console.WriteLine(@"Could not scroll calendar folders.");
            }
        }

        private static void ExpandCalendarsGroup(AutomationElement calendarsGroup)
        {
            try
            {
                if (!(calendarsGroup.GetCurrentPattern(ExpandCollapsePattern
                        .Pattern) is ExpandCollapsePattern pattern) ||
                    pattern.Current.ExpandCollapseState == ExpandCollapseState.Expanded)
                {
                    return;
                }

                pattern.Expand();
                Wait();
            }
            catch
            {
                Console.WriteLine(@"Could not expand calendars group.");
            }
        }

        private static void Wait(int ms = 1000) => Thread.Sleep(ms);
    }
}
