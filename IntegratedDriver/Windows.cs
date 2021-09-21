using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Windows.Automation;
using IntegratedDriver.ElementFinders;

namespace IntegratedDriver
{
    public class Windows
    {
        public static AutomationElement GetWindowWithName(string name, bool exactMatch, int retryLimit = Constants.RetryLimit)
        {
            var controlTypeCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);

            for (var tries = 0; tries < retryLimit; tries++)
            {
                var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, controlTypeCondition);
                foreach (AutomationElement window in windows)
                {
                    try
                    {
                        var match = exactMatch
                            ? window.Current.Name == name
                            : window.Current.Name.Contains(name);

                        if (match) return window;
                    }
                    catch (ElementNotAvailableException)
                    {
                        // ignore
                    }
                }

                Thread.Sleep(100);
            }

            Console.WriteLine($"Retry limit for window {name} from RootElement exceeded; proceeding anyway.");

            return null;
        }

        public static IList<AutomationElement> GetWindowsWithName(string name)
        {
            var controlTypeCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
            var collection = new List<AutomationElement>();

            var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, controlTypeCondition);
            foreach (AutomationElement window in windows)
            {
                var match = window.Current.Name.Contains(name);

                if (match) collection.Add(window);
            }

            return collection;
        }

        public static void CloseWindowByName(string name)
        {
            var window = GetWindowWithName(name, false);

            object pattern;
            var success = window.TryGetCurrentPattern(WindowPattern.Pattern, out pattern);
            if (success)
            {
                ((WindowPattern)pattern).Close();
            }
        }

        public static AutomationElement GetExplorerList(string name)
        {
            var parent = GetWindowWithName(name, false);
            return NativeFinder.Find(parent, ControlType.List);
        }

        public static DirectoryInfo GetWorkingTempFolder()
        {
            var path = new DirectoryInfo(Path.Combine(Path.GetTempPath(), "automation"));
            if (!path.Exists)
            {
                path.Create();
            }
            return path;
        }

        public static void ClearWorkingTempFolder()
        {
            var directory = GetWorkingTempFolder();

            foreach (var file in directory.EnumerateFiles())
            {
                file.Delete();
            }

            foreach (var subDir in directory.EnumerateDirectories())
            {
                subDir.Delete(true);
            }
        }

        public static void BringParentWindowToFront(AutomationElement element)
        {
            var tree = TreeWalker.ControlViewWalker;

            while (!Equals(element.Current.ControlType, ControlType.Window))
            {
                element = tree.GetParent(element);
                if (element == null) break;
            }

            try
            {
                element?.SetFocus();
            }
            catch
            {
                Console.WriteLine($@"Could not bring parent window to front for element {element?.Current.Name}");
            }
        }

        public static void WaitUntilFileDownloaded(string fileFullPath)
        {
            var timeout = Constants.FileDownloadTimeout;
            var fileInfo = new FileInfo(fileFullPath);

            while (timeout >= 0)
            {
                if (!IsFileLocked(fileInfo))
                {
                    return;
                }

                timeout--;
                Thread.Sleep(1000);
            }
        }

        private static bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                return true;
            }
            finally
            {
                stream?.Close();
            }

            return false;
        }
    }
}
