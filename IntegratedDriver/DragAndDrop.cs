using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using IntegratedDriver.ElementFinders;
using OpenQA.Selenium;

namespace IntegratedDriver
{
    public class DragAndDrop
    {
        public static void FromFileSystem(FileInfo file, IWebElement target)
        {
            if (file?.DirectoryName == null)
            {
                return;
            }

            Process.Start("explorer.exe", $"/select, \"{file.FullName}\"");

            var fileName = Path.GetFileNameWithoutExtension(file.FullName);
            var fileDir = file.Directory?.Name;

            var parent = Windows.GetExplorerList(fileDir);
            var item = NativeFinder.FindByPartialMatch(parent, fileName, ControlType.ListItem);

            FromElementToElement(item, target);
            Windows.CloseWindowByName(fileDir);
        }

        public static void AllFilesInFolderDndOC(DirectoryInfo directory, IWebElement target)
        {
            if (directory == null)
            {
                return;
            }

            Process.Start("explorer.exe", $"\"{directory.FullName}\"");
            var files = Directory.GetFiles(directory.FullName);
            if (files.Length == 0)
            {
                throw new IOException($"Directory : {directory.FullName} is Empty");
            }
            var fileName = Path.GetFileNameWithoutExtension(files[0]);
            var directoryName = directory.Name;

            var parent = Windows.GetExplorerList(directoryName);
            var item = NativeFinder.FindByPartialMatch(parent, fileName, ControlType.ListItem);
            UserInput.SelectAll();
            FromElementToElement(item, target);
            Windows.CloseWindowByName(directoryName);
        }

        public static void ToFileSystem(IWebElement element, DirectoryInfo path)
        {
            Process.Start("explorer.exe", $"\"{path.FullName}\"");

            var folder = Windows.GetWindowWithName(path.Name, true);
            var target = NativeFinder.Find(folder, "Items View", ControlType.List);

            FromElementToElement(element, target);
            Windows.WaitUntilFileDownloaded(path + "/email");
            Windows.CloseWindowByName(path.Name);
        }

        public static void FromElementToPoint(AutomationElement sourceElement, Point target)
        {
            Point point;

            while (!sourceElement.TryGetClickablePoint(out point))
            {
                var parent = sourceElement.CachedParent;
                UserInput.Scroll(parent, 5);
            }

            UserInput.DragAndDrop(point, target);
        }

        public static void FromPointToPoint(Point source, Point target) => UserInput.DragAndDrop(source, target);

        public static void FromElementToElement(AutomationElement sourceElement, IWebElement target)
        {
            NativeFinder.WaitForElementReady(sourceElement);
            var targetVisible = WaitForTargetVisible(target);

            if (!targetVisible)
            {
                throw new ElementNotVisibleException("Cannot drag to unavailable target.");
            }

            var ocTargetPoint = new Point(target.Location.X + (target.Size.Width / 2),
                target.Location.Y + (target.Size.Height / 2));
            var targetPoint = ConvertToAbsolutePoint(ocTargetPoint);

            Point point;

            while (!sourceElement.TryGetClickablePoint(out point))
            {
                var parent = TreeWalker.ContentViewWalker.GetParent(sourceElement);
                UserInput.Scroll(parent, 5);
            }

            UserInput.DragAndDrop(point, targetPoint);
        }

        public static void FromElementToElement(IWebElement source, AutomationElement targetElement, bool switchToWindow = true)
        {
            NativeFinder.WaitForElementReady(targetElement);
            var targetVisible = WaitForTargetVisible(source);

            if (!targetVisible)
            {
                throw new ElementNotVisibleException("Cannot drag to unavailable target.");
            }

            var ocTargetPoint = new Point(source.Location.X + (source.Size.Width / 2),
                source.Location.Y + (source.Size.Height / 2));
            var ocPoint = ConvertToAbsolutePoint(ocTargetPoint);

            Point point;

            while (!targetElement.TryGetClickablePoint(out point))
            {
                var parent = TreeWalker.ContentViewWalker.GetParent(targetElement);
                UserInput.Scroll(parent, 5);
            }

            UserInput.DragAndDrop(ocPoint, point, switchToWindow);
        }

        public static void AllFromElementToPoint(AutomationElement source, Point target)
        {
            UserInput.DragAllAndDrop(source, target);
        }

        public static void AllFromElementToElement(AutomationElement sourceElement, IWebElement target)
        {
            var ocTargetPoint = new Point(target.Location.X, target.Location.Y);
            var targetPoint = ConvertToAbsolutePoint(ocTargetPoint);

            UserInput.DragAllAndDrop(sourceElement, targetPoint);
        }

        public static bool WaitForTargetVisible(IWebElement element)
        {
            var timeout = TimeSpan.FromSeconds(20);
            var step = TimeSpan.FromSeconds(1);
            var found = false;

            while (!found && timeout.TotalMilliseconds > 0)
            {
                found = element.Enabled;
                Thread.Sleep(step);
                timeout -= step;
            }

            return found;
        }

        private static Point GetAddinPosition()
        {
            var element = NativeFinder.Find(AutomationElement.RootElement,
                Constants.AddInTitle, ControlType.Pane);
            var boundingRectangle = element.Current.BoundingRectangle;

            return new Point
            {
                X = boundingRectangle.X,
                Y = boundingRectangle.Y
            };
        }

        private static Point ConvertToAbsolutePoint(Point officeCompanionPoint)
        {
            var offset = GetAddinPosition();
            return new Point
            {
                X = offset.X + officeCompanionPoint.X,
                Y = offset.Y + officeCompanionPoint.Y
            };
        }
    }
}
