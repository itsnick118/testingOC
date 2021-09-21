using System;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Forms;
using OpenQA.Selenium;
using SimWinInput;
using Keys = OpenQA.Selenium.Keys;

namespace IntegratedDriver
{
    public class UserInput
    {
        public static void LeftClick(AutomationElement element)
        {
            Click(MouseButtons.Left, element);
        }

        public static void LeftClick(IWebElement element)
        {
            try
            {
                element.Click();
            }
            catch
            {
                var point = PointFromPoint(element.Location);
                Click(MouseButtons.Left, point);
            }
        }

        public static void LeftClick(Point point)
        {
            Click(MouseButtons.Left, point);
        }

        public static void RightClick(AutomationElement element)
        {
            Click(MouseButtons.Right, element);
        }

        public static void DoubleClick(AutomationElement element)
        {
            var clickablePoint = PointFromElement(element);
            SimMouse.Click(MouseButtons.Left, clickablePoint.X, clickablePoint.Y);
            Thread.Sleep(50);
            SimMouse.Click(MouseButtons.Left, clickablePoint.X, clickablePoint.Y);
        }

        public static void MoveMouseTo(AutomationElement element)
        {
            var clickablePoint = PointFromElement(element);
            SimMouse.Act(SimMouse.Action.MoveOnly, clickablePoint.X, clickablePoint.Y);
        }

        public static void Scroll(AutomationElement element, int numberOfClicks, bool down=true)
        {
            var clicks = Math.Abs(numberOfClicks);
            if (down) clicks *= -1;
            var clickablePoint = PointFromElement(element);

            InteropMouse.mouse_event(
                (uint)InteropMouse.MouseEventFlags.Wheel,
                clickablePoint.X,
                clickablePoint.Y,
                clicks,
                0);
        }

        public static void DragAndDrop(Point source, Point target, bool switchToWindow = false)
        {
            SimMouse.Act(SimMouse.Action.MoveOnly, (int)source.X, (int)source.Y);
            Thread.Sleep(100);
            SimMouse.Act(SimMouse.Action.LeftButtonDown, (int)source.X, (int)source.Y);
            Thread.Sleep(100);
            if (switchToWindow)
            {
                SendKeys.SendWait("%{TAB}");
                Thread.Sleep(100);
            }
            SimMouse.Act(SimMouse.Action.MoveOnly, (int)source.X + 100, (int)target.Y + 100);
            Thread.Sleep(100);
            SimMouse.Act(SimMouse.Action.MoveOnly, (int)target.X, (int)target.Y);
            Thread.Sleep(100);
            SimMouse.Act(SimMouse.Action.LeftButtonUp, (int)target.X, (int)target.Y);
            Thread.Sleep(100);
            SimMouse.Act(SimMouse.Action.MoveOnly, (int)target.X, (int)target.Y);
        }

        public static void DragAllAndDrop(AutomationElement element, Point target)
        {
            var source = PointFromElement(element);
            Click(MouseButtons.Left, element);
            SelectAll();

            // This prevents the subsequent click from registering as a double click.
            Thread.Sleep(TimeSpan.FromSeconds(1));

            DragAndDrop(PointFromPoint(source), target);
        }

        public static void Type(string inputString)
        {
            SendKeys.SendWait(inputString);
        }

        public static void Type(IWebElement element, string inputString)
        {
            element.SendKeys(inputString);
        }

        public static void DeleteAndType(IWebElement element, string inputString)
        {
            element.SendKeys(Keys.Backspace);
            Type(element, inputString);
        }

        public static void SelectAll()
        {
            Type("^a");
        }

        public static void SelectAll(IWebElement element)
        {
            element.SendKeys(Keys.Control + 'a');
        }

        public static void KeyPress(string inputString)
        {
            Type(inputString);
        }

        public static void PressEscape(IWebElement element)
        {
            element.SendKeys(Keys.Escape);
        }
        public static void PressEnter(IWebElement element)
        {
            element.SendKeys(Keys.Enter);
        }

        private static void Click(MouseButtons button, AutomationElement element)
        {
            var clickablePoint = PointFromElement(element);

            SimMouse.Click(button, clickablePoint.X, clickablePoint.Y);
        }

        private static void Click(MouseButtons button, Point point)
        {
            var convertedPoint = PointFromPoint(point);
            SimMouse.Click(button, convertedPoint.X, convertedPoint.Y);
        }

        private static System.Drawing.Point PointFromElement(AutomationElement element)
        {
            Windows.BringParentWindowToFront(element);
            var point = element.GetClickablePoint();
            return PointFromPoint(point);
        }

        private static System.Drawing.Point PointFromPoint(Point point)
        {
            return new System.Drawing.Point((int)point.X, (int)point.Y);
        }

        private static Point PointFromPoint(System.Drawing.Point point)
        {
            return new Point(point.X, point.Y);
        }
    }
}
