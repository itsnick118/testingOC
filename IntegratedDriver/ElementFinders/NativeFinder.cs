using System;
using System.Threading;
using System.Windows.Automation;
using ControlType = System.Windows.Automation.ControlType;

namespace IntegratedDriver.ElementFinders
{
    public class NativeFinder
    {
        private const int RetryLimit = 100;

        public static AutomationElement Find(AutomationElement parent, string name, int retryLimit = RetryLimit)
        {
            var nameCondition = new PropertyCondition(AutomationElement.NameProperty, name);
            return WaitFor(nameCondition, parent, retryLimit);
        }

        public static AutomationElement Find(AutomationElement parent, string name,
            ControlType controlType, int retryLimit = RetryLimit)
        {
            var multipleCondition = new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, name),
                new PropertyCondition(AutomationElement.ControlTypeProperty, controlType)
            );

            return WaitFor(multipleCondition, parent, retryLimit);
        }

        public static AutomationElement Find(AutomationElement parent, ControlType controlType, int retryLimit = RetryLimit)
        {
            var typeCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, controlType);
            return WaitFor(typeCondition, parent, retryLimit);
        }

        public static AutomationElement FindByPartialMatch(AutomationElement parent, string name,
            ControlType controlType,
            int retryLimit = RetryLimit)
        {
            if (parent == null)
            {
                parent = AutomationElement.RootElement;
            }

            var tries = 0;
            AutomationElement result = null;

            while (result == null && tries < retryLimit)
            {
                var children = FindAll(parent, controlType);

                if (children != null)
                {
                    foreach (AutomationElement childElement in children)
                    {
                        if (childElement.Current.Name.Contains(name))
                        {
                            result = childElement;
                        }
                    }
                }

                tries++;
                Thread.Sleep(200);
            }

            if (tries >= retryLimit)
            {
                throw new ElementNotAvailableException(
                    $@"Retry limit for partial match on name {name} exceeded.");
            }

            return result;
        }

        public static AutomationElementCollection FindAll(AutomationElement parent,
            ControlType controlType, int retryLimit = RetryLimit)
        {
            var controlTypeCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, controlType);

            return WaitForAll(controlTypeCondition, parent, retryLimit);
        }

        public static void WaitForElementReady(AutomationElement element)
        {
            WindowPattern windowPattern;

            try
            {
                windowPattern =
                    element.GetCurrentPattern(WindowPattern.Pattern)
                        as WindowPattern;
            }
            catch (InvalidOperationException)
            {
                Console.WriteLine(
                    @"Could not wait for an unwaitable element: " +
                    $@"{element.Current.Name}");
                return;
            }

            if (windowPattern == null || windowPattern.WaitForInputIdle(10000))
            {
                Console.WriteLine(
                    @"Timeout or null pattern waiting for an unwaitable element: " +
                    $@"{element.Current.Name}");
            }
        }

        private static AutomationElement WaitFor(Condition condition, AutomationElement parent, int retryLimit)
        {
            if (parent == null)
            {
                parent = AutomationElement.RootElement;
            }

            var tries = 0;
            AutomationElement result = null;

            while (result == null && tries < retryLimit)
            {
                try
                {
                    var first = parent.FindFirst(TreeScope.Descendants, condition);
                    if (first == null)
                    {
                        Thread.Sleep(200);
                        tries++;
                    }
                    result = first;
                }
                catch
                {
                    Thread.Sleep(200);
                    tries++;
                }
            }

            if (tries >= retryLimit)
            {
                throw new ElementNotAvailableException(
                    $@"Retry limit for {condition.GetType().Name} " +
                    $@"from {parent.Current.Name} exceeded.");
            }

            return result;
        }

        private static AutomationElementCollection WaitForAll(Condition condition, AutomationElement parent, int retryLimit)
        {
            if (parent == null)
            {
                parent = AutomationElement.RootElement;
            }

            for (var tries = 0; tries < retryLimit; tries++)
            {
                try
                {
                    var tempResult = parent.FindAll(TreeScope.Descendants, condition);
                    if (tempResult.Count > 0) return tempResult;
                }
                catch
                {
                    // ignored
                }

                Thread.Sleep(100);
            }

            Console.WriteLine($@"Retry limit for {condition.GetType().Name} " +
                              $@"from {parent.Current.Name} exceeded; proceeding anyway.");
            return null;
        }
    }
}