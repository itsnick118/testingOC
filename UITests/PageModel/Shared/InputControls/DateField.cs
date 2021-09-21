using IntegratedDriver;
using OpenQA.Selenium;
using UITests.PageModel.Selectors;

namespace UITests.PageModel.Shared.InputControls
{
    public class DateField : InputControl
    {
        private readonly IAppInstance _app;

        public DateField(IAppInstance app, string name, string className, int panelNumber = 0) : base(app)
        {
            _app = app;
            Name = name;
            ClassName = className;
            PanelNumber = panelNumber;
        }

        public override string Set(string date)
        {
            OpenPanel();

            var input = GetInputField();
            UserInput.SelectAll(input);
            UserInput.Type(input, date);

            if (!date.Contains(" - "))
            {
                input.Click();
                WaitForCalendar();

                if (date.Contains(":"))
                {
                    _app.ClickAndWait(Oc.CalendarSetButton);
                }
                else
                {
                    _app.ClickAndWait(Oc.CalendarSelectedDate);
                }
            }

            return date;
        }

        public override string GetValue()
        {
            return _app.Driver.FindElement(Oc.DateFieldInputByClass(ClassName)).GetAttribute("value");
        }

        private IWebElement GetInputField() => _app.Driver.FindElement(Oc.DateFieldInputByClass(ClassName));

        private void WaitForCalendar()
        {
            _app.WaitUntilElementAppears(Oc.CalendarControl);
            var calendarControl = _app.Driver.FindElement(Oc.CalendarControl);
            _app.WaitForAnimatedTransitionComplete(calendarControl);
        }
    }
}