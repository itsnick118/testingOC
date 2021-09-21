using IntegratedDriver;
using OpenQA.Selenium;
using UITests.PageModel.Selectors;

namespace UITests.PageModel.Shared.InputControls
{
    public class TextArea : InputControl
    {
        private readonly IAppInstance _app;

        public TextArea(IAppInstance app, string name, string className, int panelNumber = 0) : base(app)
        {
            _app = app;
            Name = name;
            ClassName = className;
            PanelNumber = panelNumber;
        }

        public override string Set(string value)
        {
            OpenPanel();
            try
            {
                var inputField = _app.Driver.FindElement(Oc.TextAreaInputByClass(ClassName));
                inputField.Click();
                UserInput.SelectAll(inputField);
                UserInput.DeleteAndType(inputField, value);
                return value;
            }
            catch (StaleElementReferenceException)
            {
                return Set(value);
            }
        }

        public override string GetValue()
        {
            return _app.Driver.FindElement(Oc.TextAreaInputByClass(ClassName)).GetAttribute("value");
        }

        public override string GetReadOnlyValue() => GetValue();
    }
}