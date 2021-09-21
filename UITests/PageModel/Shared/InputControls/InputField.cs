using IntegratedDriver;
using OpenQA.Selenium;
using UITests.PageModel.Selectors;

namespace UITests.PageModel.Shared.InputControls
{
    public class InputField : InputControl
    {
        private readonly IAppInstance _app;
        private readonly By _selector;

        public InputField(IAppInstance app, string name, string className, int panelNumber = 0) : base(app)
        {
            _app = app;
            Name = name;
            ClassName = className;
            PanelNumber = panelNumber;
        }

        public InputField(IAppInstance app, By selector) : base(app)
        {
            _app = app;
            _selector = selector;
        }

        public override string Set(string value)
        {
            OpenPanel();
            var inputField = GetInput();
            inputField.Click();
            UserInput.SelectAll(inputField);
            UserInput.Type(inputField, value);
            return value;
        }

        public override string GetValue()
        {
            return _app.Driver.FindElement(_selector ?? Oc.InputFieldByClass(ClassName)).GetAttribute("value");
        }

        private IWebElement GetInput() =>
            _app.Driver.FindElement(_selector ?? Oc.InputFieldByClass(ClassName));
    }
}