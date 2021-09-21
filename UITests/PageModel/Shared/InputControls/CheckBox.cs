using OpenQA.Selenium;
using UITests.PageModel.Selectors;

namespace UITests.PageModel.Shared.InputControls
{
    public class CheckBox : InputControl
    {
        private readonly IAppInstance _app;

        public CheckBox(IAppInstance app, string name, string className, int panelNumber = 0) : base(app)
        {
            _app = app;
            Name = name;
            ClassName = className;
            PanelNumber = panelNumber;
        }

        public override string Set(string value)
        {
            OpenPanel();
            var checkboxInnerInput = GetCheckBoxInnerInput();

            if (checkboxInnerInput.Selected != value.Equals("Yes"))
            {
                var checkbox = GetCheckBox();
                checkbox.Click();
                _app.WaitForLoadComplete();
            }

            return value;
        }

        public override string GetValue()
        {
            return GetCheckBoxInnerInput().Selected ? "Yes" : "No";
        }

        private IWebElement GetCheckBox() =>
            _app.Driver.FindElement(Oc.CheckBoxInputByClass(ClassName));

        private IWebElement GetCheckBoxInnerInput() =>
            _app.Driver.FindElement(Oc.CheckBoxInnerInputByClass(ClassName));
    }
}
