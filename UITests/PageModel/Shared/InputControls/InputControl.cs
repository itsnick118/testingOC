using OpenQA.Selenium;
using System;
using UITests.PageModel.Selectors;

namespace UITests.PageModel.Shared.InputControls
{
    public abstract class InputControl
    {
        private readonly IAppInstance _app;
        private readonly By _requiredFieldMessage = By.ClassName("help-block");

        private static By SelectPanelByIndex(int index) =>
            By.XPath($"(//div[contains(@class, \"ms-Dialog-main\")]//mat-expansion-panel-header)[{index + 1}]");

        private static By DialogLabelByText(string labelName) =>
            By.XPath($"//div[contains(text(),'{labelName}')]");

        private static By DialogFieldContainerByClass(string className) =>
            By.CssSelector($".entityform-field-{className}");

        protected InputControl(IAppInstance app)
        {
            _app = app;
        }

        public string Name { get; set; }
        public string ClassName { get; set; }
        public int PanelNumber { get; set; }

        public abstract string Set(string value);

        public abstract string GetValue();

        public virtual string SetByIndex(int index, bool clearInput = false)
        {
            throw new NotSupportedException();
        }

        public virtual string SetValueOtherthan(string value, bool clearInput = false)
        {
            throw new NotSupportedException();
        }

        public virtual string GetReadOnlyValue()
        {
            return _app.Driver.FindElement(Oc.ReadOnlyDialogControlByClass(ClassName)).Text;
        }

        public virtual bool IsRequired()
        {
            var dialogLabel = _app.Driver.FindElement(DialogLabelByText(Name));
            return dialogLabel.GetAttribute("class").Contains("is-required");
        }

        public virtual string GetRequiredWarning()
        {
            var containerField = _app.Driver.FindElement(DialogFieldContainerByClass(ClassName));
            return containerField.FindElement(_requiredFieldMessage).Text;
        }

        public virtual void SelectPersonDialog()
        {
            throw new NotSupportedException();
        }

        public virtual void Clear()
        {
            throw new NotSupportedException();
        }

        protected void OpenPanel()
        {
            if (PanelNumber != 0 && _app.Driver.FindElement(SelectPanelByIndex(PanelNumber))
                    .GetAttribute("aria-expanded") == false.ToString().ToLower())
            {
                var dialog = _app.Driver.FindElement(Oc.Dialog);
                _app.JustClick(SelectPanelByIndex(PanelNumber));
                _app.WaitForAnimatedTransitionComplete(dialog);
            }
        }
    }
}