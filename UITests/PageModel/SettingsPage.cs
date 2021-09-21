using System.Drawing;
using OpenQA.Selenium;
using UITests.PageModel.Selectors;
using UITests.PageModel.Shared;
using static UITests.Constants;

namespace UITests.PageModel
{
    public class SettingsPage
    {
        private readonly IAppInstance _app;

        public bool IsCopyEmailsWhenFiling => IsCheckboxSelected(GetCheckbox("copyEmails"));
        public bool IsAutomaticallyThreadLinking => IsCheckboxSelected(GetCheckbox("automaticallyThreadLinking"));

        public SettingsPage(IAppInstance app)
        {
            _app = app;
        }

        public void OpenConfiguration()
        {
            _app.JustClick(Oc.PanelByIndex(1));
        }

        public void OpenAdvanced()
        {
            _app.JustClick(Oc.PanelByIndex(2));
        }

        public void OpenMatterManagement()
        {
            _app.JustClick(Oc.PanelByIndex(3));
        }

        public void Cancel()
        {
            _app.JustClick(Oc.InvoiceItemRejectButton);
        }

        public void Apply()
        {
            _app.JustClick(Oc.ApplyIcon);
        }

        public IDialog LogOut()
        {
            return new Dialog(_app, _app.WaitAndClick(Oc.SignOut));
        }

        public void SelectShowCollapsedOnStart()
        {
            SelectCheckbox("showCollapsedOnStart");
        }

        public void SelectCopyEmailsWhenFiling()
        {
            SelectCheckbox("copyEmails");
        }

        public void SelectEmailThreadLinking()
        {
            SelectCheckbox("automaticallyThreadLinking");
        }

        private void SelectCheckbox(string nameAttribute)
        {
            GetCheckbox(nameAttribute).Click();
        }

        private bool IsCheckboxSelected(IWebElement checkboxElement)
        {
            if (GetCheckboxColor(checkboxElement).Name == BlueColorName)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private IWebElement GetCheckbox(string nameAttribute)
        {
            return _app.Driver.FindElement(Oc.CheckBoxByNameAttribute(nameAttribute)).FindElement(Oc.Parent);
        }

        private Color GetCheckboxColor(IWebElement element)
        {
            return _app.GetCheckboxColor(element.FindElement(Oc.CheckBoxBackground));
        }
    }
}
