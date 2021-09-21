using IntegratedDriver;
using IntegratedDriver.ElementFinders;
using System.Windows.Automation;
using UITests.PageModel.Selectors;
using UITests.PageModel.Shared;
using static IntegratedDriver.Constants;

namespace UITests.PageModel
{
    public class OutlookEmailForm : OfficeApplication
    {
        public OutlookEmailForm(TestEnvironment testEnvironment) : base(testEnvironment)
        {
        }

        public string GetEmailRecipientTo()
        {
            var to = NativeFinder.Find(GetEmailFormEditElement(RecepientTo), ControlType.Button).Current.Name;
            return to.Split(',')[0];
        }

        public string GetEmailFormValue(string controlName)
        {
            var element = GetEmailFormEditElement(controlName);
            var valuePattern = element.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
            if (valuePattern != null)
            {
                return valuePattern.Current.Value;
            }
            else
            {
                return string.Empty;
            }
        }

        public AutomationElement GetEmailPageElement()
        {
            return GetEmailFormEditElement(NewEmailPageContent);
        }

        public AutomationElement GetAttachment(string filename)
        {
            return NativeFinder.FindByPartialMatch(AppWindow, filename, ControlType.Button, 150);
        }

        public void Reply()
        {
            UserInput.LeftClick(NativeFinder.Find(AppWindow, Native.Reply, ControlType.Button));
        }

        public void UpdateSubject(string subject)
        {
            var subjectElement = GetEmailFormEditElement(EmailSubject);
            subjectElement.SetFocus();
            UserInput.SelectAll();
            UserInput.Type(subject);
        }

        public void Send()
        {
            UserInput.LeftClick(NativeFinder.Find(AppWindow, Native.Send, ControlType.Button));
        }

        public override void CloseDocument()
        {
            UserInput.Type("{ESC}");
            Wait();

            WaitForDocumentClose();
        }

        private AutomationElement GetEmailFormEditElement(string controlName)
        {
            return NativeFinder.Find(AppWindow, controlName, ControlType.Edit);
        }
    }
}
