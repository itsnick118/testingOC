using IntegratedDriver;
using OpenQA.Selenium.Interactions;
using PassportSelector = UITests.PageModel.Selectors.Passport;

namespace UITests.PageModel.Passport
{
    public class MatterPassportPage : PassportPage
    {
        public MatterPassportPage(IAppInstance app) : base(app)
        {
        }

        public void AddMatter(string matterName, bool assignmentWorkflow = false)
        {
            App.SwitchToLastDriverWindow();
            SwitchToPassportFrame();

            var matterNameInput = App.Driver.FindElement(PassportSelector.InputByDataNameTag("Name"));
            UserInput.Type(matterNameInput, matterName);

            if (assignmentWorkflow)
            {
                WaitAndClick(PassportSelector.InputByDataNameTag("Assignment Workflow"));
            }

            WaitAndClick(PassportSelector.InputLookupIconByDataNameTag("Practice Area Business Unit"));
            WaitAndClick(PassportSelector.AutocompleteDialogItemByIndex(0));
            WaitAndClick(PassportSelector.InputLookupIconByDataNameTag("Primary Matter Type"));
            WaitAndClick(PassportSelector.AutocompleteDialogItemByIndex(0));

            if (assignmentWorkflow)
            {
                WaitAndClick(PassportSelector.InputLookupIconByDataNameTag("Assignment Workflow Contact"));
            }
            else
            {
                WaitAndClick(PassportSelector.InputLookupIconByDataNameTag("Primary Internal Contact"));
            }
            WaitAndClick(PassportSelector.AutocompleteDialogItemByIndex(0));

            var saveButton = App.Driver.FindElement(PassportSelector.SaveAddModeButton);
            new Actions(App.Driver).MoveToElement(saveButton).Click().Build().Perform();

            App.WaitUntilElementDisappears(PassportSelector.SaveAddModeButton);
            App.SwitchToFirstDriverWindow();
        }

        public string GetMatterStatus()
        {
            App.SwitchToLastDriverWindow();
            SwitchToPassportFrame();
            App.WaitUntilElementAppears(PassportSelector.MatterStatusText);
            var matterStatusInPassport = App.Driver.FindElement(PassportSelector.MatterStatusText).Text;
            CloseWindowHandleSwitchToOc();
            return matterStatusInPassport.Trim();
        }
    }
}