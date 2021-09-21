using OpenQA.Selenium.Interactions;
using PassportSelector = UITests.PageModel.Selectors.Passport;

namespace UITests.PageModel.Passport
{
    public class PassportPreferencesPage : PassportPage
    {
        private bool PersistentListPages => App.IsElementSelected(PassportSelector.PersistentListFlagButton);

        public PassportPreferencesPage(IAppInstance app) : base(app)
        {
        }

        public void SetPersistentListPagesTo(bool flag)
        {
            OpenMyPreferences();
            SwitchToPassportFrame();
            App.WaitUntilElementAppears(PassportSelector.PersistentListFlagButton);

            if (PersistentListPages != flag)
            {
                CheckPersistentListFlag();
            }

        }

        private void CheckPersistentListFlag()
        {
            var flagButton = App.Driver.FindElement(PassportSelector.PersistentListFlagButton);
            var saveButton = App.Driver.FindElement(PassportSelector.SaveEditModeButton);

            var actions = new Actions(App.Driver);
            actions.MoveToElement(flagButton).Click().Perform();
            actions = new Actions(App.Driver);
            actions.MoveToElement(saveButton).Click().Perform();

            App.WaitUntilElementAppears(PassportSelector.EditButton);
        }
    }
}
