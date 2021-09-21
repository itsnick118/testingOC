using System.Windows.Automation;
using IntegratedDriver;
using IntegratedDriver.ElementFinders;
using UITests.PageModel.Selectors;
using UITests.PageModel.Shared;

namespace UITests.PageModel.OfficeApps
{
    public class Excel : OfficeApplication
    {
        public Excel(TestEnvironment testEnvironment) : base(testEnvironment)
        {
        }

        public void ReplaceTextWith(string content)
        {
            SetForegroundWindow();
            AppWindow.SetFocus();
            Wait();

            UserInput.SelectAll();
            UserInput.Type(content);
            Wait();
        }

        internal void ClickTab()
        {
            UserInput.KeyPress("{Tab}");
        }

        public void OpenNewExcel()
        {
            var fileButton = NativeFinder.Find(AppWindow, Native.FileButton, ControlType.Button, 2);
            UserInput.LeftClick(fileButton);

            var word =
                NativeFinder.Find(AppWindow, Native.NewExcelTitle, ControlType.ListItem, 2);
            UserInput.LeftClick(word);
        }
    }
}
