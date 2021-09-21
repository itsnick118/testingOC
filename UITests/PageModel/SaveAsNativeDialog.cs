using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;
using IntegratedDriver;
using IntegratedDriver.ElementFinders;

namespace UITests.PageModel
{
    public class SaveAsNativeDialog
    {
        private const string DialogTitle = "Save As";
        private const string DialogFileNameEdit = "File name:";
        private const string DialogSaveButton = "Save";

        public void SaveAs(string fileFullPath)
        {
            var saveAsDialog = Windows.GetWindowWithName(DialogTitle, true);

            if (saveAsDialog == null)
            {
                throw new ElementNotAvailableException("Save As dialog is not found.");
            }

            var saveAsDialogFileName = NativeFinder.Find(saveAsDialog, DialogFileNameEdit, ControlType.Edit);
            var saveAsDialogSaveButton = NativeFinder.Find(saveAsDialog, DialogSaveButton, ControlType.Button);

            saveAsDialog.SetFocus();
            Wait();
            saveAsDialogFileName.SetFocus();
            Wait();
            SendKeys.SendWait("^a");
            SendKeys.SendWait(fileFullPath);
            saveAsDialogSaveButton.SetFocus();
            Wait();
            SendKeys.SendWait("{ENTER}");
        }

        private static void Wait() => Thread.Sleep(100);
    }
}
