// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;
using IntegratedDriver;
using IntegratedDriver.ElementFinders;

namespace UITests.PageModel
{
    public class OpenNativeDialog
    {
        private const string DialogFileNameEdit = "File name:";
        private const string DialogOpenButton = "Open";
        private const string DialogTitle = "Open";

        public void Open(string filePath)
        {
            var openDocumentDialog = Windows.GetWindowWithName(DialogTitle, true);
            if (openDocumentDialog == null)
            {
                throw new ElementNotAvailableException("Open dialog is not found.");
            }

            var openDialogFileName = NativeFinder.Find(openDocumentDialog, DialogFileNameEdit, ControlType.Edit);
            var openDialogOpenButton = NativeFinder.Find(openDocumentDialog, DialogOpenButton, ControlType.Button);
            openDocumentDialog.SetFocus();
            Wait();

            openDialogFileName.SetFocus();
            Wait();

            SendKeys.SendWait("^a");
            SendKeys.SendWait(filePath);
            openDialogOpenButton.SetFocus();
            Wait();

            SendKeys.SendWait(Keys.Space.ToString());
            UserInput.LeftClick(openDialogOpenButton);
        }

        private static void Wait() => Thread.Sleep(100);
    }
}
