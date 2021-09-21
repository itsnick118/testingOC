using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using IntegratedDriver;
using OpenQA.Selenium;
using UITests.PageModel.Selectors;
using UITests.PageModel.Shared.InputControls;

namespace UITests.PageModel.Shared
{
    public class Dialog : IDialog
    {
        protected IAppInstance App;

        public InputControlList Controls { get; set; }
        public IWebElement Context { get; set; }

        public string HeaderText => App.Driver.FindElement(Oc.DialogHeader).Text;
        public string Text => App.Driver.FindElement(Oc.DialogActionMessage).Text;

        public Dialog(IAppInstance app, IWebElement context = null, InputControlList controls = null)
        {
            App = app;
            Context = context;
            Controls = controls;
        }

        public bool IsDisplayed() => App.IsElementDisplayed(Oc.DialogHeader);

        public string[] GetDialogButtons()
        {
            var dialogButtons = App.Driver.FindElements(Oc.DialogButton);
            return dialogButtons.Select(btn => btn.Text).ToArray();
        }

        public void ClickRadioButton(string labelName)
        {
            App.JustClick(Oc.RadioButtonByLabel(labelName));
        }

        public void Save(bool wait = true)
        {
            App.JustClick(Oc.SaveButton);

            if (!wait)
            {
                return;
            }
            App.WaitUntilElementDisappears(Oc.DialogHeader);
            App.WaitForLoadComplete();
        }

        public void SaveAs(string fileName, bool shouldSaveInTempFolder, bool withNewName = false, bool wait = true)
        {
            if (withNewName)
            {
                App.JustClick(Oc.SaveWithANewName);
            }
            else
            {
                App.JustClick(Oc.SaveAsButton);
            }
            if (!wait)
            {
                return;
            }
            App.WaitUntilElementDisappears(Oc.DialogHeader);
            App.WaitForLoadComplete();
            var saveAsDialog = new SaveAsNativeDialog();
            if (shouldSaveInTempFolder)
            {
                var fileFullName = Path.Combine(Windows.GetWorkingTempFolder().FullName, fileName);
                saveAsDialog.SaveAs(fileFullName);
            }
            else
            {
                saveAsDialog.SaveAs(fileName);
            }
        }

        public void Confirm()
        {
            App.JustClick(Oc.DialogOkButton);
            App.WaitForListLoadComplete();
            App.WaitForLoadComplete();
        }

        public void Apply()
        {
            App.JustClick(Oc.FormApplyButton);
            App.WaitUntilElementDisappears(Oc.DialogHeader);
            App.WaitForLoadComplete();
        }

        public void Approve()
        {
            App.JustClick(Oc.DialogApproveButton);
            App.WaitUntilElementDisappears(Oc.DialogHeader);
            App.WaitForLoadComplete();
        }

        public void Reject()
        {
            App.JustClick(Oc.DialogRejectButton);
            App.WaitUntilElementDisappears(Oc.DialogHeader);
            App.WaitForLoadComplete();
        }

        public void Update() => App.ClickAndWait(Oc.UpdateButton);

        public void Cancel(bool wait = true)
        {
            App.JustClick(Oc.CancelButton);

            if (!wait)
            {
                return;
            }
            App.WaitUntilElementDisappears(Oc.CancelButton);
            App.WaitForLoadComplete();
        }

        public void UploadDocument()
        {
            App.WaitAndClickThenWait(Oc.UploadDocumentButton);
        }

        public void SaveAndUpload(string fileName, bool shouldSaveInTempFolder, bool withNewName = false, bool wait = true)
        {
            App.WaitAndClickThenWait(Oc.SaveAndUploadButton);
            if (!wait)
            {
                return;
            }
            App.WaitUntilElementDisappears(Oc.DialogHeader);
            App.WaitForLoadComplete();
            var saveAsDialog = new SaveAsNativeDialog();
            if (shouldSaveInTempFolder)
            {
                var fileFullName = Path.Combine(Windows.GetWorkingTempFolder().FullName, fileName);
                saveAsDialog.SaveAs(fileFullName);
            }
            else
            {
                saveAsDialog.SaveAs(fileName);
            }
        }

        public void Proceed() => UploadDocument();

        public void RestoreDefaults()
        {
            App.JustClick(Oc.RestoreDefaults);
            App.WaitForListLoadComplete();
        }

        public void Keep()
        {
            if (IsDisplayed())
            {
                App.JustClick(Oc.DialogKeepButton);
            }
            App.WaitForLoadComplete();
        }

        public void Overwrite()
        {
            if (IsDisplayed())
            {
                App.ClickAndWait(Oc.DialogOverwriteButton);
            }
            App.WaitForLoadComplete();
        }

        public void SelectFile(string filePath)
        {
            if (IsDisplayed())
            {
                App.JustClick(Oc.DialogSelectFileButton);
            }
            App.WaitForLoadComplete();

            var openDoucmentDialog = new OpenNativeDialog();
            openDoucmentDialog.Open(filePath);
        }

        public void Remove()
        {
            if (IsDisplayed())
            {
                App.JustClick(Oc.DialogRemoveOnDiscardButton);
            }
        }

        public void DiscardChanges()
        {
            App.JustClick(Oc.DialogDiscardChangesButton);
        }

        public void DoNotDiscard()
        {
            App.JustClick(Oc.DialogDoNotDiscardButton);
        }

        public void Reset()
        {
            App.JustClick(Oc.DialogResetButton);
            App.WaitForLoadComplete();
        }

        public void Edit()
        {
            const string EditHeader = "Edit";
            App.JustClick(Oc.DialogEditButton);

            App.WaitFor(condition => App.IsElementDisplayed(Oc.DialogHeaderByName(EditHeader)));
        }

        public List<string> GetAllLabelTexts()
        {
            var labels = App.Driver.FindElements(Oc.GetAllLabelsFromDialog);
            var labelsList = new List<string>();
            foreach (var items in labels)
            {
                var labelText = items.Text;
                var index = labelText.IndexOf("\r", StringComparison.Ordinal);
                if (index > 0)
                {
                    var labelName = labelText.Substring(0, index);
                    labelsList.Add(labelName);
                }
                else
                {
                    labelsList.Add(labelText);
                }
            }
            return labelsList;
        }
    }
}
