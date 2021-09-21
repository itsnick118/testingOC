using System.Collections.Generic;
using OpenQA.Selenium;
using UITests.PageModel.Shared.InputControls;

namespace UITests.PageModel.Shared
{
    public interface IDialog
    {
        IWebElement Context { get; set; }
        InputControlList Controls { get; set; }
        string HeaderText { get; }
        string Text { get; }

        void Cancel(bool wait = true);

        void Confirm();

        bool IsDisplayed();

        string[] GetDialogButtons();

        void Save(bool wait = true);

        void ClickRadioButton(string labelName);

        void SaveAs(string fileName, bool shouldSaveInTempFolder, bool withNewName = false, bool wait = true);

        void Reset();

        void Apply();

        void Approve();

        void Update();

        void UploadDocument();

        void Proceed();

        void RestoreDefaults();

        void Keep();

        void Overwrite();

        void SelectFile(string filePath);

        void Edit();

        void Remove();

        void Reject();

        void DiscardChanges();

        void DoNotDiscard();

        List<string> GetAllLabelTexts();
    }
}
