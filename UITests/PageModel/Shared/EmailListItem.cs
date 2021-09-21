using System;
using System.Drawing;
using OpenQA.Selenium;
using UITests.PageModel.Selectors;
using static UITests.TestHelpers;

namespace UITests.PageModel.Shared
{
    public class EmailListItem : ListItem
    {
        public bool IsFolder { get; }
        public string FolderName { get; }
        public bool HasRenameButton { get; }
        public bool HasQuickFileButton { get; }

        public string From { get; }
        public string Subject { get; }
        public string EmailBody { get; }
        public DateTime ReceivedTime { get; }
        public string HasAttachment { get; }
        public IWebElement DropPoint { get; }
        public bool HasCheckBox => App.IsElementDisplayed(Oc.CheckBox, Element);

        public EmailListItem(IAppInstance app, IWebElement element) : base(app, element)
        {
            DropPoint = element;
            IsFolder = IsElementIsFolder();

            if (IsFolder)
            {
                FolderName = PrimaryText;
                HasRenameButton = app.IsElementDisplayed(Oc.RenameButton, element);
                HasQuickFileButton = app.IsElementDisplayed(Oc.ListItemQuickFile, element);
                DropPoint = element;
            }
            else
            {
                From = IsFolder ? null : PrimaryText;
                Subject = IsFolder ? null : SecondaryText;
                EmailBody = IsFolder ? null : TertiaryText;
                ReceivedTime = ParseDateTime(Meta3);
                HasAttachment = IsEmailHasAttachment();
            }
        }

        public new void Select() => base.Select();

        public IDialog Delete() => base.Delete(Oc.DeleteButton);

        public Color GetCheckBoxColor() => App.GetCheckboxColor(Element.FindElement(Oc.CheckBoxBackground));

        public new void QuickFile()
        {
            if (IsFolder)
            {
                base.QuickFile();
            }
            else
            {
                throw new NotSupportedException();
            }
        }

        public override string ToString()
        {
            return IsFolder
                ? $"{nameof(IsFolder)}:{IsFolder}, {nameof(FolderName)}:{FolderName}"
                : $"{nameof(IsFolder)}:{IsFolder}, {nameof(From)}:{From}, {nameof(ReceivedTime)}:{ReceivedTime}";
        }

        private bool IsElementIsFolder()
        {
            App.SetShortImplicitWait();

            var isFolder = !App.IsElementDisplayed(Oc.CheckBox, Element);

            App.SetLongImplicitWait();

            return isFolder;
        }

        private string IsEmailHasAttachment()
        {
            App.SetShortImplicitWait();

            var isEmailHasAttachment = App.IsElementDisplayed(Oc.EmailAttachment, Element);

            App.SetLongImplicitWait();

            return isEmailHasAttachment ? "Yes" : "No";
        }
    }
}
