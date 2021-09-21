using OpenQA.Selenium;
using System;
using UITests.PageModel.Selectors;
using static UITests.TestHelpers;

namespace UITests.PageModel.Shared
{
    public class MatterDocumentListItem : BaseDocumentListItem
    {
        public bool IsFolder { get; }

        public string FolderName { get; }

        public string Name { get; }
        public string DocumentFileName { get; }
        public string LastModifiedBy { get; }
        public string DocumentSize { get; }
        public DateTime UpdatedAt { get; }
        public string Status { get; }
        public IWebElement DropPoint { get; }
        public MatterDocumentListItem(IAppInstance app, IWebElement element) : base(app, element)
        {
            IsFolder = IsElementIsFolder();
            DropPoint = element;

            if (IsFolder)
            {
                FolderName = PrimaryText;
            }
            else
            {
                Name = PrimaryText;
                DocumentFileName = SecondaryText;
                LastModifiedBy = TertiaryText;
                DocumentSize = Meta2;
                Status = Element.FindElement(Oc.ItemOptions).Text;
                DropPoint = element;

                if (!string.IsNullOrEmpty(Meta3))
                {
                    UpdatedAt = ParseDateTime(Meta3);
                }

                FileOptions = new FileOptions(app, element);
            }
        }

        public void Rename()
        {
            App.ClickAndWait(Oc.RenameButton, Element);
        }

        public override string ToString()
        {
            return IsFolder
                ? $"{nameof(IsFolder)}:{IsFolder},{nameof(FolderName)}:{FolderName}"
                : $"{nameof(IsFolder)}:{IsFolder},{nameof(Name)}:{Name}";
        }

        private bool IsElementIsFolder()
        {
            App.SetShortImplicitWait();

            var isFolder = !App.IsElementDisplayed(Oc.SummaryIcon, Element);

            App.SetLongImplicitWait();

            return isFolder;
        }

        public bool IsFileOfType(OfficeApp appName)
        {
            return App.Driver.FindElement(Oc.OfficeAppType(appName.ToString())).Displayed;
        }
    }
}
