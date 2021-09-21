using System.Drawing;
using System.IO;
using IntegratedDriver;
using OpenQA.Selenium;
using UITests.PageModel.Selectors;

namespace UITests.PageModel.Shared
{
    public abstract class BaseDocumentListItem : ListItem
    {
        public FileOptions FileOptions { get; protected set; }

        public string NavigateToSummaryButtonTooltip => App.GetToolTip(App.Driver.FindElement(Oc.SummaryIcon)).Text;

        protected BaseDocumentListItem(IAppInstance app, IWebElement element) : base(app, element)
        {
        }

        public FileInfo Download(string filename)
        {
            App.JustClick(Oc.DownloadButton, Element);
            var saveAsDialog = new SaveAsNativeDialog();
            var fileFullName = Path.Combine(Windows.GetWorkingTempFolder().FullName, filename);
            saveAsDialog.SaveAs(fileFullName);
            Windows.WaitUntilFileDownloaded(fileFullName);
            return new FileInfo(fileFullName);
        }

        public IDialog Delete() => base.Delete(Oc.DeleteButton);

        public bool IsNavigateToSummaryVisible() => Element.FindElement(Oc.SummaryIcon).Displayed;

        public string DownloadButtonTooltip => App.GetToolTip(App.Driver.FindElement(Oc.DownloadButton)).Text;

        public bool IsDownloadIconVisible()
        {
            return App.IsElementDisplayed(Oc.DownloadButton, Element);
        }

        public Color GetColorOnHoverOverFileName() => HoverAndGetColor(PrimaryElement);

        public string GetTextDecorationOnHoverOverFileName()
        {
            HoverMouseOverListItem(PrimaryElement);
            return PrimaryElement.GetCssValue("text-decoration");
        }
    }
}
