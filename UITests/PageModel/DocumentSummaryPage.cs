using OpenQA.Selenium;
using System.Collections.Generic;
using UITests.PageModel.Shared;

namespace UITests.PageModel
{
    public class DocumentSummaryPage
    {
        private readonly IAppInstance _app;

        public IDialog AddDocumentDialog { get; }
        public IDialog DeleteDocumentDialog { get; }
        public IDialog CheckInDocumentDialog { get; }
        public SingleEntityDropPoint DropPoint { get; }
        public ItemList ItemList { get; }
        public Group SummaryPanel { get; }

        public DocumentSummaryPage(IAppInstance app)
        {
            _app = app;
            DropPoint = new SingleEntityDropPoint(_app);
            ItemList = new ItemList(app);
            SummaryPanel = new Group(_app);

            switch (app.Environment.Configuration)
            {
                default:
                    AddDocumentDialog = new Dialog(_app, null, Configurations.GA.Dialogs.AddDocumentDialogControls(_app));
                    CheckInDocumentDialog = new Dialog(_app, null, Configurations.GA.Dialogs.CheckInDialogControls(_app));
                    DeleteDocumentDialog = new Dialog(_app);
                    break;
                case EnvironmentConfiguration.EY:
                    AddDocumentDialog = new Dialog(_app, null, Configurations.EY.Dialogs.AddDocumentDialogControls(_app));
                    CheckInDocumentDialog = new Dialog(_app, null, Configurations.EY.Dialogs.CheckInDialogControls(_app));
                    break;
            }
        }

        public IReadOnlyCollection<IWebElement> GetDocumentSummaryInfo()
        {
            _app.WaitForLoadComplete();
            var documentSummaryInfo = _app.Driver.FindElements(Selectors.Oc.TextInfo);
            return documentSummaryInfo;
        }

        public void NavigateToParentMatter() => _app.ClickAndWait(Selectors.Oc.ByClassName("summary-breadcrumbs-parent-label"));

        public void NavigateToParent() => _app.ClickAndWait(Selectors.Oc.BreadCrumbsParent);
        public void QuickFile() => _app.ClickAndWait(Selectors.Oc.ItemDetailQuickFile);

        public void WaitForStatusChangeTo(string status)
        {
            _app.WaitFor(condition => IsStatusEqualsTo(status));
        }

        public bool IsStatusEqualsTo(string status)
        {
            status = status.ToLower();
            var documentSummaryInfo = _app.Driver.FindElements(Selectors.Oc.TextInfo);

            try
            {
                foreach (var webElement in documentSummaryInfo)
                {
                    if (webElement.Text.ToLower() == status)
                    {
                        return true;
                    }
                }
            }
            catch (StaleElementReferenceException)
            {
                // ignore
            }

            return false;
        }
    }
}
