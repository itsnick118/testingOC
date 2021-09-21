using OpenQA.Selenium;
using System.Collections.Generic;
using UITests.PageModel.Selectors;
using UITests.PageModel.Shared;
using static UITests.TestHelpers;

namespace UITests.PageModel
{
    public class InvoiceSummaryPage
    {
        private readonly IAppInstance _app;
        public IDialog ApproveInvoiceDialog { get; }
        public IDialog RejectInvoiceDialog { get; }
        public IDialog AdjustLineItemDialog { get; }
        public IDialog HeaderAdjustmentDialog { get; }
        public IDialog RejectLineItemDialog { get; }
        public Group SummaryPanel { get; }
        public Group InvoiceTotalsPanel { get; }
        public ItemList ItemList { get; }
        public EntityTabs Tabs { get; }
        public Dialog Dialog { get; }
        public QuickSearch QuickSearch { get; }
        public SingleEntityDropPoint DropPoint { get; }
        public EntityTabs EntityTabs { get; }

        public decimal InvoiceTotalValue => GetNumeral(_app.Driver.FindElement(Oc.InvoiceValueByClass("total")).Text);
        public decimal InvoiceNetTotalValue => GetNumeral(_app.Driver.FindElement(Oc.InvoiceValueByClass("netTotal")).Text);

        public InvoiceSummaryPage(IAppInstance app)
        {
            _app = app;
            SummaryPanel = new Group(_app);
            InvoiceTotalsPanel = new Group(_app, 1);
            ItemList = new ItemList(_app);
            Tabs = new EntityTabs(app);
            ApproveInvoiceDialog = new Dialog(app, null, Configurations.GA.Dialogs.ApproveInvoiceDialogControls(app));
            RejectInvoiceDialog = new Dialog(app, null, Configurations.GA.Dialogs.RejectInvoiceDialogControls(app));
            AdjustLineItemDialog = new Dialog(app, null, Configurations.GA.Dialogs.AdjustLineItemDialogControls(app));
            HeaderAdjustmentDialog = new Dialog(app, null, Configurations.GA.Dialogs.HeaderAdjustmentDialogControls(app));
            RejectLineItemDialog = new Dialog(app, null, Configurations.GA.Dialogs.RejectLineItemDialogControls(app));
            Dialog = new Dialog(app);
            QuickSearch = new QuickSearch(app);
            DropPoint = new SingleEntityDropPoint(app);
            EntityTabs = new EntityTabs(app);
        }

        public IReadOnlyCollection<IWebElement> GetInvoiceSummaryInfo()
        {
            _app.WaitForLoadComplete();
            var invoiceSummaryInfo = _app.Driver.FindElements(Oc.TextInfo);
            return invoiceSummaryInfo;
        }

        public void Approve() =>_app.JustClick(Oc.InvoiceItemApproveButton);

        public void Reject() =>_app.JustClick(Oc.InvoiceItemRejectButton);

        public void AccessInvoice() => _app.WaitAndClick(Oc.AccessButton);

        public void OpenMoreOptions() => _app.WaitAndClick(Oc.ListOptions);

        public void QuickFile(bool wait = true)
        {
            _app.JustClick(Oc.ItemDetailQuickFile);
            if (wait)
            {
                _app.WaitUntilElementAppears(Oc.UploadDocumentButton);
            }
        }

        public void CollapseInvoiceSummary() => _app.JustClick(Oc.PanelByIndex(2));
    }
}
