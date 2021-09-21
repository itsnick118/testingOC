using UITests.PageModel.Shared;
using UITests.PageModel.Selectors;

namespace UITests.PageModel
{
    public class InvoicesListPage
    {
        private readonly IAppInstance _app;
        public ItemList ItemList { get; }
        public IDialog ApproveInvoiceDialog { get; }
        public IDialog RejectInvoiceDialog { get; }
        public IDialog InvoiceListFilterDialog { get; }
        public ISortDialog InvoiceSortDialog { get; }
        public EntityTabs Tabs { get; }

        public InvoicesListPage(IAppInstance app)
        {
            _app = app;
            ItemList = new ItemList(app);
            Tabs = new EntityTabs(app);
            ApproveInvoiceDialog = new Dialog(app, null, Configurations.GA.Dialogs.ApproveInvoiceDialogControls(app));
            RejectInvoiceDialog = new Dialog(app, null, Configurations.GA.Dialogs.RejectInvoiceDialogControls(app));
            InvoiceListFilterDialog = new Dialog(_app, null, Configurations.GA.Dialogs.InvoiceListFilterDialogControls(_app));
            InvoiceSortDialog = new SortDialog(app);
        }

        public void Open() => _app.WaitAndClickThenWait(Oc.SpendTab);
    }
}
