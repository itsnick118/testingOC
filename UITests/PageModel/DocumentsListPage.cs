using UITests.PageModel.Shared;

namespace UITests.PageModel
{
    public class DocumentsListPage
    {
        public IDialog AddDocumentDialog { get; }
        public IDialog AddFolderDialog { get; }
        public IDialog RenameDocumentDialog { get; }
        public IDialog CheckInDocumentDialog { get; }
        public IDialog SaveCurrentViewDialog { get; }
        public IDialog MatterDocumentsListFilterDialog { get; }
        public IDialog InvoiceDocumentListFilterDialog { get; }
        public ISortDialog DocumentSortDialog { get; }
        public ItemList ItemList { get; }
        public BreadcrumbsControl BreadcrumbsControl { get; }
        public SingleEntityDropPoint DropPoint { get; }
        public QuickSearch QuickSearch { get; }

        public DocumentsListPage(IAppInstance app)
        {
            DocumentSortDialog = new SortDialog(app);
            ItemList = new ItemList(app);
            BreadcrumbsControl = new BreadcrumbsControl(app);
            DropPoint = new SingleEntityDropPoint(app);

            SaveCurrentViewDialog = new Dialog(app, null, Configurations.GA.Dialogs.SaveCurrentViewDialogControls(app));
            MatterDocumentsListFilterDialog = new Dialog(app, null, Configurations.GA.Dialogs.MatterDocumentsListFilterDialogControls(app));
            InvoiceDocumentListFilterDialog = new Dialog(app, null, Configurations.GA.Dialogs.InvoiceDocumentsListFilterDialogControls(app));
            RenameDocumentDialog = new Dialog(app, null, Configurations.GA.Dialogs.RenameDocumentDialogControls(app));
            QuickSearch = new QuickSearch(app);

            switch (app.Environment.Configuration)
            {
                default:
                    AddFolderDialog = new Dialog(app, null, Configurations.GA.Dialogs.AddFolderDialogControls(app));
                    AddDocumentDialog = new Dialog(app, null, Configurations.GA.Dialogs.AddDocumentDialogControls(app));
                    CheckInDocumentDialog = new Dialog(app, null, Configurations.GA.Dialogs.CheckInDialogControls(app));
                    break;
                case EnvironmentConfiguration.EY:
                    AddFolderDialog = new Dialog(app, null, Configurations.EY.Dialogs.AddFolderDialogControls(app));
                    AddDocumentDialog = new Dialog(app, null, Configurations.EY.Dialogs.AddDocumentDialogControls(app));
                    CheckInDocumentDialog = new Dialog(app, null, Configurations.EY.Dialogs.CheckInDialogControls(app));
                    break;
            }
        }
    }
}
