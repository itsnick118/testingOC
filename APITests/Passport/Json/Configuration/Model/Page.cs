namespace APITests.Passport.Json.Configuration.Model
{
    public class Page
    {
        public Action[] Actions { get; set; }
        public CalculatedColumns CalculatedColumns { get; set; }
        public string[] Columns { get; set; }
        public Field[] ConditionalFields { get; set; }
        public Constructor[] Constructor { get; set; }
        public ServiceMethod Drag { get; set; }
        public ServiceMethod DropFile { get; set; }
        public ServiceMethod DropOutlookItem { get; set; }
        public bool DynamicDataFilter { get; set; }
        public string Entity { get; set; }
        public Field[] Fields { get; set; }
        public FilterDefinition FilterBy { get; set; }
        public string FolderPageRef { get; set; }
        public string HeaderField { get; set; }
        public string[] HiddenFields { get; set; }
        public bool IsAvailableForCmis { get; set; }
        public ItemPageOptions ItemPage { get; set; }
        public string ListContentType { get; set; }
        public ListOptions ListOptions { get; set; }
        public string ListPage { get; set; }
        public LiveUpdateDefinition LiveUpdate { get; set; }
        public string Loader { get; set; }
        public string LongColumn { get; set; }
        public string LookupType { get; set; }
        public string Mode { get; set; }
        public string[] Names { get; set; }
        public string Navigate { get; set; }
        public string PageName { get; set; }
        public PanelOptions PanelOptions { get; set; }
        public string ParentFieldName { get; set; }
        public bool RowsSelectable { get; set; }
        public string SavedSearchItemPage { get; set; }
        public string ScreenDisplayName { get; set; }
        public string[] SelectedItemDisplayColumns { get; set; }
        public ServiceMethod Show { get; set; }
        public bool ShowSavedSearches { get; set; }
        public TableOptions TableOptions { get; set; }
        public Tab[] Tabs { get; set; }
        public ServiceMethod TabSelection { get; set; }
        public string Title { get; set; }
        public string Type { get; set; }
        public bool? UseJsonApi { get; set; }
    }
}
