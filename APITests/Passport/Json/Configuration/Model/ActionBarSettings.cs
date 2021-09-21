namespace APITests.Passport.Json.Configuration.Model
{
    public class ActionBarSettings
    {
        public ServiceMethod Click { get; set; }
        public Action DeleteAction { get; set; }
        public bool HideIfMoreThanOneItemsSelected { get; set; }
        public string Icon { get; set; }
        public string IconPendingClass { get; set; }
        public string Label { get; set; }
        public ServiceMethod MenuItemClick { get; set; }
        public string MenuType { get; set; }
        public string Name { get; set; }
        public ServiceMethod OnLoad { get; set; }
        public bool ShowInMenuBar { get; set; }
        public bool ShowOnlyOnItemSelection { get; set; }
        public string TitleClass { get; set; }
        public string ToolTip { get; set; }
    }
}
