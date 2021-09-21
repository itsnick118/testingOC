namespace APITests.Passport.Json.Configuration.Model
{
    public class Action
    {
        public ActionBarSettings ActionBarSettings { get; set; }
        public string ActionMenuClass { get; set; }
        public string ActionParentClass { get; set; }
        public ServiceMethod Click { get; set; }
        public string Column { get; set; }
        public string Command { get; set; }
        public CommandArgs CommandArgs { get; set; }
        public string Context { get; set; }
        public dynamic DefaultClass { get; set; }
        public string DynamicTooltipField { get; set; }
        public ServiceMethod GetToggledStatus { get; set; }
        public string Id { get; set; }
        public bool IsComponent { get; set; }
        public bool IsIndicator { get; set; }
        public dynamic Label { get; set; }
        public string MenuAboveViewportClass { get; set; }
        public MenuActionItem[] MenuActionItems { get; set; }
        public string MenuBelowViewportClass { get; set; }
        public string MenuContainerClass { get; set; }
        public string MenuInViewportClass { get; set; }
        public string Name { get; set; }
        public bool NavigateBack { get; set; }
        public bool OnlineOnly { get; set; }
        public Options Options { get; set; }
        public string PageContainerClass { get; set; }
        public string PendingClass { get; set; }
        public int RepeatColumnIndex { get; set; }
        public bool Responsive { get; set; }
        public bool ShowOnlyInActionBar { get; set; }
        public string TitleClass { get; set; }
        public string ToggledClass { get; set; }
        public string UsageEvent { get; set; }
        public dynamic Visible { get; set; }
    }
}
