using System.Collections.Generic;

namespace APITests.Passport.Json.Configuration.Model
{
    public class ModuleDefinition
    {
        public string FavoritesPage { get; set; }
        public Dictionary<ModulePage, Page> ItemPages { get; set; }
        public int Order { get; set; }
        public ConnectionStateAssignment Page { get; set; }
        public Dictionary<ModulePage, Page> Pages { get; set; }
        public SavedSearchFormConfig SavedSearchFormConfig { get; set; }
        public Dictionary<SearchListPage, Page> SearchListPages { get; set; }
        public Service[] Services { get; set; }
        public ConnectionStateAssignment State { get; set; }
        public Tab[] Tabs { get; set; }
        public string Title { get; set; }
        public string TitleClass { get; set; }
    }
}
