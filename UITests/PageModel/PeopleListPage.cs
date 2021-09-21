using UITests.PageModel.Shared;

namespace UITests.PageModel
{
    public class PeopleListPage
    {
        public IDialog AddPersonDialog { get; }
        public IDialog ViewPersonDialog { get; }
        public ISortDialog PeopleSortDialog { get; }
        public ItemList ItemList { get; }

        public PeopleListPage(IAppInstance app)
        {
            switch (app.Environment.Configuration)
            {
                default:
                    AddPersonDialog = new Dialog(app, null, Configurations.GA.Dialogs.AddPersonDialogControls(app));
                    break;

                case EnvironmentConfiguration.EY:
                    AddPersonDialog = new Dialog(app, null, Configurations.EY.Dialogs.AddPersonDialogControls(app));
                    break;
            }
            ViewPersonDialog = new Dialog(app);
            PeopleSortDialog = new SortDialog(app);
            ItemList = new ItemList(app);
        }

        public PeopleListItem RemoveAllPersonsExceptPIC()
        {
            for (var i = ItemList.GetCount() - 1; i >= 0; i--)
            {
                var person = ItemList.GetPeopleListItemByIndex(i);
                if (person.Role != "Primary Internal Contact")
                {
                    person.Remove().Confirm();
                }
            }

            return ItemList.GetPeopleListItemByIndex(0);
        }
    }
}
