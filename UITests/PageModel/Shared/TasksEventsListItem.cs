using OpenQA.Selenium;
using UITests.PageModel.Selectors;

namespace UITests.PageModel.Shared
{
    public class TasksEventsListItem : ListItem
    {
        public string Name { get; }
        public string Type { get; }

        public TasksEventsListItem(IAppInstance app, IWebElement element) : base(app, element)
        {
            Name = SecondaryText;
            Type = PrimaryText;
        }

        public IDialog Delete() => base.Delete(Oc.DeleteButton);

        public new void Edit() => base.Edit();

        public override string ToString()
        {
            return $"{nameof(Name)}:{Name},{nameof(Type)}:{Type}";
        }
    }
}
