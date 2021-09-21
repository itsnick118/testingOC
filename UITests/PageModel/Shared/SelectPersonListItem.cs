using OpenQA.Selenium;

namespace UITests.PageModel.Shared
{
    public class SelectPersonListItem : ListItem
    {
        public SelectPersonListItem(IAppInstance app, IWebElement element) : base(app, element)
        {
            Name = PrimaryText;
        }

        public string Email => SecondaryText;
        public string Name { get; }
        public string Role => TertiaryText;

        public new void Select() => base.Select();

        public override string ToString()
        {
            return $"{nameof(Name)}:{Name},{nameof(Email)}:{Email},{nameof(Role)}:{Role}";
        }
    }
}
