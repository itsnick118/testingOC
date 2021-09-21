using System.Drawing;
using OpenQA.Selenium;
using UITests.PageModel.Selectors;

namespace UITests.PageModel.Shared
{
    public class PeopleListItem : ListItem
    {
        public PeopleListItem(IAppInstance app, IWebElement element) : base(app, element)
        {
            PersonName = PrimaryText;
            Role = SecondaryText;
            PersonType = TertiaryText;
        }

        public string PersonName { get; }
        public string PersonType { get; }
        private IWebElement RemoveButton => Element.FindElement(Oc.RemoveButton);
        public string Role { get; }

        public new void Edit() => base.Edit();

        public string GetEmailAddress() => App.GetToolTip(Element.FindElement(Oc.EmailIcon)).Text.Replace('"', ' ');

        public Color GetRemoveButtonColor() => HoverAndGetColor(RemoveButton);

        public bool IsContactIconVisible() => Element.FindElement(Oc.ContactIcon).Displayed;

        public bool IsEmailIconVisible() => Element.FindElement(Oc.EmailIcon).Displayed;

        public bool IsRemovePersonButtonVisible() => RemoveButton.Displayed;

        public Color GetPersonNameColor() => HoverAndGetColor(PrimaryElement);

        public void OpenEmailWindow() => Element.FindElement(Oc.EmailIcon).Click();

        public IDialog Remove() => base.Delete(Oc.RemoveButton);

        public override string ToString()
        {
            return $"{nameof(PersonName)}:{PersonName}";
        }

        public void ViewContact() => App.ClickAndWait(Oc.ContactIcon, Element);
    }
}
