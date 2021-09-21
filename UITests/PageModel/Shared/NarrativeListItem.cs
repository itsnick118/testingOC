using System;
using OpenQA.Selenium;
using UITests.PageModel.Selectors;
using static UITests.TestHelpers;

namespace UITests.PageModel.Shared
{
    public class NarrativeListItem : ListItem
    {
        public string Type { get; }
        public string Description { get; }
        public string Narrative { get; }
        public string EnteredBy { get; }
        public DateTime? NarrativeDate { get; }

        public NarrativeListItem(IAppInstance app, IWebElement element) : base(app, element)
        {
            Type = PrimaryText;
            Description = SecondaryText;
            Narrative = TertiaryText;
            EnteredBy = Meta2;
            NarrativeDate = !string.IsNullOrWhiteSpace(Meta3) ? ParseDateTime(Meta3) : new DateTime?();
        }

        public IDialog Delete() => base.Delete(Oc.DeleteButton);

        public new void Edit() => base.Edit();

        public override string ToString()
        {
            return $"{nameof(Description)}:{Description}, {nameof(NarrativeDate)}:{NarrativeDate}";
        }
    }
}
