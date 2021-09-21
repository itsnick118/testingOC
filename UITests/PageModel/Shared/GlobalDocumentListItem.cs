using OpenQA.Selenium;
using System;
using UITests.PageModel.Selectors;
using static UITests.TestHelpers;

namespace UITests.PageModel.Shared
{
    public class GlobalDocumentListItem : BaseDocumentListItem
    {
        public string Name { get; }
        public string AssociatedEntityName { get; }
        public string CreatedByFullName { get; }
        public string DocumentSize { get; }
        public DateTime UpdatedAt { get; }
        public string Status { get; }

        public GlobalDocumentListItem(IAppInstance app, IWebElement element) : base(app, element)
        {
            Name = PrimaryText;
            AssociatedEntityName = SecondaryText;
            CreatedByFullName = TertiaryText;
            DocumentSize = Meta2;
            UpdatedAt = ParseDateTime(Meta3);
            Status = Element.FindElement(Oc.ItemOptions).Text;

            FileOptions = new FileOptions(app, element);
        }

        public override string ToString()
        {
            return $"{nameof(Name)}:{Name}";
        }
    }
}
