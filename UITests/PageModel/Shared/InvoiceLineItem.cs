using OpenQA.Selenium;
using UITests.PageModel.Selectors;
using static UITests.TestHelpers;

namespace UITests.PageModel.Shared
{
    public class InvoiceLineItem : ListItem
    {
        public decimal NetTotal => GetNumeral(Meta2);

        public InvoiceLineItem(IAppInstance app, IWebElement element) : base(app, element)
        {
        }

        public new void Edit() => base.Edit();

        public void Reject()
        {
            App.JustClick(Oc.InvoiceItemRejectButton, Element);
        }
    }
}
