using OpenQA.Selenium;
using UITests.PageModel.Selectors;
using static UITests.TestHelpers;

namespace UITests.PageModel.Shared
{
    public class InvoiceHeaderItem : ListItem
    {
        public decimal NetAmount => GetNumeral(Meta3);
        public decimal AdjustmentAmount => GetNumeral(Element.FindElement(Oc.LineItemAmount).Text);

        public new void Edit() => base.Edit();

        public InvoiceHeaderItem(IAppInstance app, IWebElement element) : base(app, element)
        {
        }
    }
}
