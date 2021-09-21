using System;
using OpenQA.Selenium;
using UITests.PageModel.Selectors;
using static UITests.TestHelpers;

namespace UITests.PageModel.Shared
{
    public class InvoiceListItem : ListItem
    {
        public string Number { get; }
        public string OrganizationName { get; }
        public string VendorId { get; }
        public string MatterName { get; }
        public string MatterNumber { get; }
        public DateTime ReceivedDate { get; }
        public string TotalNetAmount { get; }
        public bool HasApproveButton { get; }
        public IWebElement DropPoint { get; }

        public InvoiceListItem(IAppInstance app, IWebElement element) : base(app, element)
        {
            DropPoint = element;
            HasApproveButton = app.IsElementDisplayed(Oc.InvoiceItemApproveButton);
            Number = PrimaryText;
            ReceivedDate = ParseDateTime(Meta3);
            TotalNetAmount = Meta2;

            var secondary = SecondaryText.Split(new[] { " - " }, StringSplitOptions.None);
            var ternary = TertiaryText.Split(new[] { " - " }, StringSplitOptions.None);
            OrganizationName = secondary[0];
            VendorId = secondary[1];
            MatterName = ternary[0];
            MatterNumber = ternary[1];
        }

        public void Approve()
        {
            App.JustClick(Oc.InvoiceItemApproveButton, Element);
        }

        public void Reject()
        {
            App.JustClick(Oc.InvoiceItemRejectButton, Element);
        }

        public void AccessInvoice()
        {
            App.WaitAndClick(Oc.AccessButton, Element);
        }

        public new void QuickFile()
        {
            base.QuickFile();
        }
    }
}
