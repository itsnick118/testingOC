using System.Threading;
using PassportSelector = UITests.PageModel.Selectors.Passport;

namespace UITests.PageModel.Passport
{
    public class InvoicePassportPage : PassportPage
    {
        public PassportHeaderItemList PassportHeaderItemList { get; }

        public InvoicePassportPage(IAppInstance app) : base(app)
        {
            PassportHeaderItemList = new PassportHeaderItemList(app);
        }

        public string GetNetTotal()
        {
            SwitchToInvoiceTotalsFrame();
            return App.Driver.FindElement(PassportSelector.InvoiceNetTotal).Text;
        }

        private void SwitchToInvoiceTotalsFrame()
        {
            App.SwitchToLastDriverWindow();
            SwitchToPassportFrame();
            App.Driver.SwitchTo().Frame(App.Driver.FindElement(PassportSelector.InvoiceSummaryFrame));

            //TODO : Remove Sleep and add an explicit wait
            /*Wait for html elements to load inside iframe does not seem to work
             eg. WaitUntilElementAppears(Passport.InvoiceNetTotal)*/
            Thread.Sleep(1000);
        }

        public void NavigateToHeaderAdjustmentTab()
        {
            App.SwitchToLastDriverWindow();
            SwitchToPassportFrame();

            var headerAdjustment = App.Driver.FindElement(PassportSelector.HeaderAdjustmentTab);
            headerAdjustment.Click();
            App.WaitUntilElementAppears(PassportSelector.ListItems);
        }
    }
}