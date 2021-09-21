using PassportSelector = UITests.PageModel.Selectors.Passport;
using static UITests.TestHelpers;


namespace UITests.PageModel.Passport
{
    public class PassportHeaderAdjustmentListItem
    {
        private readonly IAppInstance _app;

        public decimal NetAmount => GetHeaderItemNetTotalFromPassport();

        public PassportHeaderAdjustmentListItem(IAppInstance app)
        {
            _app = app;
        }   

        public decimal GetHeaderItemNetTotalFromPassport()
        {
            var total = _app.Driver.FindElement(PassportSelector.HeaderItemNetTotalPassport).Text;
            return GetNumeral(total);
        }
    }
}