using System.Collections;
using NUnit.Framework;

namespace APITests.Passport.ScreenScraping
{
    [TestFixture]
    public class MatterApiTests
    {
        private PassportClient _passportClient;

        [SetUp]
        public void SetUp()
        {
            _passportClient = new PassportClient(new EnvironmentConfiguration(Environment.PASSPORT_2_5));
        }

        [Test]
        public void GetMatters_RetrievesMatters()
        {
            var query = new ScreenQuery
            {
                ScreenName = "Matter List - Office",
                SearchKeywords = "",
                CurrentPageNumber = 1,
                CurrentPageSize = 50,
                PageType = "normal",
                CssClasses = "",
                LoadImmediately = false,
                NoCache = true,
                FalseParm = 0,
                PageOffset = 0
            };

            var screen = _passportClient.GetScreen(query, true);

            var tableHeader = screen.GetTableHeader() as ICollection;

            Assert.IsNotNull(tableHeader);
            Assert.Contains("Matter Name", tableHeader);
            Assert.Contains("Matter Number", tableHeader);
        }
    }
}
