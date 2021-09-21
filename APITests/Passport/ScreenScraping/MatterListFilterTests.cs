using System.Collections;
using System.Collections.Generic;
using NUnit.Framework;

namespace APITests.Passport.ScreenScraping
{
    [TestFixture]
    public class MatterListFilterTests
    {
        private PassportClient _passportClient;
        private ScreenQuery _baseQuery;
        
        [SetUp]
        public void SetUp()
        {
            _passportClient = new PassportClient(new EnvironmentConfiguration(Environment.PASSPORT_2_5));
            _baseQuery = new ScreenQuery
            {
                ScreenName = "Matter List - Office",
                SearchKeywords = "",
                CurrentPageNumber = 1,
                CurrentPageSize = 50,
                PageType = "normal",
                CssClasses = string.Empty,
                LoadImmediately = false,
                NoCache = true,
                DocumentTitle = "abc",
                FalseParm = 0,
                PageOffset = 0
            };
        }

        [Test]
        public void MatterListFilter_BudgetReview()
        {
            var testQuery = _baseQuery.Clone() as ScreenQuery;
            if (testQuery != null)
            {
                testQuery.DynamicSearch = new DynamicSearch
                {
                    SearchCriteria = new List<DynamicSearchCriterion>
                    {
                        new DynamicSearchCriterion("GREATER_THAN", "0", new SearchAttribute(3728)),
                        new DynamicSearchCriterion("EQUALS", "1", new SearchAttribute(3728))
                    }
                };
            }

            var screen = _passportClient.GetScreen(testQuery, false);

            var tableHeader = screen.GetTableHeader() as ICollection;

            Assert.IsNotNull(tableHeader);
            Assert.Contains("Matter Name", tableHeader);
            Assert.Contains("Matter Number", tableHeader);
        }
    }
}