using System.Net;
using System.Net.Http;
using MockPassport.Mappings.Strings;
using WireMock.Matchers;
using WireMock.RequestBuilders;
using WireMock.ResponseBuilders;
using WireMock.Server;

namespace MockPassport.Mappings.Lists
{
    public class Invoices : IMapping, IUpdatable
    {
        private const string InvoiceListFile = "invoice_list_screen";
        private const string PagesUpToTen = @"searchInput.pageInfo.currentPageNumber=\d\D";

        public FluentMockServer Setup(FluentMockServer server, IEnvironment environment)
        {
            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.FetchList))
                    .WithBody(b => b.Contains(ScreenName.AsParam(ScreenName.InvoicesList)))
                    .WithBody(b => b.Contains("search-keywords=&"))
                    .WithBody(new RegexMatcher(PagesUpToTen)))
                .WithTitle("All invoices list page")
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, InvoiceListFile))
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeaders(FromFile.GetFirstNinePageHeaders(environment, InvoiceListFile)));

            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.FetchList))
                    .WithBody(b => b.Contains(ScreenName.AsParam(ScreenName.InvoicesList)))
                    .WithBody(b => b.Contains("search-keywords=&"))
                    .WithBody(b => b.Contains("searchInput.pageInfo.currentPageNumber=10")))
                .WithTitle("All invoices list page")
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, InvoiceListFile))
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeaders(FromFile.GetTenthPageHeaders(environment, InvoiceListFile)));

            return server;
        }

        public void Update(HttpClient client, IEnvironment environment, EntityIdMap entityIdMap)
        {
            new MappingUpdateRequest{ 
                Endpoint = Endpoint.FetchList,
                Content = ScreenName.AsParam(ScreenName.InvoicesList) + "&search-keywords=&searchInput.pageInfo.currentPageNumber=1" +
                          "&searchInput.pageInfo.currentPageSize=50&pageType=normal&cssClasses=&loadImmediately=false&nocache=true" +
                          "&documentTitle=abc&falseParm=0&pageOffset=0",
                Method = HttpMethod.Post,
                FileName = InvoiceListFile,
                Title = "All Invoices List Page",
                ContentType = ContentType.FormUrlEncoded,
                Environment = environment
            }.UpdateFile(client);
        }
    }
}
