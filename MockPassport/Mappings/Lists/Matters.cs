using System.Net;
using System.Net.Http;
using MockPassport.Mappings.Strings;
using WireMock.Matchers;
using WireMock.RequestBuilders;
using WireMock.ResponseBuilders;
using WireMock.Server;

namespace MockPassport.Mappings.Lists
{
    public class Matters : IMapping, IUpdatable
    {
        private const string MatterListFile = "matter_list_screen";
        private const string SpecificMattersFile = "matter_specific_matters";
        private const string PagesUpToTen = @"searchInput.pageInfo.currentPageNumber=\d\D";

        public FluentMockServer Setup(FluentMockServer server, IEnvironment environment)
        {
            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.FetchList))
                    .WithBody(b => b.Contains(ScreenName.AsParam(ScreenName.MatterList)))
                    .WithBody(b => b.Contains("search-keywords=&"))
                    .WithBody(new RegexMatcher(PagesUpToTen)))
                .WithTitle("All matters list page, first nine pages")
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, MatterListFile))
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeaders(FromFile.GetFirstNinePageHeaders(environment, MatterListFile)));

            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.FetchList))
                    .WithBody(b => b.Contains(ScreenName.AsParam(ScreenName.MatterList)))
                    .WithBody(b => b.Contains("search-keywords=&"))
                    .WithBody(b => b.Contains("searchInput.pageInfo.currentPageNumber=10")))
                .WithTitle("All matters list page, last page")
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, MatterListFile))
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeaders(FromFile.GetTenthPageHeaders(environment, MatterListFile)));

            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.FetchList))
                    .WithBody(new RegexMatcher(@"includedIDs=\d")))
                .WithTitle("All matters list page, specific matters")
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, SpecificMattersFile))
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeaders(FromFile.GetHeaders(environment, SpecificMattersFile)));

            return server;
        }

        public void Update(HttpClient client, IEnvironment environment, EntityIdMap entityIdMap)
        {
            new MappingUpdateRequest{ 
                Endpoint = Endpoint.FetchList,
                Content = ScreenName.AsParam(ScreenName.MatterList) + "&search-keywords=&searchInput.pageInfo.currentPageNumber=1" +
                          "&searchInput.pageInfo.currentPageSize=50&pageType=normal&cssClasses=&loadImmediately=false" +
                          "&nocache=true&documentTitle=abc&falseParm=0&pageOffset=0",
                Method = HttpMethod.Post,
                FileName = MatterListFile,
                Title = "All Matters List Page",
                ContentType = ContentType.FormUrlEncoded,
                Environment = environment
            }.UpdateFile(client);
        }
    }
}
