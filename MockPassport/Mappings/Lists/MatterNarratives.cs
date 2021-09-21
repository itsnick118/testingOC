using System.Net;
using System.Net.Http;
using MockPassport.Mappings.Strings;
using WireMock.RequestBuilders;
using WireMock.ResponseBuilders;
using WireMock.Server;

namespace MockPassport.Mappings.Lists
{
    public class MatterNarratives : IMapping, IUpdatable
    {
        private const string MatterNarrativesFile = "matter_narratives_screen";

        public FluentMockServer Setup(FluentMockServer server, IEnvironment environment)
        {
            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.FetchList))
                    .WithBody(b => b.Contains(ScreenName.AsParam(ScreenName.MatterNarrativesList))))
                .WithTitle("Matter narratives page")
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, MatterNarrativesFile))
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeaders(FromFile.GetHeaders(environment, MatterNarrativesFile)));

            return server;
        }

        public void Update(HttpClient client, IEnvironment environment, EntityIdMap entityIdMap)
        {
            var map = entityIdMap.Map;

            new MappingUpdateRequest
            { 
                Endpoint = Endpoint.FetchList,
                Content = ScreenName.AsParam(ScreenName.MatterNarrativesList) + "&search-keywords=&parentInstanceId=" + environment.ModelMatter +
                          "&parentFieldName=matterNarratives&parentEntityId=" + map[EntityName.Matter] + "&searchInput.pageInfo.currentPageNumber=1" +
                          "&searchInput.pageInfo.currentPageSize=50&pageType=normal&cssClasses=&loadImmediately=false" +
                          "&nocache=true&documentTitle=abc&falseParm=0&pageOffset=0",
                Method = HttpMethod.Post,
                FileName = MatterNarrativesFile,
                Title = "Matter narratives page",
                ContentType = ContentType.FormUrlEncoded,
                Environment = environment
            }.UpdateFile(client);
        }
    }
}
