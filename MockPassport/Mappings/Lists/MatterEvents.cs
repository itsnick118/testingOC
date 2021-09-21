using System.Net;
using System.Net.Http;
using MockPassport.Mappings.Strings;
using WireMock.RequestBuilders;
using WireMock.ResponseBuilders;
using WireMock.Server;

namespace MockPassport.Mappings.Lists
{
    public class MatterEvents : IMapping, IUpdatable
    {
        private const string MatterEventsFile = "matter_events_screen";
        public FluentMockServer Setup(FluentMockServer server, IEnvironment environment)
        {
            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.FetchList))
                    .WithBody(b => b.Contains(ScreenName.AsParam(ScreenName.MatterEventList))))
                .WithTitle("Matter events page")
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, MatterEventsFile))
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeaders(FromFile.GetHeaders(environment, MatterEventsFile)));

            return server;
        }

        public void Update(HttpClient client, IEnvironment environment, EntityIdMap entityIdMap)
        {
            var map = entityIdMap.Map;

            new MappingUpdateRequest
            { 
                Endpoint = Endpoint.FetchList,
                Content = ScreenName.AsParam(ScreenName.MatterEventList) + "&search-keywords=" +
                          "&parentInstanceId=" + environment.ModelMatter + "&parentFieldName=matterEvents" +
                          "&parentEntityId=" + map[EntityName.Matter] + "" +
                          "&searchInput.pageInfo.currentPageNumber=1&searchInput.pageInfo.currentPageSize=50" +
                          "&pageType=normal&cssClasses=&loadImmediately=false&nocache=true&documentTitle=abc" +
                          "&falseParm=0&pageOffset=0",
                Method = HttpMethod.Post,
                FileName = MatterEventsFile,
                Title = "Matter events page",
                ContentType = ContentType.FormUrlEncoded,
                Environment = environment
            }.UpdateFile(client);
        }
    }
}
