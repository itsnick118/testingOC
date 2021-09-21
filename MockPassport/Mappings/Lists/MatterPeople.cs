using System.Net;
using System.Net.Http;
using MockPassport.Mappings.Strings;
using WireMock.RequestBuilders;
using WireMock.ResponseBuilders;
using WireMock.Server;

namespace MockPassport.Mappings.Lists
{
    public class MatterPeople : IMapping, IUpdatable
    {
        private const string MatterPeopleFile = "matter_person_screen";

        public FluentMockServer Setup(FluentMockServer server, IEnvironment environment)
        {
            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.FetchList))
                    .WithBody(b => b.Contains(ScreenName.AsParam(ScreenName.MatterPersonList))))
                .WithTitle("Matter people page")
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, MatterPeopleFile))
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeaders(FromFile.GetHeaders(environment, MatterPeopleFile)));

            return server;
        }

        public void Update(HttpClient client, IEnvironment environment, EntityIdMap entityIdMap)
        {
            var map = entityIdMap.Map;

            var allMattersListPageRequest = new MappingUpdateRequest{ 
                Endpoint = Endpoint.FetchList,
                Content = ScreenName.AsParam(ScreenName.MatterPersonList) + "&search-keywords=&parentInstanceId=" + environment.ModelMatter +
                          "&parentFieldName=people&parentEntityId=" + map[EntityName.Matter] + "&searchInput.pageInfo.currentPageNumber=1" +
                          "&searchInput.pageInfo.currentPageSize=50&pageType=normal&cssClasses=" +
                          "&loadImmediately=false&nocache=true&documentTitle=abc&falseParm=0&pageOffset=0",
                Method = HttpMethod.Post,
                FileName = MatterPeopleFile,
                Title = "Matter people page",
                ContentType = ContentType.FormUrlEncoded,
                Environment = environment
            };

            allMattersListPageRequest.UpdateFile(client);
        }
    }
}
