using System.Net;
using System.Net.Http;
using MockPassport.Mappings.Strings;
using WireMock.RequestBuilders;
using WireMock.ResponseBuilders;
using WireMock.Server;

namespace MockPassport.Mappings.Lists
{
    public class EmailDocuments : IMapping, IUpdatable
    {
        private const string MatterEmailsListFile = "matter_emails_screen";

        public FluentMockServer Setup(FluentMockServer server, IEnvironment environment)
        {
            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.FetchList))
                    .WithBody(b => b.Contains(ScreenName.AsParam(ScreenName.EmailDocumentList))))
                .WithTitle("Matter Emails list page")
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, MatterEmailsListFile))
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeaders(FromFile.GetHeaders(environment, MatterEmailsListFile)));

            return server;
        }

        public void Update(HttpClient client, IEnvironment environment, EntityIdMap entityIdMap)
        {
            var map = entityIdMap.Map;

            new MappingUpdateRequest{ 
                Endpoint = Endpoint.FetchList,
                Content = ScreenName.AsParam(ScreenName.EmailDocumentList)+ "&search-keywords=&parentInstanceId=" + environment.ModelMatter + "" +
                          "&parentFieldName=emailDocuments&parentEntityId=" + map[EntityName.Matter] + "&searchInput.pageInfo.currentPageNumber=1" +
                          "&searchInput.pageInfo.currentPageSize=50&multimode=OneToMany&mode=show&" +
                          "pageType=normal&cssClasses=&loadImmediately=false&nocache=true&documentTitle=abc" +
                          "&falseParm=0&pageOffset=0",
                Method = HttpMethod.Post,
                FileName = MatterEmailsListFile,
                Title = "Matter Emails list page",
                ContentType = ContentType.FormUrlEncoded,
                Environment = environment
            }.UpdateFile(client);
        }
    }
}
