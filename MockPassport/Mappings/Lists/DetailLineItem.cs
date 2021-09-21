using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using MockPassport.Mappings.Strings;
using WireMock.RequestBuilders;
using WireMock.ResponseBuilders;
using WireMock.Server;

namespace MockPassport.Mappings.Lists
{
    public class DetailLineItem: IMapping, IUpdatable
    {
        private const string DetailListScreenFile = "invoice_detail_line_item_list_screen";
        private const string DetailRedlineFile = "detail_line_item_redline_total";

        public FluentMockServer Setup(FluentMockServer server, IEnvironment environment)
        {
            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.FetchList))
                    .WithBody(b => b.Contains(ScreenName.AsParam(ScreenName.DetailLineItemList))))
                .WithTitle("Invoice Detail Line Item list screen")
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, DetailListScreenFile))
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeader(HeaderKey.ContentType, ContentType.TextHtml)
                    .WithHeaders(FromFile.GetHeaders(environment, DetailListScreenFile)));

            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.GetEntity(EntityName.DetailLineItem)))
                    .WithParam(ParamKey.AttributeNames, "redLineTotal"))
                .WithTitle("Invoice Detail Red Line Total screen")
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, DetailRedlineFile))
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeaders(FromFile.GetHeaders(environment, DetailRedlineFile)));

            return server;
        }

        public void Update(HttpClient client, IEnvironment environment, EntityIdMap entityIdMap)
        {
            var map = entityIdMap.Map;

            new MappingUpdateRequest
            {
                Endpoint = Endpoint.FetchList,
                Content = ScreenName.AsParam(ScreenName.DetailLineItemList) + "&search-keywords=" +
                          "&parentInstanceId="+ environment.ModelInvoiceHeader +"&parentFieldName=detailLineItems&parentEntityId=" + 
                          map[EntityName.InvoiceHeader] + "&searchInput.pageInfo.currentPageNumber=1&searchInput.pageInfo.currentPageSize=50" +
                          "&pageType=normal&cssClasses=&loadImmediately=false&nocache=true&documentTitle=abc" +
                          "&falseParm=0&pageOffset=0",
                Method = HttpMethod.Post,
                FileName = DetailListScreenFile,
                Title = "Invoice Detail Line Item list screen",
                ContentType = ContentType.FormUrlEncoded,
                Environment = environment
            }.UpdateFile(client);

            new MappingUpdateRequest
            {
                Endpoint = Endpoint.GetEntity(EntityName.AdjustmentLineItem,
                    map[EntityName.DetailLineItem]),
                Method = HttpMethod.Get,
                FileName = DetailRedlineFile,
                Title = "Invoice Detail Red Line Total screen",
                Environment = environment,
                Parameters = new Dictionary<string, string>
                {
                    {
                        ParamKey.AttributeNames, "redLineTotal"
                    }
                }
            }.UpdateFile(client);
        }
    }
}
