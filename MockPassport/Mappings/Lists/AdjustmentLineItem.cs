using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using MockPassport.Mappings.Strings;
using WireMock.RequestBuilders;
using WireMock.ResponseBuilders;
using WireMock.Server;

namespace MockPassport.Mappings.Lists
{
    public class AdjustmentLineItem: IMapping, IUpdatable
    {
        private const string ListScreenFile = "invoice_adjustment_line_item_list_screen";
        private const string RedlineTotalFile = "adjustment_line_item_redline_total";

        public FluentMockServer Setup(FluentMockServer server, IEnvironment environment)
        {
            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.FetchList))
                    .WithBody(b => b.Contains(ScreenName.AsParam(ScreenName.AdjustmentLineItemList))))
                .WithTitle("Invoice Adjustment Line Item list screen")
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, ListScreenFile))
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeaders(FromFile.GetHeaders(environment, ListScreenFile)));

            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.GetEntity(EntityName.AdjustmentLineItem)))
                    .WithParam(ParamKey.AttributeNames, "redLineTotal"))
                .WithTitle("Invoice Adjustment Red Line Total screen")
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, RedlineTotalFile))
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeaders(FromFile.GetHeaders(environment, RedlineTotalFile)));

            return server;
        }

        public void Update(HttpClient client, IEnvironment environment, EntityIdMap entityIdMap)
        {
            var map = entityIdMap.Map;

            new MappingUpdateRequest
            {
                Endpoint = Endpoint.FetchList,
                Content = ScreenName.AsParam(ScreenName.AdjustmentLineItemList) + "&search-keywords=" +
                          "&parentInstanceId=" + environment.ModelInvoiceHeader + "&parentFieldName=adjustmentLineItems&parentEntityId=" 
                          + map[EntityName.InvoiceHeader] + "&searchInput.pageInfo.currentPageNumber=1&searchInput.pageInfo.currentPageSize=50" +
                          "&pageType=normal&cssClasses=&loadImmediately=false&nocache=true&documentTitle=abc" +
                          "&falseParm=0&pageOffset=0",
                Method = HttpMethod.Post,
                FileName = ListScreenFile,
                Title = "Invoice Adjustment Line Item list screen",
                ContentType = ContentType.FormUrlEncoded,
                Environment = environment
            }.UpdateFile(client);

            new MappingUpdateRequest
            {
                Endpoint = Endpoint.GetEntity(EntityName.AdjustmentLineItem, 
                    map[EntityName.AdjustmentLineItem]),
                Method = HttpMethod.Get,
                FileName = RedlineTotalFile,
                Title = "Invoice Adjustment Red Line Total screen",
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
