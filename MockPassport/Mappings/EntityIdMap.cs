using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Xml;
using Microsoft.AspNetCore.WebUtilities;
using MockPassport.Mappings.Strings;
// using HttpMethod = Microsoft.AspNetCore.Server.Kestrel.Core.Internal.Http.HttpMethod;

namespace MockPassport.Mappings
{
    public class EntityIdMap
    {
        private readonly IList<string> _entityList = new List<string>
        {
            EntityName.AdjustmentLineItem,
            EntityName.DetailLineItem,
            EntityName.EmailDocument,
            EntityName.InvoiceHeader,
            EntityName.Matter
        };

        public IDictionary<string, int> Map { get; set; }

        public EntityIdMap(HttpClient client)
        {
            Map = new Dictionary<string, int>();

            foreach (var entity in _entityList)
            {
                var uri = new Uri(
                    QueryHelpers.AddQueryString(
                        client.BaseAddress + Endpoint.GetEntity(EntityName.MetaData),
                        "searchParameters", $"((name)(EQUALS)({entity}))"));

                var request = new HttpRequestMessage
                {
                    RequestUri = uri,
                    Content = null,
                    Method = HttpMethod.Get
                };

                var response = client.SendAsync(request);
                if (response.Result.StatusCode != HttpStatusCode.OK) continue;

                using (var content = response.Result.Content)
                {
                    var parsedResult = new XmlDocument(); 
                    parsedResult.LoadXml(content.ReadAsStringAsync().Result);

                    if (parsedResult.DocumentElement == null) continue;

                    if (parsedResult.DocumentElement.FirstChild is XmlElement node)
                        Map[node.InnerText] = Convert.ToInt32(new string(node.Attributes["href"].Value.Where(
                            char.IsDigit).ToArray()));
                }
            }
        }
    }
}
