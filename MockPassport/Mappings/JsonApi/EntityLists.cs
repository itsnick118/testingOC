using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using MockPassport.Mappings.Strings;
using WireMock.Server;

namespace MockPassport.Mappings.JsonApi
{
    public class EntityLists : IMapping, IUpdatable
    {
        // the entities in this list will replay as the first page, as-is. Multiple pages aren't supported yet.
        private readonly IList<string> _entities = new List<string>
        {
            ScreenName.ToBeAcknowledgedByMe,
            ScreenName.InvoicesList,
            ScreenName.DetailLineItemList, 
            ScreenName.AdjustmentLineItemList,
            ScreenName.EmailDocumentCmisList,
            ScreenName.MatterDocumentCmisList,
            ScreenName.MatterNarrativesList,
            ScreenName.MatterPersonList,
            ScreenName.MatterEventList
        };

        // the entities in this list will be replayed over and over for each "page" until the app has 500 items.
        private readonly IList<string> _extendedEntities = new List<string>
        {
            ScreenName.MatterList,
            ScreenName.GlobalDocumentsCmisList
        };

        public FluentMockServer Setup(FluentMockServer server, IEnvironment environment)
        {
            foreach (var entity in _entities)
            {
                MappingHelpers.CreateJsonEntityListSetup(entity, server, environment);
            }
            foreach (var entity in _extendedEntities)
            {
                MappingHelpers.CreateExtendedJsonListSetup(entity, server, environment);
            }

            return server;
        }

        public void Update(HttpClient client, IEnvironment environment, EntityIdMap entityIdMap)
        {
            foreach (var entity in _entities.Concat(_extendedEntities))
            {
                MappingHelpers.CreateJsonListUpdate(entity, client, environment);
            }
        }
    }
}
