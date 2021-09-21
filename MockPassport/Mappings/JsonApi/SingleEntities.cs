using System.Collections.Generic;
using System.Net.Http;
using MockPassport.Mappings.Strings;
using WireMock.Server;

namespace MockPassport.Mappings.JsonApi
{
    public class SingleEntities: IMapping, IUpdatable
    {

        private readonly IDictionary<string, string> _listToSingleEntityMap = new Dictionary<string, string>
        {
            {ScreenName.MatterList, ScreenName.MatterSummary}
        };

        public FluentMockServer Setup(FluentMockServer server, IEnvironment environment)
        {
            foreach (var singleEntity in _listToSingleEntityMap.Values)
            {
                MappingHelpers.CreateJsonEntitySetup(singleEntity, server, environment);
            }
            return server;
        }

        public void Update(HttpClient client, IEnvironment environment, EntityIdMap map)
        {
            foreach (var mapping in _listToSingleEntityMap)
            {
                var topEntity = MappingHelpers.GetIdForFirstJsonEntity(mapping.Key, client, environment);

                if (topEntity >= 0)
                {
                    MappingHelpers.CreateJsonEntityUpdate(mapping.Value, topEntity, client,
                        environment);
                }
            }
        }
    }
}
