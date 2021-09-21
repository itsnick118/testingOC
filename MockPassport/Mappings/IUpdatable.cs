
using System.Net.Http;

namespace MockPassport.Mappings
{
    public interface IUpdatable
    {
        void Update(HttpClient client, IEnvironment environment, EntityIdMap map);
    }
}
