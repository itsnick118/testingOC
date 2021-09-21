using WireMock.Server;

namespace MockPassport.Mappings
{
    public interface IMapping
    {
        FluentMockServer Setup(FluentMockServer server, IEnvironment environment);
    }
}
