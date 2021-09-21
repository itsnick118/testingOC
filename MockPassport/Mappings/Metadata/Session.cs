using System.Net;
using MockPassport.Mappings.Strings;
using WireMock.RequestBuilders;
using WireMock.ResponseBuilders;
using WireMock.Server;

namespace MockPassport.Mappings.Metadata
{
    public class Session : IMapping
    {
        public FluentMockServer Setup(FluentMockServer server, IEnvironment environment)
        {
            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.InitializeSessionParam)))
                .WithTitle("Initialize session param")
                .RespondWith(Response.Create()
                    .WithStatusCode(HttpStatusCode.OK));

            return server;
        }
    }
}