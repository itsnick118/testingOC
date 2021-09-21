using System.Net;
using MockPassport.Mappings.Strings;
using WireMock.Matchers;
using WireMock.RequestBuilders;
using WireMock.ResponseBuilders;
using WireMock.Server;

namespace MockPassport.Mappings.Metadata
{
    public class Ping: IMapping
    {
        public FluentMockServer Setup(FluentMockServer server, IEnvironment environment)
        {
            server
                .Given(Request.Create()
                    .UsingHead()
                    .WithPath(p => p.Contains(Endpoint.Index)))
                .WithTitle("Ping")
                .RespondWith(Response.Create()
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeader(HeaderKey.ContentType, ContentType.TextHtml));

            server
                .Given(Request.Create()
                    .UsingGet()
                    .WithPath(Endpoint.Passport))
                .WithTitle("Base Passport call for allowing certs")
                .RespondWith(Response.Create()
                    .WithStatusCode(HttpStatusCode.OK));

            return server;
        }
    }
}
