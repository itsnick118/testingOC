using System.Net;
using MockPassport.Mappings.Strings;
using WireMock.Matchers;
using WireMock.RequestBuilders;
using WireMock.ResponseBuilders;
using WireMock.Server;

namespace MockPassport.Mappings.Metadata
{
    public class GlowRootStatusTrace:IMapping
    {
       public FluentMockServer Setup(FluentMockServer server, IEnvironment environment)
        {
            server
                .Given(Request.Create()
                    .UsingGet()
                    .WithPath(new WildcardMatcher(Endpoint.GlowRootEnabled)))
                .WithTitle("GlowRoot")
                .RespondWith(Response.Create()
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeader(HeaderKey.ContentType, ContentType.ApplicationXml));
            server
                .Given(Request.Create()
                    .UsingGet()
                    .WithPath(new WildcardMatcher(Endpoint.GlowRootDisabled)))
                .WithTitle("GlowRoot")
                .RespondWith(Response.Create()
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithBody(@"{ msg: ""Unable to access the Glowroot.The Glowroot is disabled or not installed!!!""}")
                    .WithHeader(HeaderKey.ContentType, ContentType.ApplicationXml));
            return server;
        }
    }
}
