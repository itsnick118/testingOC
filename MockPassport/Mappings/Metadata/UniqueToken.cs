using System;
using System.Net;
using MockPassport.Mappings.Strings;
using WireMock.RequestBuilders;
using WireMock.ResponseBuilders;
using WireMock.Server;

namespace MockPassport.Mappings.Metadata
{
    public class UniqueToken:IMapping
    {
       public FluentMockServer Setup(FluentMockServer server, IEnvironment environment)
        {
            server
                .Given(Request.Create()
                    .UsingGet()
                    .WithPath(p => p.Contains(Endpoint.GetUniqueToken)))
                .WithTitle("Get unique token")
                .RespondWith(Response.Create()
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeader(HeaderKey.ContentType, ContentType.TextPlain)
                    .WithBody(GetUniqueToken()));

            return server;
        }

        private string GetUniqueToken()
        {
            const string numbers = "0123456789";
            const string lowercaseChars = "abcdefghijklmnopqrstuvwxyz";
            var uppercaseChars = lowercaseChars.ToUpperInvariant();
            var chars = (lowercaseChars + uppercaseChars + numbers).ToCharArray();

            var random = new Random();
            var result = new char[12];

            for (var i = 0; i < 12; i++)
            {
                result[i] = chars[random.Next(chars.Length)];
            }

            return new string(result);
        }
    }
}
