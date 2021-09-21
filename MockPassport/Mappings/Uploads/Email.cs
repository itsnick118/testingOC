using System;
using System.Collections.Generic;
using System.Net;
using WireMock.RequestBuilders;
using WireMock.ResponseBuilders;
using WireMock.Server;

namespace MockPassport.Mappings.Uploads
{
    public class Email : IMapping
    {
        public FluentMockServer Setup(FluentMockServer server, IEnvironment environment)
        {
            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains("/Passport/datacert/api/entity/EmailDocument"))
                    .UsingPost())
                .WithTitle("Single e-mail upload")
                .RespondWith(Response.Create()
                    .WithStatusCode(HttpStatusCode.Created)
                    .WithHeaders(new Dictionary<string, string>
                    {
                        { "Server", "None" },
                        { "Location", "https://localhost:7777/Passport/datacert/api/entity/EmailDocument/12413" },
                        { Strings.HeaderKey.ContentType, "application/xml;charset=UTF-8"},
                        { "Content-Length", "0" }
                    })
                    .WithDelay(TimeSpan.FromMilliseconds(500L)));

            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains("/Passport/datacert/api/entity/PassportCmisObject"))
                    .WithParam("parentEntityName", "Matter")
                    .UsingPost())
                .WithTitle("Single e-mail upload (CMIS)")
                .RespondWith(Response.Create()
                    .WithStatusCode(HttpStatusCode.Created)
                    .WithHeaders(new Dictionary<string, string>
                    {
                        { "Server", "None" },
                        { "Location", "https://localhost:7777/Passport/datacert/api/entity/EmailDocument/12413" },
                        { Strings.HeaderKey.ContentType, "application/xml;charset=UTF-8"},
                        { "Content-Length", "0" }
                    })
                    .WithDelay(TimeSpan.FromMilliseconds(500L)));

            return server;
        }
    }
}