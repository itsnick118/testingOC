using System.Collections.Generic;
using System.Net.Http;
using MockPassport.Mappings.Strings;
using WireMock.RequestBuilders;
using WireMock.ResponseBuilders;
using WireMock.Server;

namespace MockPassport.Mappings.Metadata
{
    public class Manifests: IMapping, IUpdatable
    {
        private const string OutlookMatterManifestFile = "outlook_matter_manifest";
        private const string OutlookSpendManifestFile = "outlook_spend_manifest";
        private const string OutlookGlobalDocumentsManifestFile = "outlook_globaldocs_manifest";
        private const string RootManifestFile = "rootmanifest";

        public FluentMockServer Setup(FluentMockServer server, IEnvironment environment)
        {
            // The order in this file is not guaranteed to apply, so:
            //    -  Use priority 1 for specific app/module manifests
            //    -  Use priority 10 for fallthrough (e.g. core manifest)

            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.OcManifest))
                    .WithParam(ParamKey.OfficeApp, "outlook")
                    .WithParam(ParamKey.Module, "matter"))
                .WithTitle("OC Outlook/matter manifest")
                .AtPriority(1)
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, OutlookMatterManifestFile)));

            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.OcManifest))
                    .WithParam(ParamKey.OfficeApp, "outlook")
                    .WithParam(ParamKey.Module, "spend"))
                .WithTitle("OC Outlook/spend manifest")
                .AtPriority(1)
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, OutlookSpendManifestFile)));

            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.OcManifest)))
                .WithTitle("OC root manifest")
                .AtPriority(10)
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, RootManifestFile)));

            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.OcManifest))
                    .WithParam(ParamKey.OfficeApp, "outlook")
                    .WithParam(ParamKey.Module, "globalDocuments"))
                .WithTitle("OC global documents manifest")
                .AtPriority(1)
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, OutlookGlobalDocumentsManifestFile)));

            return server;
        }

        public void Update(HttpClient client, IEnvironment environment, EntityIdMap entityIdMap)
        {
            new MappingUpdateRequest { 
                Endpoint = Endpoint.OcManifest,
                Method = HttpMethod.Get,
                FileName = OutlookMatterManifestFile,
                Title = "OC outlook/matter manifest",
                ContentType = ContentType.TextHtml,
                Environment = environment,
                Parameters = new Dictionary<string, string>
                {
                    { ParamKey.OfficeApp, "outlook" },
                    { ParamKey.Module, "matter" },
                    { ParamKey.SpaVersion, "1.0.1.0" },
                    { ParamKey.MsiVersion, "1.0.1.0" }
                }
            }.UpdateFile(client);

            new MappingUpdateRequest { 
                Endpoint = Endpoint.OcManifest,
                Method = HttpMethod.Get,
                FileName = OutlookSpendManifestFile,
                Title = "OC outlook/spend manifest",
                ContentType = ContentType.TextHtml,
                Environment = environment,
                Parameters = new Dictionary<string, string>
                {
                    { ParamKey.OfficeApp, "outlook" },
                    { ParamKey.Module, "spend" },
                    { ParamKey.SpaVersion, "1.0.1.0" },
                    { ParamKey.MsiVersion, "1.0.1.0" }
                }
            }.UpdateFile(client);

            new MappingUpdateRequest { 
                Endpoint = Endpoint.OcManifest,
                Method = HttpMethod.Get,
                FileName = RootManifestFile,
                Title = "OC root manifest",
                ContentType = ContentType.TextHtml,
                Environment = environment,
                Parameters = new Dictionary<string, string>
                {
                    { ParamKey.SpaVersion, "1.0.1.0" },
                    { ParamKey.MsiVersion, "1.0.1.0" }
                }
            }.UpdateFile(client);

            new MappingUpdateRequest
            {
                Endpoint = Endpoint.OcManifest,
                Method = HttpMethod.Get,
                FileName = OutlookGlobalDocumentsManifestFile,
                Title = "OC outlook/global documents manifest",
                ContentType = ContentType.TextHtml,
                Environment = environment,
                Parameters = new Dictionary<string, string>
                {
                    { ParamKey.OfficeApp, "outlook" },
                    { ParamKey.Module, "globalDocuments" },
                    { ParamKey.SpaVersion, "1.0.1.0" },
                    { ParamKey.MsiVersion, "1.0.1.0" }
                }
            }.UpdateFile(client);
        }
    }
}