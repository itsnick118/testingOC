using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;
using MockPassport.Mappings.Strings;
using WireMock.RequestBuilders;
using WireMock.ResponseBuilders;
using WireMock.Server;

namespace MockPassport.Mappings.Metadata
{
    public class EntityNames: IMapping, IUpdatable
    {
        public static IDictionary<string, string> MetaDataFiles => new Dictionary<string, string>
        {
            { "AdjustmentLineItem", "entityNames_adjustmentlineitem" },
            { "DetailLineItem", "entityNames_detaillineitem" },
            { "EmailDocument", "entityNames_emailDocument" },
            { "InvoiceHeader", "entityNames_invoiceheader" },
            { "InvoiceMatterManagementDoc", "entityNames_invoicemattermanagementdocument" },
            { "ListScreen", "entityNames_listscreen" },
            { "Matter", "entityNames_matter" },
            { "MatterEvent", "entityNames_matterevent" },
            { "MatterManagementDoc", "entityNames_mattermanagementdocument" },
            { "MatterMatterManagementDoc", "entityNames_mattermattermanagementdocument" },
            { "MatterNarrative", "entityNames_matternarrative" },
            { "MatterPerson", "entityNames_matterperson" },
            { "MatterPersonRoleInvolvementType", "entityNames_matterpersonroleinvolvementtype" },
            { "PassportCmisObject", "entityNames_passportcmisobject" },
            { "Person", "entityNames_person" },
            { "SavedSearchUserPrefsDefaultView", "entityNames_searchparams_userprefs_userdefaultsavedview" }
        };

        private readonly string _metaDataEndpoint = Endpoint.GetEntity(EntityName.MetaData);

        public FluentMockServer Setup(FluentMockServer server, IEnvironment environment)
        {
            var map = new MetadataMap(environment);

            foreach (var file in MetaDataFiles)
            {
                server = SetupByName(server, file.Key, environment, file.Value + MetadataMap.ByNameSuffix);
                server = SetupById(server, map, file.Key, environment, file.Value + MetadataMap.ByIdSuffix);
            }

            server = SetupByName(server, new[] { "SavedSearch", "UserPreferences", "UserDefaultSavedView" },
                environment,
                MetaDataFiles["AdjustmentLineItem"] + MetadataMap.ByNameSuffix);

            return server;
        }

        public void Update(HttpClient client, IEnvironment environment, EntityIdMap entityIdMap)
        {
            foreach (var file in MetaDataFiles)
            {
                UpdateByName(client, environment, file.Key, file.Value + MetadataMap.ByNameSuffix);
            }

            var map = new MetadataMap(environment);
            foreach (var file in MetaDataFiles)
            {
                UpdateCallsById(client, environment, map, file.Key, file.Value + MetadataMap.ByIdSuffix);
            }

            UpdateByName(client, environment, 
                new[] { "SavedSearch", "UserPreferences", "UserDefaultSavedView" },
                MetaDataFiles["SavedSearchUserPrefsDefaultView"] + MetadataMap.ByNameSuffix);
        }
        
        private FluentMockServer SetupByName(FluentMockServer server, string searchParam, 
            IEnvironment environment, string fileName)
        {
            return SetupByName(server, new[] {searchParam}, environment, fileName);
        }

        private FluentMockServer SetupByName(FluentMockServer server, string[] searchParams, 
            IEnvironment environment, string fileName)
        {
            var searchStringBuilder = new StringBuilder("(");
            foreach (var searchParam in searchParams)
            {
                searchStringBuilder.Append($"(name)(EQUALS)({searchParam});");
            }

            var searchString = searchStringBuilder.ToString().TrimEnd(';') + ')';
            
            var title = $"{string.Join(", ", searchParams)} entity names";

            server
                .Given(Request.Create()
                    .UsingGet()
                    .WithPath(p => p.Contains(_metaDataEndpoint))
                    .WithParam(ParamKey.SearchParameters, searchString))
                .WithTitle(title)
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, fileName))
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeader("Set-Cookie", "JSESSIONID=E4BE0A348440986A570C392E4B4E6BD7; Path=/Passport; Secure; HttpOnly")
                    .WithHeader("Set-Cookie", "sessionIdForCognos=E4BE0A348440986A570C392E4B4E6BD7; Path=/; Secure; HttpOnly")
                    .WithHeader(HeaderKey.ContentType, ContentType.ApplicationXml));

            return server;
        }
        
        private FluentMockServer SetupById(FluentMockServer server, MetadataMap map, string entity, 
            IEnvironment environment, string fileName)
        {
            if (!map.EntityHrefs.ContainsKey(entity)) return server;

            var title = $"ID endpoint for {entity}";

            server
                .Given(Request.Create()
                    .UsingGet()
                    .WithPath(p => p.Contains(_metaDataEndpoint + map.EntityHrefs[entity])))
                .WithTitle(title)
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, fileName))
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeader("Set-Cookie", "JSESSIONID=E4BE0A348440986A570C392E4B4E6BD7; Path=/Passport; Secure; HttpOnly")
                    .WithHeader("Set-Cookie", "sessionIdForCognos=E4BE0A348440986A570C392E4B4E6BD7; Path=/; Secure; HttpOnly")
                    .WithHeader(HeaderKey.ContentType, ContentType.ApplicationXml));

            return server;
        }

        private void UpdateByName(HttpClient client, IEnvironment environment, string searchParam, 
            string fileName)
        {
            UpdateByName(client, environment, new []{ searchParam }, fileName);
        }
        
        private void UpdateByName(HttpClient client, IEnvironment environment, string[] searchParams,
            string fileName)
        {
            var searchStringBuilder = new StringBuilder("(");

            foreach (var searchParam in searchParams)
            {
                searchStringBuilder.Append($"(name)(EQUALS)({searchParam});");
            }

            var searchString = searchStringBuilder.ToString().TrimEnd(';') + ')';

            var title = $"{string.Join(", ", searchParams)} entity name(s)";

            new MappingUpdateRequest
            {
                Endpoint = _metaDataEndpoint,
                Method = HttpMethod.Get,
                FileName = fileName,
                Title = title,
                Environment = environment,
                Parameters = new Dictionary<string, string>
                {
                    { ParamKey.SearchParameters, searchString }
                }
            }.UpdateFile(client);
        }

        private void UpdateCallsById(HttpClient client, IEnvironment environment, MetadataMap map, string entity, string fileName)
        {
            if (!map.EntityHrefs.ContainsKey(entity)) return;
                
            var title = $"ID endpoint for {entity}";

            new MappingUpdateRequest
            {
                Endpoint = Endpoint.GetEntity(EntityName.MetaData, map.EntityHrefs[entity]),
                Method = HttpMethod.Get,
                FileName = fileName,
                Title = title,
                Environment = environment
            }.UpdateFile(client, false);
        }
    }
}
