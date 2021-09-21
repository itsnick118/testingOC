using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using Microsoft.AspNetCore.WebUtilities;
using MockPassport.Mappings.Strings;
using Newtonsoft.Json;
using WireMock.Matchers;
using WireMock.RequestBuilders;
using WireMock.ResponseBuilders;
using WireMock.Server;

namespace MockPassport.Mappings
{
    public class MappingHelpers
    {
        private const string FirstNineteenPagesPattern = @"""pageNumber"":""\d""|""pageNumber"":""1\d""";

        private const string CookieString =
            "JSESSIONID=E4BE0A348440986A570C392E4B4E6BD7; Path=/Passport; Secure; HttpOnly";

        private const string SingleEntityView = "_single_view";

        public static void CreateJsonEntityListSetup(string screenName, FluentMockServer server, IEnvironment environment)
        {
            var file = GetFileName(screenName);
            CreateJsonMetadataSetup(screenName, server, environment);

            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.ListScreenShowJson))
                    .WithParam(ParamKey.ScreenName, screenName)
                    .WithParam(ParamKey.DisableAutoLoading, "false"))
                .WithTitle(screenName)
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, file))
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeader("Set-Cookie", CookieString)
                    .WithHeader(HeaderKey.ContentType, ContentType.ApplicationJson));
        }

        public static void CreateJsonEntitySetup(string screenName, FluentMockServer server, IEnvironment environment)
        {
            var file = GetFileName(screenName + SingleEntityView);

            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.ItemScreenShowJson))
                    .WithParam(ParamKey.ScreenName, screenName)
                    .WithParam(ParamKey.Mode, "view"))
                .WithTitle(screenName)
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, file))
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeader("Set-Cookie", CookieString)
                    .WithHeader(HeaderKey.ContentType, ContentType.ApplicationJson)
                    .WithHeader(HeaderKey.MetaData, FromFile.GetMetadataHeader(environment, file)));
        }

        public static void CreateExtendedJsonListSetup(string screenName, FluentMockServer server, IEnvironment environment)
        {
            var file = GetFileName(screenName);
            CreateJsonMetadataSetup(screenName, server, environment);

            dynamic repeatingPagesBody = JsonConvert.DeserializeObject(FromFile.GetBody(environment, file));
            repeatingPagesBody.page.totalPages = 20;
            repeatingPagesBody.page.pageNumber = 1;
            repeatingPagesBody.page.totalRecords = 500;
            string repeatingPagesBodyStr = JsonConvert.SerializeObject(repeatingPagesBody);
            

            dynamic lastPageBody = JsonConvert.DeserializeObject(FromFile.GetBody(environment, file));
            lastPageBody.page.totalPages = 20;
            lastPageBody.page.pageNumber = 20;
            lastPageBody.page.totalRecords = 500;
            string lastPageBodyStr = JsonConvert.SerializeObject(lastPageBody);

            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.ListScreenShowJson))
                    .WithParam(ParamKey.ScreenName, screenName)
                    .WithParam(ParamKey.DisableAutoLoading, "false")
                    .WithBody(new RegexMatcher(FirstNineteenPagesPattern)))
                .WithTitle(screenName)
                .AtPriority(1)
                .RespondWith(Response.Create()
                    .WithBody(repeatingPagesBodyStr)
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeader(HeaderKey.ContentType, ContentType.ApplicationJson));

            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.ListScreenShowJson))
                    .WithParam(ParamKey.ScreenName, screenName)
                    .WithParam(ParamKey.DisableAutoLoading, "false")
                    .WithParam(ParamKey.Grandparent, "True"))
                .WithTitle(screenName)
                .AtPriority(10)
                .RespondWith(Response.Create()
                    .WithBody(lastPageBodyStr)
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeader(HeaderKey.ContentType, ContentType.ApplicationJson));
        }

        public static void CreateJsonListUpdate(string screenName, HttpClient client, IEnvironment environment)
        {
            new MappingUpdateRequest
            {
                Endpoint = Endpoint.ListScreenShowJson,
                Parameters = new Dictionary<string, string>
                {
                    { ParamKey.ScreenName, screenName },
                    { ParamKey.Metadata, "true" }
                },
                Content = string.Empty,
                Method = HttpMethod.Post,
                FileName = GetFileName(screenName, true),
                Title = screenName,
                ExpectContinue = true,
                Environment = environment,
                StrictTransport = false,
            }.UpdateFile(client);

            new MappingUpdateRequest
            {
                Endpoint = Endpoint.ListScreenShowJson,
                Parameters = new Dictionary<string, string>
                {
                    { ParamKey.ScreenName, screenName },
                    { ParamKey.DisableAutoLoading, "false" },
                    { ParamKey.Grandparent, "True" }
                },
                Content = "{\"page\": {\"pageNumber\":\"1\",\"pageSize\":\"50\",\"sortInfo\":null}, \"filters\": []}",
                Method = HttpMethod.Post,
                FileName = GetFileName(screenName),
                Title = screenName,
                ContentType = ContentType.ApplicationJson,
                Environment = environment,
                StrictTransport = false,
            }.UpdateFile(client);
        }

        public static void CreateJsonEntityUpdate(string screenName, int entityId, HttpClient client, IEnvironment environment)
        {
            var file = GetFileName(screenName + SingleEntityView);

            new MappingUpdateRequest
            {
                Endpoint = Endpoint.ItemScreenShowJson,
                Parameters = new Dictionary<string, string>
                {
                    { ParamKey.ScreenName, screenName },
                    { ParamKey.EntityInstanceId, entityId.ToString() },
                    { ParamKey.Mode, "view" },
                },
                Content = string.Empty,
                Method = HttpMethod.Post,
                FileName = file,
                Title = screenName,
                Environment = environment,
                StrictTransport = false,
            }.UpdateFile(client);
        }

        public static int GetIdForFirstJsonEntity(string screenName, HttpClient client, IEnvironment environment)
        {
            Console.WriteLine(Environment.NewLine + "Getting model entity ID for " + screenName);

            const string endpoint = Endpoint.ListScreenShowJson;
            var parameters = new Dictionary<string, string>
            {
                {ParamKey.ScreenName, screenName},
                {ParamKey.DisableAutoLoading, "false"},
                {ParamKey.Grandparent, "True"}
            };
            const string postContent = "{\"page\": {\"pageNumber\":\"1\",\"pageSize\":\"50\",\"sortInfo\":null}, \"filters\": []}";

            var uri = new Uri(QueryHelpers.AddQueryString(client.BaseAddress + endpoint, parameters));

            var request = new HttpRequestMessage
            {
                RequestUri = uri,
                Content = new StringContent(
                        postContent,
                        Encoding.UTF8,
                        ContentType.ApplicationJson),
                Method = HttpMethod.Post
            };

            var response = client.SendAsync(request);
            
            using (var content = response.Result.Content)
            {
                if (!response.Result.IsSuccessStatusCode)
                {
                    PrintError("Could not get model entity ID.");
                    return -1;
                }
                
                var responseBody = content.ReadAsStringAsync().Result;

                var responseJson = JsonConvert.DeserializeObject(responseBody) as dynamic;

                try
                {
                    var firstId = Convert.ToInt32(responseJson.list[0].id);
                    return firstId;
                }
                catch
                {
                    // ignoring but returning -1 below
                }
            }
            PrintError("Could not get model entity ID.");
            return -1;
        }

        private static void CreateJsonMetadataSetup(string screenName, FluentMockServer server, IEnvironment environment)
        {
            var file = GetFileName(screenName, true);
            server
                .Given(Request.Create()
                    .WithPath(p => p.Contains(Endpoint.ListScreenShowJson))
                    .WithParam(ParamKey.ScreenName, screenName)
                    .WithParam(ParamKey.Metadata, "true"))
                .WithTitle(screenName)
                .RespondWith(Response.Create()
                    .WithBody(FromFile.GetBody(environment, file))
                    .WithStatusCode(HttpStatusCode.OK)
                    .WithHeader("Set-Cookie", CookieString)
                    .WithHeader(HeaderKey.MetaData, FromFile.GetMetadataHeader(environment, file))
                    .WithHeader(HeaderKey.ContentType, ContentType.ApplicationJson));
        }

        private static string GetFileName(string screenName, bool isMetaData = false)
        {
            var baseName = Path.GetInvalidFileNameChars()
                .Aggregate(screenName, (current, badChar) => current.Replace(badChar, '_'));
            return (isMetaData ? baseName + "_meta" : baseName) + ".json";
        }

        private static void PrintError(string error)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(error);
            Console.ForegroundColor = ConsoleColor.White;
        }
    }
}