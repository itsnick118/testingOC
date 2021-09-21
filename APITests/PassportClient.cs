using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using APITests.Passport;
using APITests.Passport.Json.Configuration;
using APITests.Passport.Json.Configuration.Model;
using APITests.Passport.ScreenScraping;
using Newtonsoft.Json;

namespace APITests
{
    internal class PassportClient
    {
        private readonly EnvironmentConfiguration _config;

        public string BaseUrl;

        public PassportClient(EnvironmentConfiguration config)
        {
            _config = config;
            BaseUrl = _config.BaseUrl;
        }

        public HttpResponseMessage HttpGet(string endpoint, bool elevated)
        {
            var url = $"{BaseUrl.TrimEnd('/')}/{endpoint.TrimStart('/')}";
            HttpResponseMessage returnMessage;

            using (var httpClient = new HttpClient())
            {
                var authorizationHeaders = elevated
                    ? _config.GetStandardUserHeaders()
                    : _config.GetElevatedUserHeaders();

                httpClient.DefaultRequestHeaders.Authorization = authorizationHeaders;
                httpClient.DefaultRequestHeaders.Add("User-Agent", "Outlook");
                using (var response = httpClient.GetAsync(url))
                {
                    response.Wait();
                    returnMessage = response.Result;
                }
            }

            return returnMessage;
        } 

        private string BuildQueryString(IDictionary<string, string> queryDict)
        {
            return string.Join(
                "&", 
                queryDict.Select(queryParam => $"{queryParam.Key}={queryParam.Value}").ToArray());
        }

        private string CreateEndPointString(string endPoint, IQuery query)
        {
            var queryDict = query.AsDictionary();
            return endPoint + (queryDict.Any()
                       ? "?" + BuildQueryString(queryDict) 
                       : string.Empty);
        }
        
        public AppModuleManifest GetModuleManifest(AppModuleQuery query, bool asElevatedUser)
        {
            AppModuleManifest manifest;

            var endpoint = CreateEndPointString(EndPoints.OC_MANIFEST, query);

            using (var response = HttpGet(endpoint, asElevatedUser))
            {
                using (var content = response.Content.ReadAsStringAsync())
                {
                    content.Wait();

                    try
                    {
                        manifest = JsonConvert.DeserializeObject<AppModuleManifest>(content.Result,
                            new JsonSerializerSettings
                            {
                                MissingMemberHandling = MissingMemberHandling.Error
                            });
                    }
                    catch (Exception e)
                    {
                        throw new JsonException("Could not deserialize module manifest. Error: " + e.Message);
                    }
                }
            }

            return manifest;
        }

        public RootManifest GetRootManifest(RootManifestQuery query, bool asElevatedUser)
        {
            RootManifest manifest;

            var endpoint = CreateEndPointString(EndPoints.OC_MANIFEST, query);

            using (var response = HttpGet(endpoint, asElevatedUser))
            {
                using (var content = response.Content.ReadAsStringAsync())
                {
                    content.Wait();

                    try
                    {
                        manifest = JsonConvert.DeserializeObject<RootManifest>(content.Result,
                            new JsonSerializerSettings
                            {
                                MissingMemberHandling = MissingMemberHandling.Error
                            });
                    }
                    catch (Exception e)
                    {
                        throw new JsonException("Could not deserialize manifest. Error: " + e.Message);
                    }
                }
            }

            return manifest;
        }
        
        public PassportScreen GetScreen(ScreenQuery query, bool asElevatedUser)
        {
            var endpoint = CreateEndPointString(EndPoints.MOBILE_UI, query);
            PassportScreen result;

            using (var response = HttpGet(endpoint, asElevatedUser))
            {
                using (var content = response.Content.ReadAsStringAsync())
                {
                    content.Wait();
                    result = new PassportScreen(content.Result);
                }
            }

            return result;
        }

        public HttpStatusCode GetStatusCode(string endpoint, bool asElevatedUser)
        {
            HttpStatusCode result;

            using (var response = HttpGet(endpoint, asElevatedUser))
            {
                result = response.StatusCode;
            }

            return result;
        }

        public T StringToEnum<T>(string inputString)
        {
            return (T) Enum.Parse(typeof(T), inputString);
        }
    }
}
