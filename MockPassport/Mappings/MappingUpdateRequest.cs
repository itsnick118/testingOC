using System;
using System.Net.Http;
using System.Text;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.WebUtilities;

namespace MockPassport.Mappings
{
    public class MappingUpdateRequest
    {
        public string Endpoint { get; set; }
        public string Content { get; set; }
        public HttpMethod Method { get; set; }
        public string FileName { get; set; }
        public string Title { get; set; }
        public string ContentType { get; set; }
        public IEnvironment Environment { get; set; }
        public IDictionary<string, string> Parameters { get; set; }
        public bool ExpectContinue { get; set; }
        public bool StrictTransport { get; set; }

        public void UpdateFile(HttpClient client, bool saveHeader=true)
        {
            Console.WriteLine(System.Environment.NewLine + "Processing " + Title);

            var uri = Parameters == null
                ? new Uri(client.BaseAddress + Endpoint)
                : new Uri(QueryHelpers.AddQueryString(client.BaseAddress + Endpoint, Parameters));

            var request = new HttpRequestMessage {
                RequestUri = uri,
                Content = Content == null
                    ? null
                    : new StringContent(
                        Content,
                        Encoding.UTF8,
                        ContentType ?? Strings.ContentType.TextHtml),
                Method = Method
            };

            var response = client.SendAsync(request);
            
            var status = response.Result.StatusCode;

            Console.ForegroundColor =
                status != HttpStatusCode.OK
                    ? ConsoleColor.Yellow
                    : ConsoleColor.Green;

            var url = request.RequestUri.AbsoluteUri;
            if (!string.IsNullOrWhiteSpace(url))
            {
                Console.WriteLine("   url: " + url);
            }

            Console.WriteLine("   status: " + status);

            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(System.Environment.NewLine + "--------------------------------------------");

            using (var content = response.Result.Content)
            {
                if (!response.Result.IsSuccessStatusCode) return;

                var responseDirectory =
                    new DirectoryInfo(Path.Combine(Environment.BaseFilePath.FullName, "Responses"));
                if (!responseDirectory.Exists) responseDirectory.Create();

                var responseFilePath = Path.Combine(responseDirectory.FullName, FileName);

                var responseBody = content.ReadAsStringAsync().Result;

                File.WriteAllText(responseFilePath, responseBody);

                if (!saveHeader) return;

                var headerDirectory = new DirectoryInfo(Path.Combine(Environment.BaseFilePath.FullName, "Headers"));
                if (!headerDirectory.Exists) headerDirectory.Create();

                var headerFilePath = Path.Combine(headerDirectory.FullName, FileName);

                var responseHeader = response.Result.Headers.ToString();

                if (!StrictTransport)
                {
                    var transportHeaderPattern = @"Strict-Transport-Security.*\n";
                    responseHeader = Regex.Replace(responseHeader, transportHeaderPattern, string.Empty);
                }

                File.WriteAllText(headerFilePath, responseHeader);
            }
        }
    }
}
 