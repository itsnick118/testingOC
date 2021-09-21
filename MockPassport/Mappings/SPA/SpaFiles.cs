using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using MockPassport.Mappings.Strings;
using WireMock.Matchers;
using WireMock.RequestBuilders;
using WireMock.ResponseBuilders;
using WireMock.Server;

namespace MockPassport.Mappings.SPA
{
    public class SpaFiles : IMapping
    {
        public FluentMockServer Setup(FluentMockServer server, IEnvironment environment)
        {
            var spaFolder = new DirectoryInfo(@"..\..\..\..\Office UI");
            var distFolder = Path.Combine(spaFolder.FullName, "dist");

            if (!Directory.Exists(distFolder))
            {
                var buildPromptInfo = new ProcessStartInfo
                {
                    FileName = "cmd",
                    WorkingDirectory = spaFolder.FullName,
                    RedirectStandardInput = true,
                    UseShellExecute = false
                };

                var buildPrompt = Process.Start(buildPromptInfo);
                buildPrompt?.StandardInput.WriteLine("npm run build & exit");
                buildPrompt?.WaitForExit();

                if (buildPrompt == null || buildPrompt.ExitCode != 0 || !Directory.Exists(distFolder))
                {
                    Console.WriteLine("Warning: something went wrong building SPA code");
                    return server;
                }
            }

            var directoryStack = new Stack<DirectoryInfo>();
            directoryStack.Push(new DirectoryInfo(distFolder));

            while (directoryStack.Count > 0)
            {
                var currentDirectory = directoryStack.Pop();
                var files = currentDirectory.GetFiles();

                foreach (var file in files)
                {
                    var relativeFileName = file.FullName.Remove(0, distFolder.Length).Replace('\\', '/');

                    server
                        .Given(Request.Create()
                            .WithPath(new RegexMatcher(Endpoint.StaticResourcesRegex))
                            .WithPath(p => p.EndsWith(relativeFileName))
                            .UsingGet())
                        .WithTitle($"OC SPA: {relativeFileName}")
                        .RespondWith(Response.Create()
                            .WithHeader(HeaderKey.ContentType, GetContentTypeByExtension(file.Extension))
                            .WithBody(File.ReadAllBytes(file.FullName)));
                }

                var dirs = currentDirectory.GetDirectories();
                foreach (var dir in dirs)
                {
                    directoryStack.Push(dir);
                }
            }

            server
                .Given(Request.Create()
                    .WithPath(new RegexMatcher(Endpoint.StaticResourcesRegex))
                    .WithPath(p => p.EndsWith("/"))
                    .UsingGet())
                .WithTitle("OC SPA: root map to index.html")
                .RespondWith(Response.Create()
                    .WithHeader(HeaderKey.ContentType, ContentType.TextHtml)
                    .WithBody(File.ReadAllBytes(Path.Combine(distFolder, "index.html"))));

            return server;
        }

        private string GetContentTypeByExtension(string fileExtension)
        {
            switch (fileExtension)
            {
                case ".css":
                    return ContentType.TextCss;
                case ".svg":
                    return ContentType.ImageSvg;
                case ".woff":
                    return ContentType.FontWoff;
                case ".woff2":
                    return ContentType.FontWoff2;
                case ".js":
                    return ContentType.ApplicationJavascript;
                case ".html":
                    return ContentType.TextHtml;
                default:
                    return ContentType.TextHtml;
            }
        }
    }
}
