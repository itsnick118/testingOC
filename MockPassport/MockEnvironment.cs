using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
using MockPassport.Mappings;
using WireMock.Logging;
using WireMock.Server;
using WireMock.Settings;

namespace MockPassport
{
    public static class MockEnvironment
    {
        private static FluentMockServer _server;
        private static readonly object Padlock = new object();

        private static readonly FluentMockServerSettings Settings = new FluentMockServerSettings
        {
            UseSSL = true,
            Port = 7777
        };

        public static string GetMatchHistoryLog()
        {
            var log = new StringBuilder();
            foreach (var entry in _server.LogEntries)
            {
                var title = entry.MappingTitle;
                if (string.IsNullOrEmpty(title))
                {
                    if (!entry.MappingGuid.HasValue)
                    {
                        log.AppendLine("  Unmatched request: " + entry.RequestMessage.AbsoluteUrl);
                        log.AppendLine("    with params: " + entry.RequestMessage.RawQuery);
                        log.AppendLine("    with body: " + entry.RequestMessage.Body);
                    }
                    else
                    {
                        log.AppendLine("  Missing title for request: " + entry.RequestMessage.AbsoluteUrl);
                    }
                }
                else
                {
                    log.AppendLine("  " + title);
                }
            }
            return log.ToString();
        }

        public static void Start(string environmentName)
        {
            IEnvironment environment;
            try
            {
                environment = GetEnvironmentMetaData(environmentName);
            }
            catch
            {
                Console.WriteLine($"Could not retrieve metadata for environment {environmentName}.");
                Console.WriteLine("Check the spelling and try again.");
                return;
            }

            lock (Padlock)
            {
                if (_server != null) return;
                _server = FluentMockServer.Start(Settings);

                _server.LogEntriesChanged += (sender, args) =>
                {
                    foreach (var argsNewItem in args.NewItems)
                    {
                        if (argsNewItem is LogEntry entry)
                        {
                            Console.WriteLine("---" + entry.MappingTitle + "---");
                            Console.WriteLine(entry.MappingTitle);
                            Console.WriteLine(entry.RequestMessage.AbsoluteUrl);
                            Console.WriteLine(entry.RequestMessage.RawQuery);
                            foreach (var header in entry.RequestMessage.Headers)
                            {
                                foreach (var value in header.Value)
                                {
                                    Console.WriteLine(header.Key + ": " + value);
                                }
                            }

                            Console.WriteLine("---Response---");
                            Console.WriteLine("Status Code: " + entry.ResponseMessage.StatusCode);
                            foreach (var header in entry.ResponseMessage.Headers)
                            {
                                foreach (var value in header.Value)
                                {
                                    Console.WriteLine(header.Key + ": " + value);
                                }
                            }
                            var bytes = entry.ResponseMessage.BodyAsBytes;
                            if (bytes != null)
                            {
                                //Console.WriteLine(Encoding.UTF8.GetString(bytes));
                            }
                            Console.WriteLine("======================================================");
                        }
                    }
                };

                Console.WriteLine("Setting up shared mappings");
                foreach (var t in GetTypesInNameSpace("MockPassport.Mappings", true))
                {
                    object instance;

                    try
                    {
                        instance = Activator.CreateInstance(t);
                    }
                    catch
                    {
                        continue;
                    }

                    if (instance is IMapping mapping)
                    {
                        Console.WriteLine($"\t{mapping.GetType().Name}");
                        _server = mapping.Setup(_server, environment);
                    }
                }

                Console.WriteLine("Setting up environment-specific mappings");
                foreach (var t in GetTypesInNameSpace(environmentName))
                {
                    object instance;

                    try
                    {
                        instance = Activator.CreateInstance(t);
                    }
                    catch
                    {
                        continue;
                    }

                    if (instance is IMapping mapping) {
                        Console.WriteLine($"\t{mapping.GetType().Name}");
                        _server = mapping.Setup(_server, environment);
                    }
                }
            }

            Console.WriteLine("Press any key to stop the server");
            Console.ReadKey();
            Console.WriteLine();
            Console.WriteLine("------------");
            Console.WriteLine("Request log:");
            Console.WriteLine("------------");
            Console.WriteLine(GetMatchHistoryLog());
            Console.WriteLine("Press any key to quit");
            Console.ReadKey();
        }

        public static void Update(string environment)
        {
            var environmentMetaData = GetEnvironmentMetaData(environment);
            var authArray = Encoding.ASCII.GetBytes(
                environmentMetaData.Username + ":" + environmentMetaData.Password);

            Console.WriteLine("Updating environment using base Url: " + environmentMetaData.BaseUri);

            var client = new HttpClient
            {
                BaseAddress = environmentMetaData.BaseUri
            };

            client.DefaultRequestHeaders.Authorization = 
                new AuthenticationHeaderValue("Basic", Convert.ToBase64String(authArray));

            client.DefaultRequestHeaders.UserAgent.ParseAdd("Outlook");

            var map = new EntityIdMap(client);

            var updatableTypes = GetTypesInNameSpace("MockPassport.Mappings", true);

            foreach (var t in updatableTypes)
            {
                object instance;

                try
                {
                    instance = Activator.CreateInstance(t);
                }
                catch
                {
                    continue;
                }

                var isMapping = instance is IMapping;
                var skipping = instance is IUpdatable && t.IsDefined(typeof(SkipAttribute), false);

                if (instance is IUpdatable && !skipping)
                {
                    Console.WriteLine("Updating: " + t);
                    ((IUpdatable)instance).Update(client, environmentMetaData, map);
                }
                else if (skipping)
                {
                    Console.WriteLine("Skipping: " + t);
                }
                else if (isMapping)
                {
                    Console.WriteLine("No update instructions found for: " + t);
                }
            }

            Console.WriteLine("Press any key to quit");
            Console.ReadKey();
        }

        public static void Record(string url, string environment)
        {
            var server = FluentMockServer.Start(new FluentMockServerSettings
            {
                UseSSL = true,
                Urls = new []{ "https://localhost:7777" },
                StartAdminInterface = true,
                ProxyAndRecordSettings = new ProxyAndRecordSettings
                {
                    Url = url,
                    SaveMapping = true,
                    SaveMappingToFile = true
                }
            });

            server.LogEntriesChanged += (sender, args) =>
            {
                foreach (var argsNewItem in args.NewItems)
                {
                    if (argsNewItem is LogEntry entry)
                    {
                        Console.WriteLine("---Request---");
                        Console.WriteLine(entry.RequestMessage.AbsoluteUrl);
                        Console.WriteLine(entry.RequestMessage.RawQuery);
                        Console.WriteLine(entry.RequestMessage.Body);

                        Console.WriteLine("---Response---");
                        Console.WriteLine("Status Code: " + entry.ResponseMessage.StatusCode);
                        var bytes = entry.ResponseMessage.BodyAsBytes;
                        if (bytes != null)
                        {
                            //Console.WriteLine(Encoding.UTF8.GetString(bytes));
                        }
                    }
                }
            };

            Console.ReadKey();
        }
        
        private static List<Type> GetTypesInNameSpace(string nameSpace, bool absolute=false)
        {
            nameSpace = absolute ? nameSpace : "MockPassport.Environments." + nameSpace;
            
            var q = from t in Assembly.GetExecutingAssembly().GetTypes()
                where t.IsClass && (t.Namespace?.StartsWith(nameSpace)).GetValueOrDefault(false)
                select t;
            return q.ToList();
        }

        private static IEnvironment GetEnvironmentMetaData(string environment)
        {
            var nameSpace = "MockPassport.Environments." + environment;
            var type = GetTypesInNameSpace(nameSpace, true)
                .Where(t => t.IsClass && (t.Namespace?.StartsWith(nameSpace)).GetValueOrDefault(false))
                .First(c => c.Name == "Environment");

            return (IEnvironment)Activator.CreateInstance(type);
        }
    }
}