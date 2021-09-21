using System;
using System.Net;
using APITests.Passport.Json.Configuration.Model;
using NUnit.Framework;

namespace APITests.Passport.Json.Configuration
{
    [TestFixture]
    public class RootManifestTests
    {
        private readonly PassportClient _client = new PassportClient(
            new EnvironmentConfiguration(Environment.PASSPORT_2_5));

        private RootManifest _rootManifest;

        [OneTimeSetUp]
        public void SetUp()
        {
            var query = new RootManifestQuery(new Version(1, 0, 1, 0));
            _rootManifest = _client.GetRootManifest(query, true);
        }

        [Test]
        public void AutoUpdateUrlProperlyResolves()
        {
            var url = _rootManifest.AutoUpdate.AutoUpdateUrl;
            Assert.AreEqual(_client.GetStatusCode(url, true), HttpStatusCode.OK);
            using (var response = _client.HttpGet(url, true))
            {
                Assert.IsTrue(response.IsSuccessStatusCode);
            }
        }

        [Test]
        public void DefinesAtLeastOneModuleApp()
        {
            var totalApps = 0;

            foreach (var availableApp in _rootManifest.AvailableApps)
            {
                totalApps += availableApp.Value.Length;
            }

            Assert.NotZero(totalApps);
        }

        [Test]
        public void ModuleManifestUrlsProperlyResolve()
        {
            foreach (OfficeApp officeApp in Enum.GetValues(typeof(OfficeApp)))
            {
                foreach (var appDefinition in _rootManifest.AvailableApps[officeApp])
                {
                    Assert.AreEqual(
                        _client.GetStatusCode(AddVersionsToUrl(appDefinition.Uri), true),
                        HttpStatusCode.OK,
                        $"{appDefinition.Uri} failed to resolve");
                }
            }
        }

        [Test]
        public void OnlyOneOcAppPerOfficeAppIsMarkedDefault()
        {
            foreach (OfficeApp app in Enum.GetValues(typeof(OfficeApp)))
            {
                if (!_rootManifest.AvailableApps.ContainsKey(app)) continue;

                var numberDefault = 0;

                foreach (var module in _rootManifest.AvailableApps[app])
                {
                    if (module.Default) numberDefault++;
                }

                Assert.AreEqual(1, numberDefault);
            }
        }

        private string AddVersionsToUrl(string url)
        {
            return $"{url}&spaVersion=1.0.1.0&msiVersion=1.0.1.0";
        }
    }
}
