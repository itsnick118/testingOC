using System;
using APITests.Passport.Json.Configuration.Model;
using NUnit.Framework;

namespace APITests.Passport.Json.Configuration
{
    [TestFixture]
    public class OutlookMatterManifestTests
    {
        private readonly PassportClient _client = new PassportClient(
            new EnvironmentConfiguration(Environment.PASSPORT_2_5));

        private AppModuleManifest _moduleManifest;

        [OneTimeSetUp]
        public void SetUp()
        {
            var query = new AppModuleQuery(new Version(1, 0, 1, 0), OfficeApp.Outlook, OcModule.Matter);
            _moduleManifest = _client.GetModuleManifest(query, true);
        }
        
        [Test]
        public void DefaultPagesAreDefinedAndAvailable()
        {
            var offline = _moduleManifest.Application.Matter.Page.Offline;
            var online = _moduleManifest.Application.Matter.Page.Online;
            var favorite = _moduleManifest.Application.Matter.FavoritesPage;

            var pages = _moduleManifest.Application.Matter.Pages;

            Assert.IsTrue(pages.ContainsKey(_client.StringToEnum<ModulePage>(offline)));
            Assert.IsTrue(pages.ContainsKey(_client.StringToEnum<ModulePage>(online)));
            Assert.IsTrue(pages.ContainsKey(_client.StringToEnum<ModulePage>(favorite)));
        }
    }
}
