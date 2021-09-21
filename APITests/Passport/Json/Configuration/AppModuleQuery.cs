using System;
using System.Collections.Generic;

namespace APITests.Passport.Json.Configuration
{
    internal class AppModuleQuery: IQuery
    {
        public Version MsiVersion { get; set; }
        public OfficeApp OfficeApp { get; set; }
        public OcModule OcModule { get; set; }

        public AppModuleQuery(Version msiVersion, OfficeApp officeApp, OcModule ocModule)
        {
            MsiVersion = msiVersion;
            OfficeApp = officeApp;
            OcModule = ocModule;
        }

        public AppModuleQuery() { }

        public IDictionary<string, string> AsDictionary()
        {
            var result = new Dictionary<string, string>();

            if (MsiVersion != null) result.Add("msiVersion", MsiVersion.ToString(4));

            result.Add("officeapp", Enum.GetName(typeof(OfficeApp), OfficeApp)?.ToLowerInvariant());
            result.Add("module", Enum.GetName(typeof(OcModule), OcModule)?.ToLowerInvariant());

            return result;
        }
    }
}
