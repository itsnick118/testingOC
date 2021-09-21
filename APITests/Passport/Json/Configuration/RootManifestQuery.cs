using System;
using System.Collections.Generic;

namespace APITests.Passport.Json.Configuration
{
    internal class RootManifestQuery : IQuery
    {
        public Version MsiVersion { get; set; }

        public RootManifestQuery(Version msiVersion)
        {
            MsiVersion = msiVersion;
        }

        public RootManifestQuery() { }

        public IDictionary<string, string> AsDictionary()
        {
            var result = new Dictionary<string, string>();

            if (MsiVersion != null) result.Add("msiVersion", MsiVersion.ToString(4));

            return result;
        }
    }
}
