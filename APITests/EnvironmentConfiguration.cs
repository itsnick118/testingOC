using System;
using System.Collections.Specialized;
using System.Configuration;
using System.Net.Http.Headers;
using System.Text;

namespace APITests
{
    public class EnvironmentConfiguration
    {
        public string BaseUrl { get; protected set; }
        protected string StandardUser { get; set; }
        protected string StandardUserPassword { get; set; }
        protected string ElevatedUser { get; set; }
        protected string ElevatedUserPassword { get; set; }

        public EnvironmentConfiguration(string environmentLabel)
        {
            ParseConfigForEnvironment(environmentLabel);
        }

        public AuthenticationHeaderValue GetStandardUserHeaders()
        {
            var credentials =
                Convert.ToBase64String(
                    Encoding.ASCII.GetBytes(StandardUser + ":" + StandardUserPassword));
            return new AuthenticationHeaderValue("Basic", credentials);
        }
        public AuthenticationHeaderValue GetElevatedUserHeaders()
        {
            var credentials =
                Convert.ToBase64String(
                    Encoding.ASCII.GetBytes(ElevatedUser + ":" + ElevatedUserPassword));
            return new AuthenticationHeaderValue("Basic", credentials);
        }

        private void ParseConfigForEnvironment(string environmentLabel)
        {
            var section = ConfigurationManager.GetSection(environmentLabel) as NameValueCollection;
            if (section == null) return;

            BaseUrl = section["baseUrl"];
            if (!BaseUrl.EndsWith("/")) BaseUrl += '/';

            StandardUser = section["standardUser"];
            StandardUserPassword = section["standardPass"];
            ElevatedUser = section["elevatedUser"];
            ElevatedUserPassword = section["elevatedPass"];
        }
    }
}
