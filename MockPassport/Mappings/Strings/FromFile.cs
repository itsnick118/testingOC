using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace MockPassport.Mappings.Strings
{
    public class FromFile
    {
        public static IDictionary<string, string> GetHeaders(IEnvironment environment, string filename)
        {
            var dictionary = new Dictionary<string, string>();

            var filePath = Path.Combine(environment.BaseFilePath.FullName, "Headers", filename);

            if (!File.Exists(filePath))
            {
                return dictionary;
            }

            foreach (var line in File.ReadAllLines(filePath))
            {
                var splitPoint = line.IndexOf(':');
                dictionary[line.Substring(0, splitPoint)] = line.Substring(splitPoint + 1).Trim();
            }

            return dictionary;
        }

        public static string GetMetadataHeader(IEnvironment environment, string filename)
        {
            var headers = GetHeaders(environment, filename);
            return headers.ContainsKey("metadata") ? headers["metadata"] : string.Empty;
        }

        public static IDictionary<string, string> GetFirstNinePageHeaders(IEnvironment environment, string filename)
        {
            var currentPageNumber = new Regex(@",""searchInput\.pageInfo\.currentPageNumber"":""\d""");
            var totalRecords = new Regex(@",""searchInput\.pageInfo\.totalRecords"":""\d+""");
            var headers = GetHeaders(environment, filename);
            if (headers.ContainsKey("metadata"))
            {
                headers["metadata"] = currentPageNumber.Replace(headers["metadata"], string.Empty);
                headers["metadata"] = totalRecords.Replace(headers["metadata"],
                    @",""searchInput\.pageInfo\.totalRecords"":""500""");
            }
            
            return headers;
        }

        public static IDictionary<string, string> GetTenthPageHeaders(IEnvironment environment, string filename)
        {
            var currentPageNumber = new Regex(@",""searchInput\.pageInfo\.currentPageNumber"":""\d""");
            var totalRecords = new Regex(@",""searchInput\.pageInfo\.totalRecords"":""\d+""");
            var headers = GetHeaders(environment, filename);
            if (headers.ContainsKey("metadata"))
            {
                headers["metadata"] = currentPageNumber.Replace(headers["metadata"],
                    @",""searchInput\.pageInfo\.currentPageNumber"":""10""");
                headers["metadata"] = totalRecords.Replace(headers["metadata"],
                    @",""searchInput\.pageInfo\.totalRecords"":""500""");
            }

            return headers;
        }

        public static IDictionary<string, string> GetFirstNineteenPageHeaders(IEnvironment environment, string filename)
        {
            var currentPageNumber = new Regex(@",""searchInput\.pageInfo\.currentPageNumber"":""\d""|,""searchInput\.pageInfo\.currentPageNumber"":""1\d""");
            var totalRecords = new Regex(@",""searchInput\.pageInfo\.totalRecords"":""\d+""");
            var headers = GetHeaders(environment, filename);
            if (headers.ContainsKey("metadata"))
            {
                headers["metadata"] = currentPageNumber.Replace(headers["metadata"], string.Empty);
                headers["metadata"] = totalRecords.Replace(headers["metadata"],
                    @",""searchInput\.pageInfo\.totalRecords"":""500""");
            }

            return headers;
        }

        public static IDictionary<string, string> GetTwentiethPageHeaders(IEnvironment environment, string filename)
        {
            var currentPageNumber = new Regex(@",""searchInput\.pageInfo\.currentPageNumber"":""\d""|,""searchInput\.pageInfo\.currentPageNumber"":""1\d""");
            var totalRecords = new Regex(@",""searchInput\.pageInfo\.totalRecords"":""\d+""");
            var headers = GetHeaders(environment, filename);
            if (headers.ContainsKey("metadata"))
            {
                headers["metadata"] = currentPageNumber.Replace(headers["metadata"],
                    @",""searchInput\.pageInfo\.currentPageNumber"":""10""");
                headers["metadata"] = totalRecords.Replace(headers["metadata"],
                    @",""searchInput\.pageInfo\.totalRecords"":""500""");
            }

            return headers;
        }

        public static string GetBody(IEnvironment environment, string filename)
        {
            var filePath = Path.Combine(environment.BaseFilePath.FullName, "Responses", filename);

            return File.Exists(filePath)
                ? File.ReadAllText(filePath)
                : "This endpoint is not supported by the source system.";
        }
    }
}
