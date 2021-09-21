using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using NUnit.Framework;

namespace APITests
{
    public class PassportScreen
    {
        private readonly string _htmlString;

        public PassportScreen(string htmlString)
        {
            _htmlString = htmlString;
        }

        public IList<string> GetTableHeader()
        {
            var begin = _htmlString.IndexOf("<thead>", StringComparison.Ordinal) + "<thead>".Length;
            var end = _htmlString.IndexOf("</thead>", StringComparison.Ordinal);

            if (begin == -1 || end == -1)
            {
                Assert.Fail("Response to be evaluated does not contain a table header.");
            }

            const string regex = @"<th[^>]*>([^<]*)</th>";
            var returnCollection = new List<string>();

            var matches = Regex.Matches(_htmlString.Substring(begin, end - begin), regex);
            foreach (Match match in matches)
            {
                returnCollection.Add(match.Groups[1].Value.Replace("\t", "").Replace("\n", ""));
            }
            return returnCollection;
        }
    }
}
