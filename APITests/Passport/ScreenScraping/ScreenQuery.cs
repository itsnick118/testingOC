using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace APITests.Passport.ScreenScraping
{
    internal class ScreenQuery: ICloneable, IQuery
    {
        public string ScreenName { get; set; }
        public string SearchKeywords { get; set; }
        public int? CurrentPageNumber { get; set; }
        public int? CurrentPageSize { get; set; }
        public string PageType { get; set; }
        public string CssClasses { get; set; }
        public bool? LoadImmediately { get; set; }
        public bool? NoCache { get; set; }
        public string DocumentTitle { get; set; }
        public int? FalseParm { get; set; }
        public int? PageOffset { get; set; }
        public DynamicSearch DynamicSearch { get; set; }

        public IDictionary<string, string> AsDictionary()
        {
            var dict = new Dictionary<string, string>();

            if (ScreenName != null) dict.Add("screenName", ScreenName);
            if (SearchKeywords != null) dict.Add("search-keywords", SearchKeywords);
            if (PageType != null) dict.Add("pageType", PageType);
            if (CssClasses != null) dict.Add("cssClasses", CssClasses);
            if (DocumentTitle != null) dict.Add("documentTitle", DocumentTitle);

            if (CurrentPageNumber.HasValue)
            {
                dict.Add("searchInput.pageInfo.currentPageNumber", CurrentPageNumber.Value.ToString());
            }
            if (CurrentPageSize.HasValue)
            {
                dict.Add("searchInput.pageInfo.currentPageSize", CurrentPageSize.Value.ToString());
            }
            if (LoadImmediately.HasValue)
            {
                dict.Add("loadImmediately", LoadImmediately.Value.ToString().ToLowerInvariant());
            }
            if (NoCache.HasValue)
            {
                dict.Add("nocache", NoCache.Value.ToString().ToLowerInvariant());
            }
            if (FalseParm.HasValue)
            {
                dict.Add("falseParm", FalseParm.Value.ToString());
            }
            if (PageOffset.HasValue)
            {
                dict.Add("pageOffset", PageOffset.Value.ToString());
            }

            if (DynamicSearch != null)
            {
                foreach (var keyValuePair in DynamicSearch.AsDictionary())
                {
                    dict.Add(keyValuePair.Key, keyValuePair.Value);
                }
            }

            return dict;
        }

        public object Clone()
        {
            var result = MemberwiseClone() as ScreenQuery;
            if (result != null)
            {
                result.DynamicSearch = DynamicSearch?.Clone() as DynamicSearch;
            }

            Debug.Assert(result != null, nameof(result) + " != null");
            return result;
        }
    }
}
