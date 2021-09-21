using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace APITests.Passport.ScreenScraping
{
    internal class DynamicSearch: ICloneable
    {
        public IList<DynamicSearchCriterion> SearchCriteria { get; set; }

        public IDictionary<string, string> AsDictionary()
        {
            var dict = new Dictionary<string, string>
            {
                {"searchInput.searchCriteria.dynamicSearchCriteria[]", string.Empty}
            };

            var index = 0;
            foreach (var searchCriterion in SearchCriteria)
            {
                dict.Add(CriteriaKey(index, ".leftValue"), searchCriterion.LeftValue);
                dict.Add(CriteriaKey(index, ".comparisonType"), searchCriterion.ComparisonType);

                if (searchCriterion.Attributes.Count <= 0) continue;

                dict.Add(CriteriaKey(index, ".attributes[]"), string.Empty);

                var attributeIndex = 0;
                foreach (var attribute in searchCriterion.Attributes)
                {
                    dict.Add(CriteriaKey(index, attributeIndex, ".id"), attribute.Id.ToString());

                    attributeIndex++;
                }

                index++;
            }

            return dict;
        }

        private static string CriteriaKey(int index, string suffix)
        {
            return $"searchInput.searchCriteria.dynamicSearchCriteria[{index}]{suffix}";
        }

        private static string CriteriaKey(int index, int attributeIndex, string attributeSuffix)
        {
            return $"{CriteriaKey(index, ".attributes")}[{attributeIndex}]{attributeSuffix}";
        }

        public object Clone()
        {
            var result = MemberwiseClone() as DynamicSearch;
            if (result != null)
            {
                result.SearchCriteria = new List<DynamicSearchCriterion>();
                foreach (var searchCriteria in SearchCriteria)
                {
                    result.SearchCriteria.Add(searchCriteria.Clone() as DynamicSearchCriterion);
                }
            }

            Debug.Assert(result != null, nameof(result) + " != null");
            return result;
        }
    }
}