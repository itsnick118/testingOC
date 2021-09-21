using System;
using System.Collections.Generic;

namespace APITests.Passport.ScreenScraping
{
    internal class DynamicSearchCriterion: ICloneable
    {
        public string ComparisonType { get; set; }
        public IList<SearchAttribute> Attributes { get; set; }
        public string LeftValue { get; set; }

        public DynamicSearchCriterion(string comparisonType, string leftValue, IList<SearchAttribute> attributes)
        {
            ComparisonType = comparisonType;
            LeftValue = leftValue;
            Attributes = attributes;
        }
        public DynamicSearchCriterion(string comparisonType, string leftValue, SearchAttribute attribute)
        {
            ComparisonType = comparisonType;
            LeftValue = leftValue;
            Attributes = new List<SearchAttribute> {attribute};
        }

        public DynamicSearchCriterion() { }

        public object Clone()
        {
            var result = (DynamicSearchCriterion) MemberwiseClone();

            result.Attributes = new List<SearchAttribute>();

            foreach (var attribute in Attributes)
            {
                result.Attributes.Add(attribute.Clone() as SearchAttribute);
            }

            return result;
        }
    }
}