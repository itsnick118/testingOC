using System;

namespace APITests.Passport.ScreenScraping
{
    internal class SearchAttribute: ICloneable
    {
        public int Id { get; set; }

        public SearchAttribute(int id)
        {
            Id = id;
        }

        public SearchAttribute() {}
        public object Clone()
        {
            var result = (SearchAttribute) MemberwiseClone();
            return result;
        }
    }
}
