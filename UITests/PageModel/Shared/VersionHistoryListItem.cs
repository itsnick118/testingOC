using OpenQA.Selenium;
using System;
using static UITests.TestHelpers;

namespace UITests.PageModel.Shared
{
    public class VersionHistoryListItem : BaseDocumentListItem
    {
        public double Version { get; }
        public string CreatedBy { get; }
        public string Comments { get; }
        public string Size { get; }
        public DateTime UploadedAt { get; }

        public VersionHistoryListItem(IAppInstance app, IWebElement element) : base(app, element)
        {
            Version = Convert.ToDouble(PrimaryText);
            CreatedBy = SecondaryText;
            Comments = TertiaryText;
            Size = Meta2;
            UploadedAt = ParseDateTime(Meta3);
        }

        public override string ToString()
        {
            return $"{nameof(Version)}:{Version},{nameof(CreatedBy)}:{CreatedBy},{nameof(Comments)}:{Comments}";
        }
    }
}