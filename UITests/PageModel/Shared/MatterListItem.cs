using OpenQA.Selenium;
using System;
using UITests.PageModel.Selectors;
using static UITests.TestHelpers;

namespace UITests.PageModel.Shared
{
    public class MatterListItem : ListItem
    {
        public string Name { get; }
        public string Number { get; }
        public string PrimaryInternalContact { get; }
        public string Status { get; }
        public DateTime? StatusDate { get; }
        public string SpendToDate { get; }
        public bool HasQuickFileIcon { get; }
        public IWebElement DropPoint { get; }

        private readonly IAppInstance _app;

        public MatterListItem(IAppInstance app, IWebElement element) : base(app, element)
        {
            _app = app;

            Name = PrimaryText;
            DropPoint = element;

            var tertiary = TertiaryText.Split(new[] { " ● " }, StringSplitOptions.None);
            var status = tertiary[2].Split(new[] { " - " }, StringSplitOptions.None);
            Number = tertiary[0];
            PrimaryInternalContact = tertiary[1];
            Status = tertiary[2].Contains("-") ? status[0] : tertiary[2];
            StatusDate = tertiary[2].Contains("-") ? ParseDateTime(status[1]) : new DateTime?();
            SpendToDate = tertiary.Length == 4 ? tertiary[3] : null;
            HasQuickFileIcon = app.IsElementDisplayed(Oc.ListItemQuickFile, element);
        }

        public void ClearFavorite() => _app.ClickAndWait(Oc.FavoriteToggle, Element);

        public void SetAsFavorite()
        {
            var favoriteToggle = Element.FindElement(Oc.FavoriteToggle);
            if (!favoriteToggle.GetAttribute("class").Contains("FavoriteStarFill"))
            {
                favoriteToggle.Click();
            }

            App.WaitForLoadComplete();
        }

        public override string ToString()
        {
            return $"{nameof(Name)}:{Name}, {nameof(Number)}:{Number}, " +
                $"{nameof(PrimaryInternalContact)}:{PrimaryInternalContact}, {nameof(Status)}:{Status}, " +
                $"{nameof(StatusDate)}:{StatusDate}, {nameof(SpendToDate)}:{SpendToDate}";
        }

        public void AccessMatter()
        {
            _app.ClickAndWait(Oc.AccessButton);
        }
    }
}