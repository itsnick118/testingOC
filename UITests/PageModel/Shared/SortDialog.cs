using System.Linq;
using UITests.Extensions;
using UITests.PageModel.Selectors;

namespace UITests.PageModel.Shared
{
    public class SortDialog : ISortDialog
    {
        private readonly IAppInstance _app;

        public SortDialog(IAppInstance app)
        {
            _app = app;
        }

        public void OpenSortDialog() => _app.ClickAndWait(Oc.SortIcon);

        public void CloseSortDialog() => _app.ClickAndWait(Oc.Backdrop);

        public string[] GetSortOptions()
        {
            var sortOptions = _app.Driver.FindElements(Oc.SortOptions);
            return sortOptions.Select(btn => btn.Text).ToArray();
        }
        public void RestoreSortDefaults()
        {
            OpenSortDialog();
            _app.JustClick(Oc.RestoreSortDefaults);
            _app.WaitForListLoadComplete();
        }

        public bool IsSortRestoreDefaultPresent() => _app.IsElementDisplayed(Oc.RestoreSortDefaults);

        public void Sort(string option, SortOrder sortOrder)
        {
            OpenSortDialog();

            var currentSortDirection = GetCurrentSortDirection(option);

            switch (currentSortDirection)
            {
                default:
                    ApplySortByOption(option);
                    if (sortOrder == SortOrder.Descending)
                    {
                        OpenSortDialog();
                        ApplySortByOption(option);
                    }

                    break;
                case SortOrder.Ascending:
                case SortOrder.Descending:
                    if (currentSortDirection != sortOrder)
                    {
                        ApplySortByOption(option);
                    }
                    else
                    {
                        CloseSortDialog();
                    }

                    break;
            }
        }

        private SortOrder? GetCurrentSortDirection(string option)
        {
            var iconClassAttribute = _app.Driver.FindElement(Oc.SortIconByOption(option)).GetAttribute("class");

            if (iconClassAttribute.Contains(SortOrder.Ascending.GetDescription()))
            {
                return SortOrder.Ascending;
            }

            if (iconClassAttribute.Contains(SortOrder.Descending.GetDescription()))
            {
                return SortOrder.Descending;
            }

            return null;
        }

        private void ApplySortByOption(string option)
        {
            _app.JustClick(Oc.SortButtonByOption(option));
            _app.WaitForListLoadComplete();
        }
    }
}
