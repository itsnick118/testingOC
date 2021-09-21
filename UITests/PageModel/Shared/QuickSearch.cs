using UITests.PageModel.Selectors;
using UITests.PageModel.Shared.InputControls;

namespace UITests.PageModel.Shared
{
    public class QuickSearch
    {
        private readonly IAppInstance _app;

        private readonly InputField _searchInputField;

        public QuickSearch(IAppInstance app)
        {
            _app = app;
            _searchInputField = new InputField(_app, Oc.SearchInput);
        }

        public void SearchBy(string text)
        {
            if (!_app.IsElementDisplayed(Oc.SearchInput))
            {
                _app.JustClick(Oc.QuickSearchIcon);
            }

            _app.WaitUntilElementAppears(Oc.SearchInput);

            _searchInputField.Set(text);

            _app.WaitForListLoadComplete();
        }

        public void Close()
        {
            if (_app.IsElementDisplayed(Oc.SearchInput))
            {
                _app.JustClick(Oc.QuickSearchIcon);
            }

            _app.WaitUntilElementDisappears(Oc.SearchInput);

            _app.WaitForListLoadComplete();
        }
    }
}