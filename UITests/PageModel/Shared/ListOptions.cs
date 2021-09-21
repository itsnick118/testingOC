using System.Windows.Automation;

namespace UITests.PageModel.Shared
{
    public class ListOptions
    {
        private readonly IAppInstance _app;

        public ListOptions(IAppInstance app)
        {
            _app = app;
        }

        public void OpenCreateListFilterDialog()
        {
            _app.ClickAndWait(Selectors.Oc.FilterList);
        }

        public void RestoreDefaults()
        {
            _app.JustClick(Selectors.Oc.RestoreDefaults);
            _app.WaitForListLoadComplete();
        }

        public void SaveCurrentView()
        {
            _app.ClickAndWait(Selectors.Oc.FilterSaveView);
        }

        public void RemoveSavedView(string viewName)
        {
            _app.HoverMouseOverElement(Selectors.Oc.FilterSavedViews);
            _app.ClickAndWait(Selectors.Oc.RemoveSavedView(viewName));
        }

        public void ApplySavedView(string viewName)
        {
            _app.HoverMouseOverElement(Selectors.Oc.FilterSavedViews);
            _app.JustClick(Selectors.Oc.ButtonByName(viewName));
            _app.WaitForListLoadComplete();
        }

        public void HoverMouseOnSavedViewMenuOption()
        {
            if (_app.IsElementDisplayed(Selectors.Oc.FilterSavedViews))
            {
                _app.HoverMouseOverElement(Selectors.Oc.FilterSavedViews);
            }
            else
            {
                throw new ElementNotAvailableException("Saved Views Button Not Visible/Available");
            }
        }

        public void HoverMouseOnSavedViewByName(string viewName)
        {
            _app.HoverMouseOverElement(Selectors.Oc.FilterSavedViews);
            if (_app.IsElementDisplayed(Selectors.Oc.ButtonByName(viewName)))
            {
                _app.HoverMouseOverElement(Selectors.Oc.ButtonByName(viewName));
                _app.IsElementDisplayed(Selectors.Oc.RemoveSavedView(viewName));
            }
            else
            {
                throw new ElementNotAvailableException("Element Not Visible/Available");
            }
        }

        public void SetCurrentViewAsDefault()
        {
            _app.JustClick(Selectors.Oc.FilterSetViewAsDefault);
        }

        public void ClearUserDefault()
        {
            _app.JustClick(Selectors.Oc.FilterClearUserDefault);
            _app.WaitForListLoadComplete();
        }

        public bool IsClearUserDefaultDisplayed() => _app.IsElementDisplayed(Selectors.Oc.FilterClearUserDefault);
    }
}
