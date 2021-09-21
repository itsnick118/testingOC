using System.Collections.Generic;
using OpenQA.Selenium;
using UITests.PageModel.Selectors;

namespace UITests.PageModel.Shared
{
    public class BreadcrumbsControl
    {
        private readonly IAppInstance _app;

        public BreadcrumbsControl(IAppInstance app)
        {
            _app = app;
        }

        public string GetCurrentPath()
        {
            // returns the path in a format "/folder1/folder2"
            var breadcrumbs = FindBreadcrumbsFolders();

            var breadcrumbsPath = "";
            foreach (var crumb in breadcrumbs)
            {
                breadcrumbsPath += $"/{crumb.Text}";
            }

            return breadcrumbsPath;
        }

        public void NavigateToTheRoot() => _app.ClickAndWait(Oc.BreadcrumbsRootFolder);

        public void NavigateToFolder(string folderName)
        {
            var breadcrumbs = FindBreadcrumbsFolders();
            foreach (var crumb in breadcrumbs)
            {
                if (!crumb.Text.Equals(folderName)) continue;
                crumb.Click();
                _app.WaitForLoadComplete();
                return;
            }

            throw new NoSuchElementException($"There is no {folderName} in breadcrumbs control");
        }

        private IEnumerable<IWebElement> FindBreadcrumbsFolders() => _app.Driver.FindElements(Oc.BreadcrumbsFolders);
    }
}