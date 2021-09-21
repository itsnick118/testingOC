using System;
using System.Drawing;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using UITests.PageModel.Selectors;

namespace UITests.PageModel.Shared
{
    public class ListItem
    {
        public readonly IWebElement PrimaryElement;
        protected readonly IAppInstance App;
        protected readonly IWebElement Element;

        public ListItem(IAppInstance app, IWebElement element)
        {
            App = app;
            Element = element;
            PrimaryElement = GetListItemProperty("primary-text");
        }

        public IWebElement ActionItems => GetListItemProperty("actions-items");
        public string DeleteButtonTooltip => App.GetToolTip(App.Driver.FindElement(Oc.DeleteButton)).Text;
        public bool HasActionItems => App.IsElementDisplayed(Oc.ActionItems);
        public IWebElement Meta2Element => GetListItemProperty("meta2");
        public IWebElement Meta3Element => GetListItemProperty("meta3");
        public IWebElement SecondaryElement => GetListItemProperty("secondary-text");
        public IWebElement TernaryElement => GetListItemProperty("ternary-text");
        protected string Meta2 => Meta2Element.Text;
        protected string Meta3 => Meta3Element.Text;
        protected string PrimaryText => PrimaryElement.Text.Trim();
        protected string SecondaryText => SecondaryElement.Text;
        protected string TertiaryText => TernaryElement.Text;

        protected void Edit()
        {
            App.JustClick(Oc.EditButton, Element);
            WaitForDialog();
        }

        public Color GetDeleteButtonColor() => HoverAndGetColor(Element.FindElement(Oc.DeleteButton));

        public IWebElement GetListItemProperty(string className) =>
            Element.FindElement(Oc.ByClassName(className));

        public string GetParentClass(IWebElement element) => element.FindElement(Oc.Parent).GetAttribute("class");

        public Color GetToolTipBackground(IWebElement element) =>
            App.GetColor(App.GetToolTip(element), "background-color");

        public Color GetToolTipFontColor(IWebElement element) => App.GetColor(App.GetToolTip(element));

        public string GetToolTipShape(IWebElement element) => App.GetToolTip(element).GetCssValue("text-overflow");

        public bool IsDeleteButtonVisible() => Element.FindElement(Oc.DeleteButton).Displayed;

        public bool IsEditButtonVisible() => Element.FindElement(Oc.EditButton).Displayed;

        public void NavigateToSummary()
        {
            HoverMouseOverListItem(Element);
            App.JustClick(Oc.SummaryIcon, Element);

            App.WaitForLoadComplete();

            var wait = new WebDriverWait(App.Driver, TimeSpan.FromSeconds(30));
            wait.Until(condition =>
            {
                try
                {
                    Element.FindElement(Oc.SummaryIcon);
                    return false;
                }
                catch
                {
                    return true;
                }
            });

            App.WaitForLoadComplete();
        }

        public void Open()
        {
            Element.Click();

            App.WaitForLoadComplete();
        }

        protected IDialog Delete(By selector)
        {
            HoverMouseOverListItem(Element);

            App.JustClick(selector, Element);
            App.WaitUntilElementAppears(Oc.DialogHeader);

            return new Dialog(App, Element);
        }

        protected Color HoverAndGetColor(IWebElement element)
        {
            App.HoverMouseOverElement(element);
            return App.GetColor(element);
        }

        protected void HoverMouseOverListItem(IWebElement listItem) => App.HoverMouseOverElement(listItem);

        protected void Select() => App.ClickAndWait(Oc.CheckBox, Element);

        private void WaitForDialog()
        {
            App.WaitFor(condition => App.IsElementDisplayed(Oc.DialogHeader));
        }
        protected void QuickFile()
        {
            App.ClickAndWait(Oc.ListItemQuickFile, Element);
            App.WaitForQueueComplete();
        }
    }
}
