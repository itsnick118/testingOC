using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Automation;
using IntegratedDriver;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using UITests.PageModel.Selectors;
using static UITests.TestHelpers;

namespace UITests.PageModel.Shared
{
    public class ItemList
    {
        private const string GetListScrollPosition =
            @"var itemlist = document.querySelector(""cdk-virtual-scroll-viewport"");
                  return (Math.ceil(itemlist.scrollHeight - itemlist.scrollTop) - itemlist.clientHeight)";

        private const string ScrollWindowScript =
            "document.querySelector(\"cdk-virtual-scroll-viewport\").scrollBy(0, document.querySelector(\"cdk-virtual-scroll-viewport\").scrollHeight);";

        private readonly IAppInstance _app;

        public ItemList(IAppInstance app)
        {
            _app = app;
        }

        public bool IsFilterIconVisible => _app.IsElementDisplayed(Oc.FilterIcon);
        public bool IsSortIconVisible => _app.IsElementDisplayed(Oc.SortIcon);
        public bool IsAddButtonVisible => _app.IsElementDisplayed(Oc.AddButton);
        public bool IsAddFolderButtonVisible => _app.IsElementDisplayed(Oc.AddFolderButton);
        public bool IsQuickSearchIconDisplayed => _app.IsElementDisplayed(Oc.QuickSearchIcon);
        public bool IsListOptionsDisplayed => _app.IsElementDisplayed(Oc.ListOptions);
        public string ListOptionsToolTip => _app.GetToolTip(_app.Driver.FindElement(Oc.ListOptions)).Text;
        public string FilterIconToolTip => _app.GetToolTip(_app.Driver.FindElement(Oc.FilterIcon)).Text;

        public void OpenFirst() => Open();

        public void OpenRandom()
        {
            var randomElementNumber = GetRandomNumber(GetCount() - 1);
            Open(randomElementNumber);
        }

        public void Open(int n = 0)
        {
            var element = GetWebElementByIndex(n);
            GetElementInView(element);
            element.Click();
            _app.WaitForListLoadComplete();
        }

        public void AddMatter() => _app.WaitAndClick(Oc.AddMatter);

        public ListOptions OpenListOptionsMenu()
        {
            _app.WaitAndClick(Oc.ListOptions);

            var menuPanel = _app.Driver.FindElement(Oc.MenuPanel);
            _app.WaitForAnimatedTransitionComplete(menuPanel);

            return new ListOptions(_app);
        }
        public string GetSavedViewNameToolTip(string viewName)
        {
            return _app.GetToolTip(_app.Driver.FindElement(Oc.ButtonByName(viewName))).Text;
        }

        public List<string> GetListOptionsMenu()
        {
            _app.WaitAndClick(Oc.ListOptions);
            var listOptionsMenuText = new List<string>();
            var menuPanel = _app.Driver.FindElement(Oc.MenuPanel);
            _app.WaitForAnimatedTransitionComplete(menuPanel);

            var listOptionsMenu = _app.Driver.FindElements(Oc.MenuOptions());
            foreach (var option in listOptionsMenu) {
                listOptionsMenuText.Add(option.Text);
            }
            RandomOverlayClick();
            return listOptionsMenuText;
        }

        public void RandomOverlayClick()
        {
            try
            {
                var overlayElement = _app.Driver.FindElement(Oc.RandomOverlayClick);
                var y = overlayElement.Size.Width;
                new Actions(_app.Driver).MoveToElement(overlayElement).MoveByOffset(0, y / 2).Click().Perform();
            }
            catch
            {
                throw new ElementNotAvailableException("Overlay Not avilable.");
            }
        }

        public void OpenAddDialog()
        {
            var addDialog = _app.Driver.FindElement(Oc.AddButton);
            _app.WaitForAnimatedTransitionComplete(addDialog);
            UserInput.LeftClick(addDialog);

            _app.WaitForLoadComplete();
        }

        public void OpenAddFolderDialog()
        {
            var addFolderDialog = _app.Driver.FindElement(Oc.AddFolderButton);
            _app.WaitForAnimatedTransitionComplete(addFolderDialog);
            UserInput.LeftClick(addFolderDialog);

            _app.WaitUntilElementAppears(Oc.SaveButton);
        }

        public Color GetFilterIconColor() => _app.GetColor(_app.Driver.FindElement(Oc.FilterIcon));

        public Color GetSortIconColor() => _app.GetColor(_app.Driver.FindElement(Oc.SortIcon));

        public bool ScrollDownIfNotAtBottom(bool finalCheck = false)
        {
            _app.WaitForLoadComplete();

            var scriptExecutor = (IJavaScriptExecutor)_app.Driver;

            scriptExecutor.ExecuteScript(ScrollWindowScript);

            _app.WaitForLoadComplete();

            if (IsAtBottomOfList() && !finalCheck)
            {
                ScrollDownIfNotAtBottom(true);
            }

            return !IsAtBottomOfList();
        }

        public ListItem GetListItemFromText(string content, bool wait = true) => GetListItemFromText<ListItem>(content, wait);

        public MatterDocumentListItem GetMatterDocumentListItemFromText(string content, bool wait = true) => GetListItemFromText<MatterDocumentListItem>(content, wait);

        public GlobalDocumentListItem GetGlobalDocumentListItemFromText(string content, bool wait = true) => GetListItemFromText<GlobalDocumentListItem>(content, wait);

        public InvoiceDocumentListItem GetInvoiceDocumentListItemFromText(string content, bool wait = true) => GetListItemFromText<InvoiceDocumentListItem>(content, wait);

        public EmailListItem GetEmailListItemFromText(string content, bool wait = true) => GetListItemFromText<EmailListItem>(content, wait);

        public MatterListItem GetMatterListItemFromText(string content, bool wait = true) => GetListItemFromText<MatterListItem>(content, wait);

        public NarrativeListItem GetNarrativeListItemFromText(string content, bool wait = true) => GetListItemFromText<NarrativeListItem>(content, wait);

        public InvoiceListItem GetInvoiceListItemFromText(string content, bool wait = true) => GetListItemFromText<InvoiceListItem>(content, wait);

        public InvoiceLineItem GetInvoiceLineItemFromText(string content, bool wait = true) => GetListItemFromText<InvoiceLineItem>(content, wait);

        public InvoiceHeaderItem GetInvoiceHeaderItemFromText(string content, bool wait = true) => GetListItemFromText<InvoiceHeaderItem>(content, wait);

        public SelectPersonListItem GetMultiSelectPersonListItemFromText(string content, bool wait = true) => GetListItemFromText<SelectPersonListItem>(content, wait);

        public ListItem GetListItemByIndex(int index) => GetListItemByIndex<ListItem>(index);

        public MatterDocumentListItem GetMatterDocumentListItemByIndex(int index) => GetListItemByIndex<MatterDocumentListItem>(index);

        public GlobalDocumentListItem GetGlobalDocumentListItemByIndex(int index) => GetListItemByIndex<GlobalDocumentListItem>(index);

        public InvoiceDocumentListItem GetInvoiceDocumentListItemByIndex(int index) => GetListItemByIndex<InvoiceDocumentListItem>(index);

        public VersionHistoryListItem GetVersionHistoryListItemByIndex(int index) => GetListItemByIndex<VersionHistoryListItem>(index);

        public EmailListItem GetEmailListItemByIndex(int index) => GetListItemByIndex<EmailListItem>(index);

        public MatterListItem GetMatterListItemByIndex(int index) => GetListItemByIndex<MatterListItem>(index);

        public PeopleListItem GetPeopleListItemByIndex(int index) => GetListItemByIndex<PeopleListItem>(index);

        public PeopleListItem GetPeopleListItemFromText(string content, bool wait = true) => GetListItemFromText<PeopleListItem>(content, wait);

        public TasksEventsListItem GetTasksEventsListItemFromText(string content, bool wait = true) => GetListItemFromText<TasksEventsListItem>(content, wait);

        public TasksEventsListItem GetTasksEventsListItemByIndex(int index) => GetListItemByIndex<TasksEventsListItem>(index);

        public NarrativeListItem GetNarrativeListItemByIndex(int index) => GetListItemByIndex<NarrativeListItem>(index);

        public InvoiceListItem GetInvoiceListItemByIndex(int index) => GetListItemByIndex<InvoiceListItem>(index);

        public InvoiceLineItem GetInvoiceLineItemByIndex(int index) => GetListItemByIndex<InvoiceLineItem>(index);

        public InvoiceHeaderItem GetInvoiceHeaderItemByIndex(int index) => GetListItemByIndex<InvoiceHeaderItem>(index);

        public SelectPersonListItem GetMultiSelectPersonListItemByIndex(int index) => GetListItemByIndex<SelectPersonListItem>(index);

        public List<MatterDocumentListItem> GetAllMatterDocumentListItems() => GetAllListItems<MatterDocumentListItem>();

        public List<GlobalDocumentListItem> GetAllGlobalDocumentListItems() => GetAllListItems<GlobalDocumentListItem>();

        public List<InvoiceDocumentListItem> GetAllInvoiceDocumentListItems() => GetAllListItems<InvoiceDocumentListItem>();

        public List<VersionHistoryListItem> GetAllVersionHistoryListItems() => GetAllListItems<VersionHistoryListItem>();

        public List<MatterListItem> GetAllMatterListItems() => GetAllListItems<MatterListItem>();

        public List<PeopleListItem> GetAllPeopleListItems() => GetAllListItems<PeopleListItem>();

        public List<EmailListItem> GetAllEmailListItems() => GetAllListItems<EmailListItem>();

        public List<NarrativeListItem> GetAllNarrativeListItems() => GetAllListItems<NarrativeListItem>();

        public List<TasksEventsListItem> GetAllTasksEventsListItems() => GetAllListItems<TasksEventsListItem>();

        public List<InvoiceListItem> GetAllInvoiceListItems() => GetAllListItems<InvoiceListItem>();

        public List<SelectPersonListItem> GetAllMultiSelectPersonListItem() => GetAllListItems<SelectPersonListItem>();

        public IReadOnlyCollection<ListItem> GetAllListItems()
        {
            var listItems = GetAllWebElementListItems();
            return listItems.Select(item => new ListItem(_app, item)).ToList();
        }

        public IReadOnlyCollection<IWebElement> GetAllWebElementListItems()
        {
            _app.WaitForLoadComplete();

            return _app.Driver.FindElements(Oc.ListItems);
        }

        public IWebElement GetWebElementByIndex(int index)
        {
            _app.WaitForListLoadComplete();
            return _app.Driver.FindElement(Oc.NthListItem(index));
        }

        public int GetCount() => GetAllWebElementListItems().Count;

        public int GetFooterCount()
        {
            _app.WaitForListLoadComplete();
            var footer = _app.Driver.FindElement(Oc.ListCount).Text;
            var footerArray = footer.Split(' ');
            return footer.Contains("of") ? Convert.ToInt32(footerArray[2]) : Convert.ToInt32(footerArray[0]);
        }

        private void GetElementInView(IWebElement element)
        {
            var driver = (IJavaScriptExecutor)_app.Driver;
            driver.ExecuteScript("arguments[0].scrollIntoView(true);", element);
        }

        private IWebElement GetListItem(string content, bool wait = true)
        {
            IWebElement result = null;

            if (!wait) _app.SetShortImplicitWait();

            _app.WaitForLoadComplete();

            do
            {
                try
                {
                    result = _app.Driver.FindElement(Oc.ListItemByContent(content));
                    GetElementInView(result);
                }
                catch
                {
                    // swallowing
                }
            } while (result == null && ScrollDownIfNotAtBottom());

            _app.SetLongImplicitWait();
            return result;
        }

        private bool IsAtBottomOfList()
        {
            var driver = (IJavaScriptExecutor)_app.Driver;
            _app.WaitForLoadComplete();

            var scrollPosition = driver.ExecuteScript(GetListScrollPosition);

            return (long)scrollPosition <= 1;
        }

        private T GetListItemFromText<T>(string content, bool wait = true) where T : ListItem
        {
            var element = GetListItem(content, wait);
            return element == null ? null : (T)Activator.CreateInstance(typeof(T), _app, element);
        }

        private T GetListItemByIndex<T>(int index) where T : ListItem
        {
            var element = GetWebElementByIndex(index);
            GetElementInView(element);
            return (T)Activator.CreateInstance(typeof(T), _app, element);
        }

        private List<T> GetAllListItems<T>() where T : ListItem
        {
            var items = new List<T>();
            foreach (var element in GetAllWebElementListItems())
            {
                items.Add((T)Activator.CreateInstance(typeof(T), _app, element));
            }
            return items;
        }

        public string ListEmptyMessage()
        {
            return _app.Driver.FindElement(Oc.ListEmptyMessage).Text;
        }
    }
}
