using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using UITests.PageModel.Selectors;
using UITests.PageModel.Shared;

namespace UITests.PageModel
{
    public class SelectPersonDialog
    {
        private readonly IAppInstance _app;

        public SelectPersonDialog(IAppInstance app)
        {
            _app = app;
            ItemList = new ItemList(app);
        }

        public ItemList ItemList { get; }

        public void Close()
        {
            _app.JustClick(Oc.DialogClose);
            _app.WaitUntilElementDisappears(Oc.MultiselectWindow);
        }

        public void Done()
        {
            _app.SetLongImplicitWait();
            _app.JustClick(Oc.DialogDone);
            _app.WaitUntilElementAppears(Oc.SaveButton);
        }

        public string[] GetSelectedPersons()
        {
            var items = GetAddedItems();
            return items.Select(item => item.Text).ToArray();
        }

        public string GetValue()
        {
            return string.Join(",", GetSelectedPersons());
        }

        public void Remove(string assignees)
        {
            var index = GetIndexOfAddedItemFromMultiSelectDialog(assignees);
            _app.JustClick(Oc.ItemInMultiSelectDialogListByIndex(index));
        }

        public void RemoveAll()
        {
            for (var i = GetAddedItems().Count - 1; i >= 0; i--)
            {
                _app.JustClick(Oc.ItemInMultiSelectDialogListByIndex(i));
                _app.SetLongImplicitWait();
            }
        }
        private IReadOnlyCollection<IWebElement> GetAddedItems()
        {
            return _app.Driver.FindElements(Oc.ItemsInMultiSelectWindow);
        }

        private int GetIndexOfAddedItemFromMultiSelectDialog(string item)
        {
            var addedItem = GetSelectedPersons();
            return Array.IndexOf(addedItem, item);
        }
    }
}
