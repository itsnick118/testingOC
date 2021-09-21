using IntegratedDriver;
using OpenQA.Selenium;
using UITests.PageModel.Selectors;
using static UITests.TestHelpers;

namespace UITests.PageModel.Shared.InputControls
{
    public class AutoComplete : InputControl, IIndexedInputControl
    {
        private readonly IAppInstance _app;
        private readonly bool _usePipes;

        public AutoComplete(IAppInstance app, string name, string className, bool usePipes = true, int panelNumber = 0) : base(app)
        {
            _app = app;
            _usePipes = usePipes;
            Name = name;
            ClassName = className;
            PanelNumber = panelNumber;
        }

        public override void SelectPersonDialog()
        {
            _app.JustClick(Oc.MultiSelectDialog(ClassName));
            _app.WaitUntilElementAppears(Oc.MultiselectListItem);
        }

        public override string GetReadOnlyValue() => GetValue();

        public override string GetValue()
        {
            OpenPanel();
            var chipElements = _app.Driver.FindElements(Oc.MultiSelectInputChips(ClassName));
            var chipsValues = new string[chipElements.Count];
            for (var i = 0; i < chipElements.Count; i++)
            {
                chipsValues[i] = chipElements[i].Text;
            }

            return string.Join(",", chipsValues);
        }

        public override string Set(string value)
        {
            _app.WaitForLoadComplete();
            OpenPanel();
            var input = _app.Driver.FindElement(Oc.MultiSelectInputByClass(ClassName));
            input.Click();
            UserInput.Type(input, value);
            var searchValue = _usePipes ? AddPipesToSearchItem(value) : value;
            _app.WaitAndClick(Oc.MultiSelectItemByName(ClassName, searchValue));
            return value;
        }

        public override string SetByIndex(int index, bool clearInput = false)
        {
            _app.WaitForLoadComplete();
            OpenPanel();
            if (clearInput) Clear();
            var input = _app.Driver.FindElement(Oc.MultiSelectInputByClass(ClassName));
            input.Click();
            UserInput.Type(input, "e");
            try
            {
                _app.WaitUntilElementAppears(Oc.MultiSelectItemByIndex(ClassName, index));
                _app.WaitForLoadComplete();
                var itemToSelect = _app.Driver.FindElement(Oc.MultiSelectItemByIndex(ClassName, index));
                var itemText = GetItemText(itemToSelect);
                itemToSelect.Click();
                return itemText;
            }
            catch (WebDriverTimeoutException)
            {
                return null;
            }
        }

        public override string SetValueOtherthan(string value, bool clearInput = false)
        {
            _app.WaitForLoadComplete();
            OpenPanel();
            if (clearInput) Clear();
            var input = _app.Driver.FindElement(Oc.MultiSelectInputByClass(ClassName));
            input.Click();
            UserInput.Type(input, " ");
            try
            {
                var listItemsCount = _app.Driver.FindElements(Oc.MultiSelectItems()).Count;
                var flag = true;
                while (flag)
                {
                    var randomIndex = GetRandomNumber(listItemsCount - 1);
                    var itemToSelect = _app.Driver.FindElement(Oc.MultiSelectItemByIndex(ClassName, randomIndex));
                    var itemText = GetItemText(itemToSelect);
                    if (!itemText.Equals(value))
                    {
                        flag = false;
                        itemToSelect.Click();
                    }
                }
                return value;
            }
            catch (WebDriverTimeoutException)
            {
                return null;
            }
        }

        private static string AddPipesToSearchItem(string value)
        {
            // Let's look forward to this not being necessary
            return value.Replace(" ", " | ");
        }

        public override void Clear()
        {
            var clearChipElements = _app.Driver.FindElements(Oc.RemoveInputChips(ClassName));
            for (var i = clearChipElements.Count - 1; i >= 0; i--)
            {
                _app.JustClick(Oc.MatChipListCrossIcon(ClassName, i));
            }
        }

        private string GetItemText(IWebElement itemToSelect)
        {
            /*TODO : Need to remove below if checks as in person autocomplete
             we can see double pipe ( | | ), and all other places it's single pipe( | )
             between the names.
            Mingle defect: http://mingle/projects/growth/cards/22938*/
            var itemToSelectValues = itemToSelect.Text.Split('|');
            if (itemToSelectValues[1].Trim().Equals(""))
                itemToSelectValues[1] = itemToSelectValues[2];
            return $"{itemToSelectValues[0].Trim()} {itemToSelectValues[1].Trim()}";
        }
    }
}
