using UITests.PageModel.Selectors;

namespace UITests.PageModel.Shared.InputControls
{
    public class Dropdown : InputControl, IIndexedInputControl
    {
        private readonly IAppInstance _app;

        public Dropdown(IAppInstance app, string name, string className, int panelNumber = 0) : base(app)
        {
            _app = app;
            Name = name;
            ClassName = className;
            PanelNumber = panelNumber;
        }

        public override string Set(string value)
        {
            OpenPanel();
            _app.JustClick(Oc.DropDownInputByClass(ClassName));
            _app.JustClick(Oc.DropDownItemByText(value));
            return value;
        }

        public override string SetByIndex(int index, bool clearInput = false)
        {
            OpenPanel();
            _app.JustClick(Oc.DropDownInputByClass(ClassName));
            _app.WaitForLoadComplete();
            _app.WaitUntilElementAppears(Oc.DropDownItemByIndex(index));
            var itemToSelect = _app.Driver.FindElement(Oc.DropDownItemByIndex(index));
            var itemText = itemToSelect.Text;
            itemToSelect.Click();
            _app.WaitForAnimatedTransitionComplete(itemToSelect);
            return itemText;
        }

        public override string GetValue()
        {
            return _app.Driver.FindElement(Oc.DropDownInputByClass(ClassName)).Text.Trim();
        }
    }
}