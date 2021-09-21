using OpenQA.Selenium;

// ReSharper disable InconsistentNaming

namespace UITests.PageModel.Selectors
{
    public static class Passport
    {
        // Layout regions
        public static By PageLoading = By.Id("pageLoading");

        // List Items
        public static By LinkItems = By.TagName("a");

        // Frame Items
        public static By PassportFrame = By.Id("iFrameResizer0");
        public static By MiniNavFrame = By.XPath("//frame[@src='whskin_frmset01.htm']");
        public static By MiniBarFrame = By.XPath("//frame[@id='minibar_navpane']");
        public static By NavPaneFrame = By.Id("navpane");
        public static By HelpItemsFrame = By.XPath("//*[@id='tocIFrame']");
        public static By InvoiceSummaryFrame = By.XPath("(//iframe)[2]");

        // Passport Buttons
        public static By MySettingsIcon = By.CssSelector(".navbar-text");
        public static By MyPreferencesButton = By.XPath("//div[normalize-space(text())='My Preferences']");
        public static By EditButton = By.XPath("//button[text()='Edit']");
        public static By PersistentListFlagButton = By.CssSelector("input[data-name='Is Sticky']");
        public static By SaveEditModeButton = By.XPath("//button[@class='editMode' and text()='Save']");
        public static By SaveAddModeButton = By.XPath("//button[@class='addMode' and text()='Save']");

        // Info Items
        public static By InvoiceNetTotal = By.Id("NetTotal");

        // Invoice Sub Tabs
        public static By HeaderAdjustmentTab = By.XPath("//span[@data-name='Adjustment Line Items']");

        // Passport List Item
        public static By ListItems = By.XPath("//tbody //tr");
        public static By HeaderItemNetTotalPassport = By.XPath("//td[6]");

        public static By ListItemByContent(string content)
        {
            return By.XPath($@"//td[.//*[contains(text(),'{content}')]]");
        }

        // Inputs
        public static By InputByDataNameTag(string name) => By.CssSelector($"input[data-name='{name}']");
        public static By InputLookupIconByDataNameTag(string name) =>
            By.XPath($"//div[contains(@class,'search') and following-sibling::input[@data-name='{name}']]");
        public static By AutocompleteDialogItemByIndex(int index) =>
            By.CssSelector($"table.searchResults tbody tr:nth-of-type({index + 1})");

        public static By MatterStatusText = By.XPath("//div[contains(@id,'instance_matterStatus')]");
    }
}