using OpenQA.Selenium;

// ReSharper disable InconsistentNaming

namespace UITests.PageModel.Selectors
{
    public static class Oc
    {
        // Layout regions
        public static By DropTarget = By.XPath(@"//div[@oc-ppo-drop]|//oc-drop-zone//div");

        public static By ProgressIndicator = By.Id("progressIndicatorContainer");
        public static By SpinnerAction = By.XPath("//img[contains(@src,'ajax-loader.gif')]");
        public static By onLoadExistsElm = By.CssSelector("div.top-left");
        public static By SpinnerInitial = By.Id("initialSpinner");
        public static By SpinnerLarge = By.CssSelector(".spinner-large");
        public static By SpinnerLoading = By.CssSelector("div.spinner,core-icon.ms-Icon--actionsSpinner");
        public static By MatProgressBar = By.CssSelector("mat-progress-bar");
        public static By MenuPanel = By.CssSelector("div.mat-menu-panel");
        public static By MenuOptions(bool showHiddenItems = false)
        {
            return By.XPath($"//*[@class='cdk-overlay-container']//div[@class='mat-menu-content']" + (!showHiddenItems ? "//button[not(@hidden)]" : "") + "//span[@class='item-text']");
        }
        public static By CheckBoxBackground = By.ClassName("mat-checkbox-background");
        public static By OcOptionsLabel = By.ClassName("title");
        public static By AllDocumentShowResult = By.ClassName("show-results-button");

        public static By RandomOverlayClick = By.XPath($"//*[@class='cdk-overlay-container']");
        public static By ByClassName(string className) => By.CssSelector($".{className}");

        public static By GroupByIndex(int index) => By.XPath($"(//div[contains(@class, \"panel-header-title\")])[{index + 1}]");

        public static By Group = By.CssSelector(".panel-header-title span");
        public static By GroupHeaderValue = By.CssSelector(".header-value");
        public static By Parent = By.XPath("..");

        // Tabs
        public static By TabByName(string name)
        {
            return By.XPath($"//span[normalize-space(text())='{name}']");
        }

        public static By EntityTabByName(string name)
        {
            return By.XPath($"//oc-navigation-tabs//span[normalize-space(text())='{name}']");
        }

        public static By MattersTab = TabByName("Matters");
        public static By SpendTab = TabByName("Spend");
        public static By DocumentsTab = TabByName("Documents");

        public static By FavoritesTab = EntityTabByName("Favorites");
        public static By AllMattersTab = EntityTabByName("All Matters");
        public static By MyMattersTab = EntityTabByName("My Matters");

        public static By EntityAllDocumentsTab = EntityTabByName("All Documents");
        public static By EntityCheckedOutTab = EntityTabByName("Checked Out");
        public static By EntityRecentDocumentsTab = EntityTabByName("Recent Documents");

        public static By EntityActiveTab = By.XPath("//oc-navigation-tabs//div[contains(@class, 'active')]//span[not(core-icon)]");

        // Form inputs
        public static By HostServerUrlInputBox = By.Id("hostServerUrl");

        public static By UserNameInputBox = By.Id("userName");
        public static By PasswordInputBox = By.Id("passWord");

        public static By DropDownInputByClass(string name)
        {
            return By.CssSelector($".entityform-field-{name} mat-select");
        }

        public static By DropDownItemByIndex(int index)
        {
            return By.CssSelector($".cdk-overlay-container mat-option:nth-child({index})");
        }

        public static By DropDownItemByText(string value)
        {
            return By.XPath($"//*[@class='cdk-overlay-container']//mat-option/span[normalize-space(text())='{value}']");
        }

        public static By MultiSelectInputByClass(string name)
        {
            return By.CssSelector($"[class*=entityform-field-{name}] input");
        }

        public static By MultiSelectItemByName(string name, string item)
        {
            return By.XPath($"//div[contains(@class, \"entityform-field-{name}\")]//a[contains(text(), \"{item}\")]|" +
                            $"//span[@class=\"mat-option-text\"]//span[contains(text(), \"{item}\")]");
        }

        public static By MultiSelectItemByIndex(string name, int index)
        {
            return By.CssSelector($"div.mat-autocomplete-visible .mat-option:nth-child({index + 1}) span div.chip-text");
        }

        public static By MultiSelectInputChips(string name)
        {
            return By.CssSelector($"[class*=entityform-field-{name}] mat-chip");
        }

        public static By MultiSelectItems()
        {
            return By.CssSelector($"div.chip-text");
        }

        public static By RemoveInputChips(string name)
        {
            return By.CssSelector($"[class*=entityform-field-{name}] .mat-chip-remove");
        }

        public static By MultiSelectDialog(string name)
        {
            return By.CssSelector($"[class*=entityform-field-{name}] oc-select-list");
        }

        public static By InputFieldByClass(string inputClass)
        {
            return By.CssSelector($".entityform-field-{inputClass} input");
        }

        public static By TextAreaInputByClass(string name)
        {
            return By.CssSelector($".entityform-field-{name} textarea");
        }

        public static By DateFieldInputByClass(string name)
        {
            return By.CssSelector($".entityform-field-{name} input");
        }

        public static By CheckBoxInputByClass(string name)
        {
            return By.CssSelector($".entityform-field-{name} mat-checkbox");
        }

        public static By CheckBoxInnerInputByClass(string name)
        {
            return By.CssSelector($".entityform-field-{name} input");
        }

        public static By ReadOnlyDialogControlByClass(string name)
        {
            return By.CssSelector($".entityform-field-{name} div.entity-form_readonly_text");
        }

        public static By GetAllLabelsFromDialog = By.XPath("//div[contains(@class, 'ms-Dialog-main')] //div[contains(@class, 'entityform-field-')]");

        public static By MatChipListCrossIcon(string name, int index)
        {
            return By.CssSelector($"[class*=entityform-field-{name}] mat-chip:nth-of-type({index + 1}) core-icon");
        }

        public static By ItemInMultiSelectDialogListByIndex(int index)
        {
            return By.CssSelector($".ms-Dialog-main.multiSelectList-dialog mat-chip:nth-of-type({index + 1}) core-icon");
        }

        public static By MultiselectWindow = By.CssSelector(".ms-Dialog-main.multiSelectList-dialog");
        public static By ItemsInMultiSelectWindow = By.CssSelector(".ms-Dialog-main.multiSelectList-dialog mat-chip");
        public static By MultiselectListItem = By.CssSelector(".ms-Dialog-main.multiSelectList-dialog mat-row");

        // Calendar control
        public static By CalendarControl = By.CssSelector(".owl-dt-calendar-body");

        public static By CalendarDefault = By.CssSelector(".owl-dt-calendar-cell-today");
        public static By CalendarSelectedDate = By.CssSelector(".owl-dt-calendar-cell-active");
        public static By CalendarSetButton = By.XPath("//span[normalize-space(text())='Set']");

        public static By CalendarPreviousMonthButton =
            By.XPath("//button[contains(@class, 'owl-dt-control-button') and contains(@aria-label, 'Previous')]");

        public static By CalendarNextMonthButton =
            By.XPath("//button[contains(@class, 'owl-dt-control-button') and contains(@aria-label, 'Next')]");

        public static By Backdrop = By.CssSelector("div.cdk-overlay-container");

        // Breadcrumbs control
        public static By BreadcrumbsFolders =
            By.CssSelector("div.breadcrumbs-item.breadcrumbs-item-active div.breadcrumbs-item-text");

        public static By BreadcrumbsRootFolder = By.CssSelector("div.breadcrumbs-items-folder-icon");

        // Buttons
        public static By ButtonByName(string name)
        {
            return By.XPath($@"//span[contains(text(), '{name}')]");
        }

        public static By FilterList = ButtonByName("Filter List");
        public static By FilterSaveView = ButtonByName("Save Current View");
        public static By FilterSavedViews = ButtonByName("Saved Views");
        public static By FilterSetViewAsDefault = ButtonByName("Set Current View as Default");
        public static By FilterClearUserDefault = ButtonByName("Clear User Default");
        public static By AddMatter = ButtonByName("Add Matter");

        public static By ButtonIcon(string btnName)
        {
            return By.CssSelector($@".ms-Icon--{btnName}");
        }

        public static By DeleteButton = ButtonIcon("Delete");
        public static By EditButton = ButtonIcon("EditSolid12");
        public static By ListOptions = ButtonIcon("More");
        public static By RemoveButton = ButtonIcon("Blocked2");
        public static By RenameButton = ButtonIcon("Rename");
        public static By DownloadButton = ButtonIcon("Download");
        public static By SummaryIcon = ButtonIcon("Info");
        public static By QuickSearchIcon = ButtonIcon("Search");
        public static By FilterIcon = ButtonIcon("FilterSolid");
        public static By EmailAttachment = ButtonIcon("attachment");
        public static By AccessButton = ButtonIcon("Link12");
        public static By EmailIcon = ButtonIcon("Mail");
        public static By ContactIcon = ButtonIcon("ContactCard");
        public static By ApplyIcon = ButtonIcon("Save");

        public static By FormApplyButton = By.CssSelector("[id*='Apply']");
        public static By FavoriteToggle = By.CssSelector("oc-action-favorite span");
        public static By RestoreDefaults = By.XPath("//div[contains(text(), 'Restore Defaults')]|//span[normalize-space(text())='Restore Defaults']");
        public static By ListItemQuickFile = By.CssSelector(".ms-Icon--Upload");
        public static By UploadIndicatorCounter = By.CssSelector("[name='Upload'] + core-icon");
        public static By AddButton = By.XPath("//span[normalize-space(text())='Add']");
        public static By AddFolderButton = By.XPath("//span[normalize-space(text())='Add Folder']");
        public static By SaveButton = By.XPath("//div[normalize-space(text())='Save']");
        public static By SaveAsButton = By.XPath("//div[normalize-space(text())='Save as']");
        public static By SaveWithANewName = By.XPath("//div[normalize-space(text())='Save with a new name']");
        public static By UpdateButton = By.CssSelector("[id*='Update']");
        public static By UploadDocumentButton = By.XPath("//div[contains(text(), 'Upload Document')]|//div[contains(text(), 'Upload document')]|//span[normalize-space(text())='Upload Document']|//div[contains(text(), 'Proceed')]");
        public static By SaveAndUploadButton = By.XPath("//div[normalize-space(text())='Save & Upload']");
        public static By CancelButton = By.XPath("//div[contains(text(), 'Cancel')]");
        public static By SearchInput = By.CssSelector("core-table-search input");
        public static By CheckBox = By.CssSelector("div.mat-checkbox-inner-container");
        public static By DeleteEmailsButton = By.XPath("//span[normalize-space(text())='Delete Emails']");
        public static By ImportMatterCalendarButton = By.XPath("//div[contains(@class, 'wide')]//core-icon[contains(@class, 'ms-Icon--Calendar')]");
        public static By ToastMessageCloseButton = By.CssSelector("button.toast-close-button");
        public static By ActionItems = By.ClassName("actions-items");
        public static By UploadIndicatorStopWatch = By.XPath("//*[@name='Clock']");
        public static By QueuedEmailCount = By.XPath("//div[@class='icon-set']/core-icon[2]");
        public static By UploadEmailWaitingQueueCount = By.XPath("//*[@name='Clock']//following::core-icon[1]");
        public static By SignOut = By.CssSelector("[data-qa='signOutBtn']");

        public static By CheckBoxByNameAttribute(string name)
        {
            return By.CssSelector($"div.mat-checkbox-inner-container input[name = '{name}']");
        }

        // Header
        public static By RefreshIcon = By.CssSelector("[name='Refresh']");

        public static By UploadQueueIcon = By.CssSelector("[name='Upload']");
        public static By OpenUploadHistory = By.CssSelector("[name='FullHistory']");
        public static By CancelAll = By.CssSelector("[name='RemoveFromShoppingList']");
        public static By ClearUploadHistory = By.CssSelector("[name='Delete']");
        public static By CloseUploadHistory = By.CssSelector("[name='Cancel']");
        public static By HelpIcon = By.CssSelector("[name='Help']");
        public static By BackButton = By.CssSelector("[name='NavigateBack']");
        public static By SettingsButton = ButtonIcon("Settings");

        // sort buttons
        public static By RestoreSortDefaults = By.XPath("//span[normalize-space(text())='Restore Sort Defaults']");

        public static By SortIcon = ButtonIcon("SortLines");
        public static By SortOptions = By.CssSelector(".item-text");

        public static By SortButtonByOption(string option) => By.XPath($"//button[span[normalize-space(text())='{option}']]");

        public static By SortIconByOption(string option) => By.XPath($"//button[span[normalize-space(text())='{option}']]/core-icon");

        // Filter buttons
        public static By RemoveSavedView(string viewName) => By.XPath($"//button[span[normalize-space(text())='{viewName}']]/button");

        public static By RadioButtonByLabel(string radioButtonLabel) => By.XPath($"//label[normalize-space(text())='{radioButtonLabel}']/input[@type='radio']");

        // invoice buttons
        public static By InvoiceItemApproveButton = ButtonIcon("CheckMark");

        public static By InvoiceItemRejectButton = ButtonIcon("Cancel");

        // List Items
        public static By ListItems = By.CssSelector("mat-row");

        public static By NthListItem(int index) => By.CssSelector($"mat-row:nth-of-type({index + 1})");

        public static By ListItemByContent(string content)
        {
            return By.XPath($@"//mat-row[.//*[contains(text(),'{content.Trim()}')]]");
        }

        public static By MatterPropertyByClass(string name)
        {
            return By.CssSelector($".entityform-field-{name} span span:not(.dotted_view_separator)");
        }

        // Info Items
        public static By InvoiceValueByClass(string className)
        {
            return By.CssSelector($".entityform-field-{className} div.value");
        }

        public static By Tooltip = By.CssSelector("div.mat-tooltip");
        public static By ToastMessage = By.CssSelector("div.toast-message");
        public static By LineItemAmount = By.XPath("(//div[@class='items-stretch-container']//div[@class='left-item'])[last()]//span[last()]");
        public static By ListCount = By.CssSelector(".info-footer");
        public static By ListEmptyMessage = By.CssSelector(".no-records-row");

        public static By OfficeAppType(string appName) => ButtonIcon(appName + "Document");

        // CheckIn/CheckOut
        public static By ItemOptions = By.CssSelector("span[data-qa='Options']");

        public static By AllOptions = By.XPath("//*[@class='cdk-overlay-container']//div[@class='mat-menu-content']//button");
        public static By AllOptionsOverlay = By.XPath("//*[@class='cdk-overlay-container']//div[contains(@class, 'mat-menu-panel')]");

        public static By ItemOptionsCheckOutButton = By.CssSelector("button[data-qa='Check Out']");
        public static By ItemOptionsCheckInButton = By.CssSelector("button[data-qa='Check In']");
        public static By ItemOptionsDiscardCheckOutButton = By.CssSelector("button[data-qa='Discard Checkout']");

        //Frame Items
        public static By Dialog = By.XPath("(//*[@class='ms-Dialog-main '])[last()]");

        public static By DialogHeader = By.XPath("(//*[@class='ms-Dialog-header'])[last()]");
        public static By DialogActionMessage = By.CssSelector(".ms-Dialog-content.dialog-content.ng-star-inserted");
        public static By DialogButton = By.XPath("(//*[contains(@class, 'ms-Dialog-actions')])[last()]//button");

        public static By DialogHeaderByName(string headerName)
        {
            return By.XPath($@"//p[contains(@class, 'ms-Dialog-title') and contains(text(),'{headerName}')]");
        }

        public static By DialogKeepButton = DialogButtonById("Keep");
        public static By DialogRemoveOnDiscardButton = DialogButtonById("Remove");
        public static By DialogOkButton = DialogButtonById("Ok");
        public static By DialogDiscardChangesButton = DialogButtonById("Discard Changes");
        public static By DialogResetButton = DialogButtonById("Reset");
        public static By DialogEditButton = DialogButtonById("Edit");
        public static By DialogApproveButton = DialogButtonById("Approve");
        public static By DialogRejectButton = DialogButtonById("Reject");
        public static By DialogDoNotDiscardButton = DialogButtonById("Don't Discard");
        public static By DialogDone = DialogButtonById("Done");
        public static By DialogClose = DialogButtonById("Close");
        public static By DialogOverwriteButton = DialogButtonById("Overwrite");
        public static By DialogSelectFileButton = DialogButtonById("Select file");

        private static By DialogButtonById(string btnName)
        {
            return By.Id($@"dialog_button_{btnName}_");
        }

        public static By PanelByIndex(int index)
        {
            return By.XPath($"(//mat-expansion-panel-header)[{index}]");
        }

        // Document Summmary
        public static By TextInfo = By.CssSelector(".entity-form_readonly_text");

        public static By BreadCrumbsParent = By.XPath("(//div[contains(@class,'summary-breadcrumbs-parent-label')])[last()]");
        public static By ItemDetailQuickFile = By.XPath("//mat-expansion-panel-header//core-icon[contains(@class, 'ms-Icon--Upload')]");
    }
}
