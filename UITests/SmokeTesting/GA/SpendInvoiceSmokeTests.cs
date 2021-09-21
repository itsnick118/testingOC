using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using IntegratedDriver;
using NUnit.Framework;
using UITests.PageModel;
using UITests.PageModel.OfficeApps;
using UITests.PageModel.Shared;
using UITests.PageModel.Shared.Comparators;
using static UITests.Constants;
using static UITests.TestHelpers;

namespace UITests.SmokeTesting.GA
{
    public class SpendInvoiceSmokeTests : UITestBase
    {
        private Outlook _outlook;
        private Word _word;

        [SetUp]
        public void SetUp()
        {
            _outlook = new Outlook(TestEnvironment);
            _outlook.Launch();
            _outlook.Oc.BasicSettingsPage.LogInAsAttorneyUser();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC_16723 : Verify Invoice List Page Actions Approve and Reject")]
        public void VerifyInvoiceListApproveReject()
        {
            var invoicesListPage = _outlook.Oc.InvoicesListPage;
            var myInvoicesList = invoicesListPage.ItemList;

            invoicesListPage.Open();
            var myInvoices = myInvoicesList.GetAllInvoiceListItems();
            Assert.GreaterOrEqual(myInvoices.Count, 1, "Invoices list is not loaded or has no items");
            Assert.IsFalse(_outlook.Oc.IsErrorDisplayed(), "Invoices list is loaded with error message");

            Assert.That(myInvoices, Has.All.Property(nameof(InvoiceListItem.HasApproveButton)).True, "My Invoice list has items without Approve button");

            // select random invoice to approve
            var selectedInvoice = myInvoicesList.GetInvoiceListItemByIndex(GetRandomNumber(myInvoices.Count - 1));
            var selectedInvoiceName = selectedInvoice.PrimaryElement.Text;
            selectedInvoice.Approve();
            invoicesListPage.ApproveInvoiceDialog.Controls["Internal Comment"].Set(AutomatedComment);
            invoicesListPage.ApproveInvoiceDialog.Controls["External Comment"].Set(AutomatedComment);

            // approve invoice
            invoicesListPage.ApproveInvoiceDialog.Approve();
            Assert.IsNull(myInvoicesList.GetInvoiceListItemFromText(selectedInvoiceName));

            // select random invoice to reject
            myInvoices = myInvoicesList.GetAllInvoiceListItems();
            selectedInvoice = myInvoicesList.GetInvoiceListItemByIndex(GetRandomNumber(myInvoices.Count - 1));
            selectedInvoiceName = selectedInvoice.PrimaryElement.Text;
            selectedInvoice.Reject();
            invoicesListPage.RejectInvoiceDialog.Controls["Reject Reason Codes"].Set("Billing");
            invoicesListPage.RejectInvoiceDialog.Controls["Internal Comment"].Set(AutomatedComment);
            invoicesListPage.RejectInvoiceDialog.Controls["External Comment"].Set(AutomatedComment);

            // reject invoice
            invoicesListPage.RejectInvoiceDialog.Reject();
            Assert.IsNull(myInvoicesList.GetInvoiceListItemFromText(selectedInvoiceName));
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC_16671 : Verify Filter Options in Invoice documents list (End to End funtional test cases: Step 1 & 2)")]
        public void VerifyFilterOptionsInInvoiceDocumentsListEndToEnd()
        {
            var invoicesListPage = _outlook.Oc.InvoicesListPage;
            var myInvoicesList = invoicesListPage.ItemList;
            var invoiceListPage = _outlook.Oc.InvoicesListPage;
            var invoiceSummaryPage = _outlook.Oc.InvoiceSummaryPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var documentSummaryPage = _outlook.Oc.DocumentSummaryPage;
            var randomFileContent = GetRandomText(10);

            var wordFile = CreateDocument(OfficeApp.Word, randomFileContent);
            invoiceListPage.Open();
            Assert.GreaterOrEqual(myInvoicesList.GetCount(), 1);

            myInvoicesList.OpenRandom();
            invoiceSummaryPage.EntityTabs.Open("Documents");

            DragAndDrop.FromFileSystem(wordFile, documentsListPage.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.Controls["Comments"].Set(AutomatedComment + " for " + wordFile.Name);

            documentsListPage.AddDocumentDialog.UploadDocument();
            invoiceSummaryPage.QuickSearch.SearchBy(wordFile.Name);

            var selectedDocument = documentsListPage.ItemList.GetInvoiceDocumentListItemFromText(wordFile.Name);

            // Fetch Document Properties
            var _fileName = selectedDocument.Name;
            var _name = selectedDocument.Name;
            var _fileStatus = selectedDocument.Status;

            var _fileUpdatedBy = selectedDocument.LastModifiedBy;
            var _updatedAt = FormatDateRange(selectedDocument.UpdatedAt.Date, selectedDocument.UpdatedAt.AddDays(3).Date);

            // To Fetch Document Comments
            selectedDocument.NavigateToSummary();
            documentSummaryPage.SummaryPanel.Toggle();
            var documentVersion = documentsListPage.ItemList.GetVersionHistoryListItemByIndex(0);

            documentSummaryPage.NavigateToParent();
            var _filecomments = documentVersion.Comments;

            // To Fetch Document Content
            selectedDocument = documentsListPage.ItemList.GetInvoiceDocumentListItemFromText(wordFile.Name);
            selectedDocument.Open();
            _word = new Word(TestEnvironment);
            _word.Attach(wordFile.Name);
            var _fileContent = _word.ReadActiveFileContent();
            _word.Close();

            documentsListPage = _outlook.Oc.DocumentsListPage;
            var documentFilterDialog = documentsListPage.InvoiceDocumentListFilterDialog;
            var myDocumentList = documentsListPage.ItemList;

            var filterOptions = new Dictionary<string, string>()
            {
                    { "File Name", _fileName },
                    { "Name", _name },
                    { "Status", _fileStatus },
                    { "Updated By", _fileUpdatedBy },
                    { "Comments", _filecomments },
                    { "Updated At", _updatedAt},
                    { "Content",_fileContent }
             };

            foreach (var filterOption in filterOptions)
            {
                myDocumentList.OpenListOptionsMenu().OpenCreateListFilterDialog();
                documentFilterDialog.Controls[filterOption.Key].Set(filterOption.Value);
                documentFilterDialog.Apply();

                Assert.IsTrue(myDocumentList.IsFilterIconVisible, "Filter Icon is not visible");
                Assert.AreEqual(myDocumentList.GetFilterIconColor().Name, BlueColorName, "Filter Icon is not expected color");
                Assert.AreEqual(myDocumentList.FilterIconToolTip, FilterIconToolTip, "Filter Icon tooltip differs from expected");

                var invoiceDocumentItems = myDocumentList.GetAllInvoiceDocumentListItems();
                Assert.IsNotEmpty(invoiceDocumentItems, $"No records found on applying filter by { filterOption.Key} - {filterOption.Value}");

                Assert.NotNull(invoiceDocumentItems.Select(x => x.GetType().GetProperty(filterOption.Value)));

                // Restore Defaults
                myDocumentList.OpenListOptionsMenu().RestoreDefaults();
                Assert.IsFalse(myDocumentList.IsFilterIconVisible, "Filter Icon is visible");
            }

            //Clean up
            selectedDocument = documentsListPage.ItemList.GetInvoiceDocumentListItemFromText(wordFile.Name);

            if (selectedDocument.Status == CheckInStatus.CheckedOut)
            {
                selectedDocument.FileOptions.CheckIn();
                documentFilterDialog.Controls["Comments"].Set("Clean up Operation");
                documentFilterDialog.UploadDocument();
            }
            selectedDocument.Delete().Confirm();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC_16671 : Verify Filter Options in Invoice documents list (End to End funtional test cases): Step 3,4,5 Save Current View")]
        public void VerifyInvoiceDocumentUpdateDeleteCurrentViewInFilterOperations()
        {
            var invoicesListPage = _outlook.Oc.InvoicesListPage;
            var myInvoicesList = invoicesListPage.ItemList;
            var invoiceListPage = _outlook.Oc.InvoicesListPage;
            var myInvoiceList = invoiceListPage.ItemList;
            var invoiceSummaryPage = _outlook.Oc.InvoiceSummaryPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var myDocumentList = documentsListPage.ItemList;

            invoiceListPage.Open();
            var myInvoices = myInvoiceList.GetAllInvoiceListItems();
            Assert.GreaterOrEqual(myInvoices.Count, 1, "Invoices list is not loaded or has no items");

            // select invoice
            Assert.GreaterOrEqual(myInvoicesList.GetCount(), 1);
            myInvoicesList.OpenFirst();

            invoiceSummaryPage.EntityTabs.Open("Documents");
            Assert.GreaterOrEqual(myDocumentList.GetCount(), 1, "Document list is not loaded or has no items");

            // Select a document
            var selectedDocument = myDocumentList.GetMatterDocumentListItemByIndex(GetRandomNumber(myDocumentList.GetCount() - 1));

            // Fetch Document Properties
            var fileName = selectedDocument.DocumentFileName;
            var name = selectedDocument.Name;

            myDocumentList.OpenListOptionsMenu().OpenCreateListFilterDialog();

            var documentFilterDialog = documentsListPage.InvoiceDocumentListFilterDialog;

            // Filter by File Name
            documentFilterDialog.Controls["File Name"].Set(fileName);
            documentFilterDialog.Apply();

            Assert.IsTrue(myDocumentList.IsFilterIconVisible, "Filter Icon is not visible");
            Assert.AreEqual(myDocumentList.GetFilterIconColor().Name, BlueColorName, "Filter Icon is not expected color");
            Assert.AreEqual(myDocumentList.FilterIconToolTip, FilterIconToolTip, "Filter Icon tooltip differs from expected");

            var filteredList = myDocumentList.GetAllInvoiceDocumentListItems();
            Assert.IsNotEmpty(filteredList, $"No records found on applying filter by FileName - {fileName}");
            Assert.That(filteredList, Has.All.Property(nameof(InvoiceDocumentListItem.Name)).Contains(fileName).IgnoreCase, $"Filtered list has items not containing Name - {fileName}");

            // Save View
            myDocumentList.OpenListOptionsMenu().SaveCurrentView();
            documentsListPage.SaveCurrentViewDialog.Controls["Create New"].Set(ViewName);
            documentsListPage.SaveCurrentViewDialog.Save();

            // Clear Filters By Restauring Defaults
            myDocumentList.OpenListOptionsMenu().RestoreDefaults();
            Assert.IsFalse(myDocumentList.IsFilterIconVisible, "Filter Icon is visible");

            // Apply Saved View
            myDocumentList.OpenListOptionsMenu().ApplySavedView(ViewName);
            Assert.AreEqual(filteredList.Count, myDocumentList.GetCount());
            Assert.IsTrue(myDocumentList.IsFilterIconVisible, "Filter Icon is not visible");

            // Filter by File Name
            myDocumentList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentFilterDialog = documentsListPage.InvoiceDocumentListFilterDialog;
            documentFilterDialog.Controls["Name"].Set(name);
            documentFilterDialog.Apply();

            Assert.IsTrue(myDocumentList.IsFilterIconVisible, "Filter Icon is not visible");
            Assert.AreEqual(myDocumentList.GetFilterIconColor().Name, BlueColorName, "Filter Icon is not expected color");
            Assert.AreEqual(myDocumentList.FilterIconToolTip, FilterIconToolTip, "Filter Icon tooltip differs from expected");

            filteredList = myDocumentList.GetAllInvoiceDocumentListItems();
            Assert.IsNotEmpty(filteredList, $"No records found on applying filter by FileName - {name}");
            Assert.That(filteredList, Has.All.Property(nameof(InvoiceDocumentListItem.Name)).Contains(name).IgnoreCase, $"Filtered list has items not containing Name - {name}");

            //Update existing View
            myDocumentList.OpenListOptionsMenu().SaveCurrentView();
            documentsListPage.SaveCurrentViewDialog.ClickRadioButton("Update Existing");
            documentsListPage.SaveCurrentViewDialog.Controls["Update Existing"].SetByIndex(1);
            documentsListPage.SaveCurrentViewDialog.Save();

            // Cancel Setting a View
            myDocumentList.OpenListOptionsMenu().SaveCurrentView();
            documentsListPage.SaveCurrentViewDialog.Controls["Create New"].Set(ViewName);
            documentsListPage.SaveCurrentViewDialog.Cancel();

            // Delete Existing view
            myDocumentList.OpenListOptionsMenu().RestoreDefaults();
            myDocumentList.OpenListOptionsMenu().RemoveSavedView(ViewName);
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("Verify Filter Options in Invoice documents list (End to End funtional test cases): Step 3-4, Save Current View")]
        public void VerifyInvoiceSetCurrentViewAsDefaultInFilterOperations()
        {
            var invoicesListPage = _outlook.Oc.InvoicesListPage;
            var myInvoicesList = invoicesListPage.ItemList;
            var invoiceListPage = _outlook.Oc.InvoicesListPage;
            var myInvoiceList = invoiceListPage.ItemList;
            var invoiceSummaryPage = _outlook.Oc.InvoiceSummaryPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var myDocumentList = documentsListPage.ItemList;
            var _settingsPage = _outlook.Oc.SettingsPage;
            invoiceListPage.Open();

            // select invoice
            Assert.GreaterOrEqual(myInvoicesList.GetCount(), 1);
            myInvoicesList.OpenFirst();

            invoiceSummaryPage.EntityTabs.Open("Documents");
            Assert.GreaterOrEqual(myDocumentList.GetCount(), 1, "Document list is not loaded or has no items");

            // Select a document
            var selectedDocument = myDocumentList.GetMatterDocumentListItemByIndex(GetRandomNumber(myDocumentList.GetCount() - 1));

            // Fetch Document Properties
            var fileName = selectedDocument.DocumentFileName;

            myDocumentList.OpenListOptionsMenu().OpenCreateListFilterDialog();

            var documentFilterDialog = documentsListPage.InvoiceDocumentListFilterDialog;

            // Filter by File Name
            documentFilterDialog.Controls["File Name"].Set(fileName);
            documentFilterDialog.Apply();
            var filterCount = myDocumentList.GetCount();

            Assert.IsTrue(myDocumentList.IsFilterIconVisible, "Filter Icon is not visible");
            Assert.AreEqual(myDocumentList.GetFilterIconColor().Name, BlueColorName, "Filter Icon is not expected color");
            Assert.AreEqual(myDocumentList.FilterIconToolTip, FilterIconToolTip, "Filter Icon tooltip differs from expected");

            var filteredList = myDocumentList.GetAllInvoiceDocumentListItems();
            Assert.IsNotEmpty(filteredList, $"No records found on applying filter by FileName - {fileName}");
            Assert.That(filteredList, Has.All.Property(nameof(InvoiceDocumentListItem.Name)).Contains(fileName).IgnoreCase, $"Filtered list has items not containing Name - {fileName}");

            // Set view as default
            myDocumentList.OpenListOptionsMenu().SetCurrentViewAsDefault();

            // Verify current view is set as default
            myDocumentList.OpenListOptionsMenu().RestoreDefaults();
            Assert.IsTrue(myDocumentList.IsFilterIconVisible, "Filter Icon is not visible");
            Assert.AreEqual(myDocumentList.GetFilterIconColor().Name, BlueColorName, "Filter Icon is not expected color");

            // Logout sbrown
            _outlook.Oc.OpenSettings();
            _settingsPage.OpenConfiguration();
            _outlook.Oc.WaitForLoadComplete();
            _settingsPage.LogOut().Confirm();

            // Log in office companion again
            _outlook.Oc.BasicSettingsPage.LogInAsAttorneyUser();
            invoiceListPage.Open();

            var myInvoices = myInvoiceList.GetAllInvoiceListItems();
            myDocumentList = documentsListPage.ItemList;
            Assert.GreaterOrEqual(myInvoices.Count, 1, "Invoices list is not loaded or has no items");

            // select invoice
            myInvoicesList.OpenFirst();

            invoiceSummaryPage.EntityTabs.Open("Documents");
            Assert.GreaterOrEqual(myDocumentList.GetCount(), 1, "Document list is not loaded or has no items");

            Assert.AreEqual(myDocumentList.GetCount(), filterCount);

            myDocumentList.OpenListOptionsMenu().ClearUserDefault();
            Assert.IsFalse(myDocumentList.IsFilterIconVisible, "Filter Icon is visible");
            Assert.IsFalse(myDocumentList.OpenListOptionsMenu().IsClearUserDefaultDisplayed(), "Clear User Default Link is Visible");
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC_16724 : Verify Invoice List Filters")]
        public void InvoiceListFilters()
        {
            var invoicesListPage = _outlook.Oc.InvoicesListPage;
            var myInvoicesList = invoicesListPage.ItemList;

            invoicesListPage.Open();
            var myInvoices = myInvoicesList.GetAllInvoiceListItems();
            Assert.GreaterOrEqual(myInvoices.Count, 1, "Invoices list is not loaded or has no items");

            // select random invoice
            var selectedInvoice = myInvoicesList.GetInvoiceListItemByIndex(GetRandomNumber(myInvoices.Count - 1));

            // fetch invoice properties
            var invoiceNumber = GetRandomSubstring(selectedInvoice.Number);
            var receivedDate = selectedInvoice.ReceivedDate.Date;
            var organizationName = GetRandomSubstring(selectedInvoice.OrganizationName);
            var vendorId = GetRandomSubstring(selectedInvoice.VendorId);
            var matterName = GetRandomSubstring(selectedInvoice.MatterName);
            var matterNumber = GetRandomSubstring(selectedInvoice.MatterNumber);

            // apply filter - number
            myInvoicesList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            var invoiceFilterDialog = invoicesListPage.InvoiceListFilterDialog;
            invoiceFilterDialog.Controls["Invoice Number"].Set(invoiceNumber);
            invoiceFilterDialog.Apply();
            Assert.IsTrue(myInvoicesList.IsFilterIconVisible, "Filter Icon is not visible");
            Assert.AreEqual(myInvoicesList.GetFilterIconColor().Name, BlueColorName, "Filter Icon is not expected color");
            Assert.AreEqual(myInvoicesList.FilterIconToolTip, FilterIconToolTip, "Filter Icon tooltip differs from expected");

            var filteredList = myInvoicesList.GetAllInvoiceListItems();
            Assert.IsNotEmpty(filteredList, $"No records found on applying filter by FileName - {invoiceNumber}");
            Assert.That(filteredList, Has.All.Property(nameof(InvoiceListItem.Number)).Contains(invoiceNumber).IgnoreCase, $"Filtered list has items not containing Name - {invoiceNumber}");

            // restore defaults
            myInvoicesList.OpenListOptionsMenu().RestoreDefaults();
            Assert.IsFalse(myInvoicesList.IsFilterIconVisible, "Filter Icon is visible");

            // apply filter - received date
            var dateTo = receivedDate.AddDays(3);
            myInvoicesList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            invoiceFilterDialog.Controls["Received Date"].Set(FormatDateRange(receivedDate, dateTo));
            invoiceFilterDialog.Apply();
            Assert.IsTrue(myInvoicesList.IsFilterIconVisible, "Filter Icon is not visible");
            Assert.AreEqual(myInvoicesList.GetFilterIconColor().Name, BlueColorName, "Filter Icon is not expected color");
            Assert.AreEqual(myInvoicesList.FilterIconToolTip, FilterIconToolTip, "Filter Icon tooltip differs from expected");

            var endDate = dateTo.AddSeconds(-1).AddDays(1);
            filteredList = myInvoicesList.GetAllInvoiceListItems();
            Assert.IsNotEmpty(filteredList, $"(No records found on applying filter by Received Date : {receivedDate} - {dateTo}");
            Assert.That(filteredList, Has.All.Property(nameof(InvoiceListItem.ReceivedDate)).InRange(receivedDate, endDate), $"Filtered list has items out of Received Date : {receivedDate} - {endDate}");

            // apply filter - organization name
            myInvoicesList.OpenListOptionsMenu().RestoreDefaults();
            myInvoicesList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            invoiceFilterDialog.Controls["Organization Name"].Set(organizationName);
            invoiceFilterDialog.Apply();
            Assert.IsTrue(myInvoicesList.IsFilterIconVisible, "Filter Icon is not visible");

            filteredList = myInvoicesList.GetAllInvoiceListItems();
            Assert.IsNotEmpty(filteredList, $"No records found on applying filter by Organization Name - {organizationName}");
            Assert.That(filteredList, Has.All.Property(nameof(InvoiceListItem.OrganizationName)).Contains(organizationName).IgnoreCase, $"Filtered list has items not containing Organization Name - {organizationName}");

            // apply filter - vendor id
            myInvoicesList.OpenListOptionsMenu().RestoreDefaults();
            myInvoicesList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            invoiceFilterDialog.Controls["Vendor Id"].Set(vendorId);
            invoiceFilterDialog.Apply();
            Assert.IsTrue(myInvoicesList.IsFilterIconVisible, "Filter Icon is not visible");

            filteredList = myInvoicesList.GetAllInvoiceListItems();
            Assert.IsNotEmpty(filteredList, $"No records found on applying filter by Vendor Id - {vendorId}");
            Assert.That(filteredList, Has.All.Property(nameof(InvoiceListItem.VendorId)).Contains(vendorId).IgnoreCase, $"Filtered list has items not containing Vendor Id - {vendorId}");

            // apply filter - matter name
            myInvoicesList.OpenListOptionsMenu().RestoreDefaults();
            myInvoicesList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            invoiceFilterDialog.Controls["Matter Name"].Set(matterName);
            invoiceFilterDialog.Apply();
            Assert.IsTrue(myInvoicesList.IsFilterIconVisible, "Filter Icon is not visible");

            filteredList = myInvoicesList.GetAllInvoiceListItems();
            Assert.IsNotEmpty(filteredList, $"No records found on applying filter by Matter Name - {matterName}");
            Assert.That(filteredList, Has.All.Property(nameof(InvoiceListItem.MatterName)).Contains(matterName).IgnoreCase, $"Filtered list has items not containing Matter Name - {matterName}");

            // apply filter - matter number
            myInvoicesList.OpenListOptionsMenu().RestoreDefaults();
            myInvoicesList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            invoiceFilterDialog.Controls["Matter Number"].Set(matterNumber);
            invoiceFilterDialog.Apply();
            Assert.IsTrue(myInvoicesList.IsFilterIconVisible, "Filter Icon is not visible");

            filteredList = myInvoicesList.GetAllInvoiceListItems();
            Assert.IsNotEmpty(filteredList, $"No records found on applying filter by Matter Number - {matterNumber}");
            Assert.That(filteredList, Has.All.Property(nameof(InvoiceListItem.MatterNumber)).Contains(matterNumber).IgnoreCase, $"Filtered list has items not containing Matter Number - {matterNumber}");
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC_16724 : Verify Invoice List Sort")]
        public void InvoiceListSort()
        {
            const int UniqueInvoicesNumber = 2;

            var invoiceListPage = _outlook.Oc.InvoicesListPage;
            var myInvoicesList = invoiceListPage.ItemList;
            var invoiceSortOptions = new[] { "Invoice Number", "Total Net Amount", "Received Date", "Matter", "Organization Name", "Vendor ID" };

            invoiceListPage.Open();
            var myInvoices = myInvoicesList.GetAllInvoiceListItems();
            Assert.GreaterOrEqual(myInvoices.Count, 1, "Invoices list is not loaded or has no items");

            // verify sort list options
            Assert.IsTrue(myInvoicesList.IsSortIconVisible, "Sort Icon is not visible");

            invoiceListPage.InvoiceSortDialog.OpenSortDialog();
            Assert.AreEqual(invoiceListPage.InvoiceSortDialog.GetSortOptions(), invoiceSortOptions, "Sort options are not displayed");
            invoiceListPage.InvoiceSortDialog.CloseSortDialog();

            // verify default sort
            Assert.That(myInvoices.GroupBy(x => new { x.Number, x.ReceivedDate }).Count, Is.GreaterThanOrEqualTo(UniqueInvoicesNumber),
                "Two or more invoices with different numbers and received date are required to verify sorting.");
            Assert.That(myInvoices, Is.Ordered.Ascending.By(nameof(InvoiceListItem.Number)).
                Then.Ascending.By(nameof(InvoiceListItem.ReceivedDate)));

            // verify to sort invoice list based on invoice number
            invoiceListPage.InvoiceSortDialog.Sort("Invoice Number", SortOrder.Descending);
            myInvoices = myInvoicesList.GetAllInvoiceListItems();
            Assert.That(myInvoices, Is.Ordered.Descending.By(nameof(InvoiceListItem.Number)));

            invoiceListPage.InvoiceSortDialog.Sort("Invoice Number", SortOrder.Ascending);
            myInvoices = myInvoicesList.GetAllInvoiceListItems();
            Assert.That(myInvoices, Is.Ordered.Ascending.By(nameof(InvoiceListItem.Number)));

            // restore defaults
            invoiceListPage.InvoiceSortDialog.RestoreSortDefaults();
            myInvoices = myInvoicesList.GetAllInvoiceListItems();
            Assert.That(myInvoices, Is.Ordered.Ascending.By(nameof(InvoiceListItem.Number)).
                Then.Ascending.By(nameof(InvoiceListItem.ReceivedDate)));

            // verify to sort invoice list based on total net amount
            invoiceListPage.InvoiceSortDialog.Sort("Total Net Amount", SortOrder.Descending);
            myInvoices = myInvoicesList.GetAllInvoiceListItems();
            Assert.That(myInvoices, Is.Ordered.Descending.By(nameof(InvoiceListItem.TotalNetAmount)).Using(new CurrencyComparer()));

            invoiceListPage.InvoiceSortDialog.Sort("Total Net Amount", SortOrder.Ascending);
            myInvoices = myInvoicesList.GetAllInvoiceListItems();
            Assert.That(myInvoices, Is.Ordered.Ascending.By(nameof(InvoiceListItem.TotalNetAmount)).Using(new CurrencyComparer()));

            // restore defaults
            invoiceListPage.InvoiceSortDialog.RestoreSortDefaults();
            myInvoices = myInvoicesList.GetAllInvoiceListItems();
            Assert.That(myInvoices, Is.Ordered.Ascending.By(nameof(InvoiceListItem.Number)).
                Then.Ascending.By(nameof(InvoiceListItem.ReceivedDate)));
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC_16725 : Verify Invoice Summary Actions Approve and Reject")]
        public void InvoiceSummaryApproveReject()
        {
            const string MyInvoicesTab = "My Invoices";
            var invoiceListPage = _outlook.Oc.InvoicesListPage;
            var myInvoicesList = invoiceListPage.ItemList;
            var invoiceSummaryPage = _outlook.Oc.InvoiceSummaryPage;

            invoiceListPage.Open();

            var myInvoices = myInvoicesList.GetCount();
            var selectedInvoice = myInvoicesList.GetInvoiceListItemByIndex(GetRandomNumber(myInvoices - 1));
            var selectedInvoiceName = selectedInvoice.PrimaryElement.Text;

            // verify invoice summary
            selectedInvoice.Open();
            var summaryInfo = invoiceSummaryPage.GetInvoiceSummaryInfo();
            Assert.IsNotEmpty(summaryInfo, "Invoice Summary fields are not retrieved or empty.");

            invoiceSummaryPage.InvoiceTotalsPanel.Toggle();
            foreach (var webElement in summaryInfo)
            {
                Assert.IsNotEmpty(webElement.Text);
            }

            // verify invoice summary approve
            invoiceSummaryPage.Approve();
            invoiceSummaryPage.ApproveInvoiceDialog.Controls["Internal Comment"].Set(AutomatedComment);
            invoiceSummaryPage.ApproveInvoiceDialog.Controls["External Comment"].Set(AutomatedComment);
            invoiceSummaryPage.ApproveInvoiceDialog.Approve();
            var selectedTab = invoiceListPage.Tabs.GetActiveTab();
            Assert.AreEqual(selectedTab, MyInvoicesTab, $"Default selected tab should be {MyInvoicesTab} but is {selectedTab}");
            Assert.IsNull(myInvoicesList.GetInvoiceListItemFromText(selectedInvoiceName, false));

            // verify invoice summary reject
            selectedInvoice = myInvoicesList.GetInvoiceListItemByIndex(GetRandomNumber(myInvoices - 1));
            selectedInvoiceName = selectedInvoice.PrimaryElement.Text;
            selectedInvoice.Open();

            invoiceSummaryPage.Reject();
            invoiceSummaryPage.RejectInvoiceDialog.Controls["Reject Reason Codes"].Set("Billing");
            invoiceSummaryPage.ApproveInvoiceDialog.Controls["Internal Comment"].Set(AutomatedComment);
            invoiceSummaryPage.ApproveInvoiceDialog.Controls["External Comment"].Set(AutomatedComment);
            invoiceSummaryPage.ApproveInvoiceDialog.Reject();
            selectedTab = invoiceListPage.Tabs.GetActiveTab();
            Assert.AreEqual(selectedTab, MyInvoicesTab, $"Default selected tab should be {MyInvoicesTab} but is {selectedTab}");
            Assert.IsNull(myInvoicesList.GetInvoiceListItemFromText(selectedInvoiceName, false));
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC_16726 : Verify to make an adjustment on Line Item")]
        public void InvoiceLineItemAdjustment()
        {
            const string LineItemTab = "Line Items";
            const int Delta = 1;
            var invoiceListPage = _outlook.Oc.InvoicesListPage;
            var myInvoicesList = invoiceListPage.ItemList;
            var invoiceSummaryPage = _outlook.Oc.InvoiceSummaryPage;

            invoiceListPage.Open();

            var myInvoices = myInvoicesList.GetCount();
            var selectedInvoice = myInvoicesList.GetInvoiceListItemByIndex(GetRandomNumber(myInvoices - 1));
            var selectedInvoiceNumber = selectedInvoice.Number;

            // verify line items in invoice summary
            selectedInvoice.Open();
            invoiceSummaryPage.SummaryPanel.Toggle();
            var selectedTab = invoiceSummaryPage.Tabs.GetActiveTab();
            Assert.AreEqual(selectedTab, LineItemTab, $"Default selected tab should be {LineItemTab} but is {selectedTab}");

            // select line item
            var selectedLineItem = myInvoicesList.GetInvoiceLineItemByIndex(0);
            var lineItemType = selectedLineItem.PrimaryElement.Text;
            var lineItemNetTotal = selectedLineItem.NetTotal;

            // get header adjustment value
            invoiceSummaryPage.Tabs.Open("Header Adjustments");
            var selectedHeaderItem = myInvoicesList.GetInvoiceHeaderItemFromText(lineItemType);
            var headerItemNetAmount = selectedHeaderItem.NetAmount;

            // make adjustment
            invoiceSummaryPage.Tabs.Open("Line Items");
            selectedLineItem = myInvoicesList.GetInvoiceLineItemByIndex(0);
            selectedLineItem.Edit();
            invoiceSummaryPage.AdjustLineItemDialog.Controls["Adjustment Type"].Set("Increase by amount");
            invoiceSummaryPage.AdjustLineItemDialog.Controls["Adjustment Value"].Set(Delta.ToString());
            invoiceSummaryPage.AdjustLineItemDialog.Controls["Adjustment Reason"].Set("Billing for Indexing");
            invoiceSummaryPage.AdjustLineItemDialog.Controls["Adjustment Description"].Set(AutomatedComment);

            invoiceSummaryPage.AdjustLineItemDialog.Save();
            selectedLineItem = myInvoicesList.GetInvoiceLineItemByIndex(0);
            var updatedLineItemNetTotal = selectedLineItem.NetTotal;

            Assert.AreNotSame(updatedLineItemNetTotal, lineItemNetTotal, "Line Item is not updated");

            // verify header adjustments
            invoiceSummaryPage.Tabs.Open("Header Adjustments");
            var headerItem = invoiceSummaryPage.ItemList.GetInvoiceHeaderItemFromText(lineItemType);
            var updatedHeaderAdjustment = headerItem.AdjustmentAmount;
            var updatedHeaderItemNetAmount = headerItem.NetAmount;
            Assert.AreNotSame(updatedHeaderItemNetAmount, headerItemNetAmount, "Header Item net amount amount is not updated");
            Assert.Zero(updatedHeaderAdjustment);

            // verify invoice totals
            invoiceSummaryPage.InvoiceTotalsPanel.Toggle();
            var invoiceSummaryTotal = invoiceSummaryPage.InvoiceTotalValue;
            var updatedInvoiceHeader = invoiceSummaryPage.InvoiceTotalsPanel.HeaderValue;
            Assert.AreNotSame(GetNumeral(updatedInvoiceHeader), invoiceSummaryTotal, "Invoice Totals section is not updated");

            // verify invoice net total in invoice list
            invoiceListPage.Open();
            selectedInvoice = myInvoicesList.GetInvoiceListItemFromText(selectedInvoiceNumber);
            Assert.That(selectedInvoice.TotalNetAmount, Is.EqualTo(updatedInvoiceHeader).Using(new CurrencyComparer()));

            // verify invoice net total in passport
            selectedInvoice.AccessInvoice();
            var invoicePassportPage = _outlook.Oc.InvoicePassportPage;
            Assert.That(invoicePassportPage.GetNetTotal(), Is.EqualTo(updatedInvoiceHeader).Using(new CurrencyComparer()));
            invoicePassportPage.CloseWindowHandleSwitchToOc();

            selectedInvoice = myInvoicesList.GetInvoiceListItemFromText(selectedInvoiceNumber);
            selectedInvoice.Open();

            // reject line item
            selectedLineItem = myInvoicesList.GetInvoiceLineItemByIndex(0);
            selectedLineItem.Reject();
            invoiceSummaryPage.RejectLineItemDialog.Controls["Reason Code"].Set("Billing for Indexing");
            invoiceSummaryPage.RejectLineItemDialog.Controls["Internal Comment"].Set(AutomatedComment);
            invoiceSummaryPage.RejectLineItemDialog.Controls["External Comment"].Set(AutomatedComment);
            invoiceSummaryPage.RejectLineItemDialog.Save();

            selectedLineItem = myInvoicesList.GetInvoiceLineItemByIndex(0);
            var rejectedLineItemNetTotal = selectedLineItem.NetTotal;
            Assert.Zero(rejectedLineItemNetTotal);

            // verify header adjustments
            invoiceSummaryPage.Tabs.Open("Header Adjustments");
            headerItem = invoiceSummaryPage.ItemList.GetInvoiceHeaderItemFromText(lineItemType);
            Assert.AreNotSame(headerItem.NetAmount, updatedHeaderItemNetAmount);

            // verify invoice totals
            var recentInvoiceHeaderValue = invoiceSummaryPage.InvoiceTotalsPanel.HeaderValue;
            Assert.AreNotSame(GetNumeral(recentInvoiceHeaderValue), GetNumeral(updatedInvoiceHeader));

            // verify invoice net total in invoice list
            invoiceListPage.Open();
            selectedInvoice = myInvoicesList.GetInvoiceListItemFromText(selectedInvoiceNumber);
            Assert.That(selectedInvoice.TotalNetAmount, Is.EqualTo(recentInvoiceHeaderValue).Using(new CurrencyComparer()));

            // verify invoice net total in passport
            selectedInvoice.AccessInvoice();
            Assert.That(invoicePassportPage.GetNetTotal(), Is.EqualTo(recentInvoiceHeaderValue).Using(new CurrencyComparer()));
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC_16727 : Verify to make an adjustment on Header Item")]
        public void InvoiceHeaderAdjustment()
        {
            const int Delta = 1;
            var invoiceListPage = _outlook.Oc.InvoicesListPage;
            var myInvoicesList = invoiceListPage.ItemList;
            var invoiceSummaryPage = _outlook.Oc.InvoiceSummaryPage;

            invoiceListPage.Open();
            var myInvoices = myInvoicesList.GetCount();
            var selectedInvoice = myInvoicesList.GetInvoiceListItemByIndex(GetRandomNumber(myInvoices - 1));

            // Navigate to Header Adjustment tab
            selectedInvoice.Open();
            invoiceSummaryPage.Tabs.Open("Header Adjustments");
            invoiceSummaryPage.SummaryPanel.Toggle();

            var selectedHeaderItem = myInvoicesList.GetInvoiceHeaderItemByIndex(0);
            var headerItemType = selectedHeaderItem.PrimaryElement.Text;
            var headerItemNetAmount = selectedHeaderItem.NetAmount;

            // Make an adjustment
            selectedHeaderItem.Edit();
            invoiceSummaryPage.HeaderAdjustmentDialog.Controls["Adjustment Type"].Set("Increase by percentage");
            invoiceSummaryPage.HeaderAdjustmentDialog.Controls["Adjustment Value"].Set(Delta.ToString());
            invoiceSummaryPage.HeaderAdjustmentDialog.Controls["Adjustment Reason"].Set("Billing for Indexing");
            invoiceSummaryPage.HeaderAdjustmentDialog.Controls["Adjustment Description"].Set(AutomatedComment);
            invoiceSummaryPage.AdjustLineItemDialog.Save();

            // Verify net total amount get changed after an adjustment
            selectedHeaderItem = myInvoicesList.GetInvoiceHeaderItemByIndex(0);
            var updatedHeaderItemNetAmount = selectedHeaderItem.NetAmount;
            Assert.AreNotSame(updatedHeaderItemNetAmount, headerItemNetAmount, "Net Total value not being updated");

            // verify invoice net total with passport net total value
            invoiceSummaryPage.InvoiceTotalsPanel.Toggle();
            var invoiceTotalInOc = invoiceSummaryPage.InvoiceNetTotalValue;
            invoiceSummaryPage.AccessInvoice();
            var invoicePassportPage = _outlook.Oc.InvoicePassportPage;
            Assert.AreEqual(GetNumeral(invoicePassportPage.GetNetTotal()), invoiceTotalInOc);

            // Verify header adjustment net value with passport
            invoicePassportPage.NavigateToHeaderAdjustmentTab();
            var headerItemFromPassport = invoicePassportPage.PassportHeaderItemList.GetPassportListItemFromText(headerItemType);
            var headerItemNetTotalFromPassport = headerItemFromPassport.NetAmount;
            Assert.AreEqual(headerItemNetTotalFromPassport, updatedHeaderItemNetAmount);
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC_16732 : Validate Document Operations in Context Menu (CheckIn/CheckOut/Discard CheckOut)")]
        public void ValidateDocumentOperationFromContextMenu()
        {
            var checkedIn = CheckInStatus.CheckedIn.ToLower();
            var checkedOut = CheckInStatus.CheckedOut.ToLower();

            var invoiceListPage = _outlook.Oc.InvoicesListPage;
            var myInvoicesList = invoiceListPage.ItemList;
            var invoiceSummaryPage = _outlook.Oc.InvoiceSummaryPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;

            invoiceListPage.Open();
            Assert.GreaterOrEqual(myInvoicesList.GetCount(), 1);
            myInvoicesList.OpenRandom();
            invoiceSummaryPage.EntityTabs.Open("Documents");
            invoiceSummaryPage.CollapseInvoiceSummary();

            var wordFile = CreateDocument(OfficeApp.Word);
            DragAndDrop.FromFileSystem(wordFile, documentsListPage.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();
            invoiceSummaryPage.QuickSearch.SearchBy(wordFile.Name);
            var uploadedDocument = documentsListPage.ItemList.GetMatterDocumentListItemFromText(wordFile.Name);
            var documentStatus = uploadedDocument.Status.ToLower();
            Assert.AreEqual(checkedIn, documentStatus);

            //check out document
            uploadedDocument.FileOptions.CheckOut();
            _word = new Word(TestEnvironment);
            _word.Attach(wordFile.Name);
            const string EditedContent = "Content is edited by automated test.";
            _word.ReplaceTextWith(EditedContent);
            _word.SaveDocument();
            _word.Close();
            uploadedDocument = documentsListPage.ItemList.GetMatterDocumentListItemFromText(wordFile.Name);
            documentStatus = uploadedDocument.Status.ToLower();
            Assert.AreEqual(checkedOut, documentStatus);

            //check in document
            uploadedDocument.FileOptions.CheckIn();
            invoiceSummaryPage.Dialog.UploadDocument();
            uploadedDocument = documentsListPage.ItemList.GetMatterDocumentListItemFromText(wordFile.Name);
            documentStatus = uploadedDocument.Status.ToLower();
            Assert.AreEqual(checkedIn, documentStatus);

            //check out again for discard check out.
            uploadedDocument.FileOptions.CheckOut();
            _word = new Word(TestEnvironment);
            _word.Attach(wordFile.Name);
            _word.SaveDocument();
            _word.Close();
            uploadedDocument = documentsListPage.ItemList.GetMatterDocumentListItemFromText(wordFile.Name);
            documentStatus = uploadedDocument.Status.ToLower();
            Assert.AreEqual(checkedOut, documentStatus);

            //discard check out
            uploadedDocument.FileOptions.DiscardCheckOutAndRemoveLocalCopy();
            uploadedDocument = documentsListPage.ItemList.GetMatterDocumentListItemFromText(wordFile.Name);
            documentStatus = uploadedDocument.Status.ToLower();
            Assert.AreEqual(checkedIn, documentStatus);
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC_16673 : Validate Checkin & Checkout Document in Context Menu : Step 1, 2, 3")]
        public void ValidateCheckInCheckOutDocumentOperationFromContextMenu()
        {
            var checkedIn = CheckInStatus.CheckedIn.ToLower();
            var invoiceListPage = _outlook.Oc.InvoicesListPage;
            var myInvoicesList = invoiceListPage.ItemList;
            var invoiceSummaryPage = _outlook.Oc.InvoiceSummaryPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var myDocumentList = documentsListPage.ItemList;
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var checkInDocumentDialog = globalDocumentsPage.CheckInDocumentDialog;

            invoiceListPage.Open();
            Assert.GreaterOrEqual(myInvoicesList.GetCount(), 1);

            myInvoicesList.OpenRandom();
            invoiceSummaryPage.EntityTabs.Open("Documents");

            var wordFile = CreateDocument(OfficeApp.Word);
            DragAndDrop.FromFileSystem(wordFile, documentsListPage.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();

            invoiceSummaryPage.QuickSearch.SearchBy(wordFile.Name);
            var uploadedDocument = myDocumentList.GetInvoiceDocumentListItemFromText(wordFile.Name);
            var documentStatus = uploadedDocument.Status.ToLower();
            Assert.AreEqual(checkedIn, documentStatus);

            // Checkout
            var randomString = GetRandomText(10);
            uploadedDocument.FileOptions.CheckOut();
            _word = new Word(TestEnvironment);
            _word.Attach(wordFile.Name);
            Assert.True(_word.IsDocumentOpened);
            Assert.False(_word.IsReadOnly);

            _word.ReplaceTextWith(randomString);
            _word.SaveDocument();
            _word.Close();

            // Checkin
            uploadedDocument = myDocumentList.GetInvoiceDocumentListItemFromText(wordFile.Name);
            uploadedDocument.FileOptions.CheckIn();
            checkInDocumentDialog.Controls["Comments"].Set("Test Document : CheckIn Operation");
            checkInDocumentDialog.UploadDocument();

            // Search and Checkout again

            uploadedDocument = myDocumentList.GetInvoiceDocumentListItemFromText(wordFile.Name);
            documentStatus = uploadedDocument.Status.ToLower();
            Assert.AreEqual(checkedIn, documentStatus);

            // Clean Up

            invoiceSummaryPage.QuickSearch.SearchBy(wordFile.Name);

            if (uploadedDocument.Status == CheckInStatus.CheckedOut)
            {
                uploadedDocument = invoiceListPage.ItemList.GetInvoiceDocumentListItemFromText(wordFile.Name);
                uploadedDocument.FileOptions.CheckIn();
                checkInDocumentDialog.Controls["Comments"].Set("Clean up Operation");
                checkInDocumentDialog.UploadDocument();
            }
            uploadedDocument.Delete().Confirm();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC_16673 : Validate Keep and Discard Operations while Checkin Document in Context Menu)): #Steps 3,4,8,9,10 ")]
        public void VerifyKeepDiscardCheckinScenarios()
        {
            var checkedIn = CheckInStatus.CheckedIn.ToLower();
            var inVoiceListPage = _outlook.Oc.InvoicesListPage;
            var myInvoiceList = inVoiceListPage.ItemList;
            var invoiceSummaryPage = _outlook.Oc.InvoiceSummaryPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var myDocumentList = documentsListPage.ItemList;
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var checkOutDocumentDialog = globalDocumentsPage.CheckOutDocumentDialog;
            var checkInDocumentDialog = globalDocumentsPage.CheckInDocumentDialog;

            inVoiceListPage.Open();
            Assert.GreaterOrEqual(myInvoiceList.GetCount(), 1);

            myInvoiceList.OpenFirst();
            invoiceSummaryPage.EntityTabs.Open("Documents");

            var wordFile = CreateDocument(OfficeApp.Word);
            DragAndDrop.FromFileSystem(wordFile, documentsListPage.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();
            invoiceSummaryPage.QuickSearch.SearchBy(wordFile.Name);
            var uploadedDocument = myDocumentList.GetMatterDocumentListItemFromText(wordFile.Name);
            var documentStatus = uploadedDocument.Status.ToLower();
            Assert.AreEqual(checkedIn, documentStatus);

            // Checkout , Modify and Save
            var randomString = GetRandomText(10);
            uploadedDocument.FileOptions.CheckOut();
            _word = new Word(TestEnvironment);
            _word.Attach(wordFile.Name);
            Assert.True(_word.IsDocumentOpened);
            Assert.False(_word.IsReadOnly);
            _word.ReplaceTextWith(randomString);
            _word.SaveDocument();
            var localFilePath = _word.GetActiveFilePath();
            _word.Close();

            // Discard Checkout and Keep : Step 4
            uploadedDocument = myDocumentList.GetMatterDocumentListItemFromText(wordFile.Name);
            uploadedDocument.FileOptions.DiscardCheckOut();
            checkInDocumentDialog.Keep();

            // Verify old Text after Keeping
            _word = new Word(TestEnvironment);
            var localFileContent = _word.ReadWordContent(localFilePath);
            Assert.AreEqual(localFileContent.ToLower(), randomString.ToLower());

            // Checkout again to verify discard and remove : Step 3
            randomString = GetRandomText(10);
            uploadedDocument.FileOptions.CheckOut();
            checkOutDocumentDialog.Overwrite();
            _word = new Word(TestEnvironment);
            _word.Attach(wordFile.Name);
            Assert.True(_word.IsDocumentOpened);
            Assert.False(_word.IsReadOnly);
            _word.ReplaceTextWith(randomString);
            _word.SaveDocument();

            localFilePath = _word.GetActiveFilePath();
            _word.Close();

            uploadedDocument = myDocumentList.GetMatterDocumentListItemFromText(wordFile.Name);
            uploadedDocument.FileOptions.DiscardCheckOutAndRemoveLocalCopy();

            Assert.IsFalse(File.Exists(localFilePath));

            // Clean Up

            uploadedDocument = myDocumentList.GetMatterDocumentListItemFromText(wordFile.Name);
            if (uploadedDocument.Status == CheckInStatus.CheckedOut)
            {
                uploadedDocument.FileOptions.CheckIn();

                checkInDocumentDialog.Controls["Comments"].Set("Clean up Operation");

                checkInDocumentDialog.UploadDocument();
            }
            uploadedDocument.Delete().Confirm();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC 16672 : Verify Sort Operations in Invoice documents list End to End)")]
        public void VerifyDefaultSortOptionsInInvoiceDocuments()
        {
            const int uniqueDocumentsAndFoldersCount = 2;
            var folderNames = new string[] { };
            IDictionary<string, string> emails = new Dictionary<string, string>();
            var invoicesListPage = _outlook.Oc.InvoicesListPage;
            var myInvoicesList = invoicesListPage.ItemList;
            var invoiceSummaryPage = _outlook.Oc.InvoiceSummaryPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;

            invoicesListPage.Open();
            var myInvoices = myInvoicesList.GetAllInvoiceListItems();
            Assert.Greater(myInvoices.Count, 0, "Invoices list is not loaded or has 0 items");
            Assert.IsFalse(_outlook.Oc.IsErrorDisplayed(), "Invoices list is loaded with error message");
            myInvoicesList.OpenRandom();
            invoiceSummaryPage.EntityTabs.Open("Documents");
            Assert.IsTrue(documentsListPage.ItemList.IsSortIconVisible, "Sort Icon is not visible");
            var invoiceDocumentItems = myInvoicesList.GetAllInvoiceDocumentListItems();

            if (invoiceDocumentItems.Where(x => x.IsFolder).GroupBy(x => x.FolderName).Count() < uniqueDocumentsAndFoldersCount)
            {
                // Create folders
                folderNames = new[] { GetRandomText(5), GetRandomText(10) };

                foreach (var folderName in folderNames)
                {
                    myInvoicesList.OpenAddFolderDialog();
                    documentsListPage.AddFolderDialog.Controls["Name"].Set(folderName);
                    documentsListPage.AddFolderDialog.Save();
                }
            }

            if (invoiceDocumentItems.Where(x => !x.IsFolder).GroupBy(x => x.Name).Count() < uniqueDocumentsAndFoldersCount)
            {
                // Upload documents
                emails = _outlook.AddTestEmailsToFolder(uniqueDocumentsAndFoldersCount, FileSize.VerySmall, true);
                _outlook.OpenTestEmailFolder();
                _outlook.TurnOnReadingPane();

                for (var i = 0; i < emails.Count; i++)
                {
                    _outlook.SelectNthItem(i);
                    var filename = new FileInfo(emails.ElementAt(i).Value).Name;
                    var attachment = _outlook.GetAttachmentFromReadingPane(filename);
                    DragAndDrop.FromElementToElement(attachment, documentsListPage.DropPoint.GetElement());
                    documentsListPage.AddDocumentDialog.UploadDocument();

                    var uploadedDocument = myInvoicesList.GetInvoiceDocumentListItemFromText(filename);
                    Assert.IsNotNull(uploadedDocument);
                }
            }

            // Verify scenario
            documentsListPage.DocumentSortDialog.Sort("Document File Name", SortOrder.Descending);
            invoiceDocumentItems = myInvoicesList.GetAllInvoiceDocumentListItems();

            Assert.That(invoiceDocumentItems.Count(x => x.IsFolder), Is.GreaterThanOrEqualTo(uniqueDocumentsAndFoldersCount),
                "There are not enough folders to verify folders sorting. Need two folders at least.");
            Assert.That(invoiceDocumentItems.Where(x => !x.IsFolder).GroupBy(x => x.Name).Count(), Is.GreaterThanOrEqualTo(uniqueDocumentsAndFoldersCount),
                "There are no documents on the list. Need two documents to check sorting.");
            Assert.That(invoiceDocumentItems, Is.Ordered.Ascending.By(nameof(InvoiceDocumentListItem.IsFolder)).Using(new BooleanInverterComparer())
                .Then.Descending.By(nameof(InvoiceDocumentListItem.Name)));

            documentsListPage.DocumentSortDialog.RestoreSortDefaults();

            invoiceDocumentItems = myInvoicesList.GetAllInvoiceDocumentListItems();
            Assert.That(invoiceDocumentItems, Is.Ordered.Ascending.By(nameof(InvoiceDocumentListItem.IsFolder)).Using(new BooleanInverterComparer())
                .Then.Ascending.By(nameof(MatterDocumentListItem.Name)));
            Assert.That(invoiceDocumentItems.Where(x => x.IsFolder).ToList(), Is.Ordered.Ascending.By(nameof(InvoiceDocumentListItem.FolderName)));

            // Cleanup
            for (var i = 0; i < emails.Count; i++)
            {
                var filename = new FileInfo(emails.ElementAt(i).Value).Name;
                var document = myInvoicesList.GetInvoiceDocumentListItemFromText(filename);
                document.Delete().Confirm();
            }

            foreach (var folderName in folderNames)
            {
                myInvoicesList.GetEmailListItemFromText(folderName).Delete().Confirm();
            }
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC 16672 : Verify Sort Operations in Invoice documents list End to End)")]
        public void VerifySortOperationsInInvoiceDocuments()
        {
            var invoicesListPage = _outlook.Oc.InvoicesListPage;
            var myInvoicesList = invoicesListPage.ItemList;
            var invoiceSummaryPage = _outlook.Oc.InvoiceSummaryPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var sortOptions = new Dictionary<string, string>()
            {
                 { "Is Folder",  nameof(InvoiceDocumentListItem.IsFolder) },
                 { "Name", nameof(InvoiceDocumentListItem.Name) },
                 { "Last Modified By Full Name", nameof(InvoiceDocumentListItem.LastModifiedBy) },
                 { "Document File Name", nameof(InvoiceDocumentListItem.DocumentFileName) },
                 { "Updated At", nameof(InvoiceDocumentListItem.UpdatedAt) },
                 { "Document Size", nameof(InvoiceDocumentListItem.DocumentSize) },
                 { "Status", nameof(InvoiceDocumentListItem.Status) }
            };

            invoicesListPage.Open();
            var myInvoicesCount = myInvoicesList.GetCount();
            Assert.GreaterOrEqual(myInvoicesCount, 2, "Invoices list is not loaded or has less than two items");
            Assert.IsFalse(_outlook.Oc.IsErrorDisplayed(), "Invoices list is loaded with error message");

            myInvoicesList.OpenFirst();
            invoiceSummaryPage.EntityTabs.Open("Documents");
            Assert.IsTrue(documentsListPage.ItemList.IsSortIconVisible, "Sort Icon is not visible");

            foreach (var sortOption in sortOptions)
            {
                documentsListPage.DocumentSortDialog.Sort(sortOption.Key, SortOrder.Descending);

                var invoiceDocumentItems = myInvoicesList.GetAllInvoiceDocumentListItems();
                Assert.NotNull(invoiceDocumentItems.Select(x => x.GetType().GetProperty(sortOption.Value)));
                Assert.That(invoiceDocumentItems.Where(x => !x.IsFolder).ToList(), Is.Ordered.Descending.By(sortOption.Value));

                documentsListPage.DocumentSortDialog.Sort(sortOption.Key, SortOrder.Ascending);

                invoiceDocumentItems = myInvoicesList.GetAllInvoiceDocumentListItems();
                Assert.NotNull(invoiceDocumentItems.Select(x => x.GetType().GetProperty(sortOption.Value)));
                Assert.That(invoiceDocumentItems.Where(x => !x.IsFolder).ToList(), Is.Ordered.Ascending.By(sortOption.Value));
            }

            invoicesListPage.InvoiceSortDialog.RestoreSortDefaults();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC 16669 : Verify Sort Options in Invoice documents list (UI Verification))")]
        public void VerifySortOptionsInInvoiceDocumentsList()
        {
            var invoicesListPage = _outlook.Oc.InvoicesListPage;
            var myInvoicesList = invoicesListPage.ItemList;
            var invoiceSummaryPage = _outlook.Oc.InvoiceSummaryPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var expectedDocumentSortOptions = new[] { "Is Folder", "Name", "Created At", "Created By Full Name", "Document File Name", "Comments", "Updated At", "Document Size", "Status", "Last Modified By Full Name", "Locked By", "Created By" };

            invoicesListPage.Open();
            var myInvoices = myInvoicesList.GetAllInvoiceListItems();

            Assert.GreaterOrEqual(myInvoices.Count, 1, "Invoices list is not loaded or has 0 items");
            Assert.IsFalse(_outlook.Oc.IsErrorDisplayed(), "Invoices list is loaded with error message");

            myInvoicesList.OpenRandom();
            invoiceSummaryPage.EntityTabs.Open("Documents");

            Assert.IsTrue(documentsListPage.ItemList.IsSortIconVisible, "Sort Icon is not visible");

            documentsListPage.DocumentSortDialog.OpenSortDialog();
            var documentSortOptions = documentsListPage.DocumentSortDialog.GetSortOptions();

            CollectionAssert.IsSubsetOf(documentSortOptions, expectedDocumentSortOptions, "Sort options are not displayed");
            Assert.AreEqual(expectedDocumentSortOptions.Length, documentSortOptions.Length, "Difference in Sort options Count");

            CollectionAssert.AreEquivalent(expectedDocumentSortOptions, documentSortOptions, "Sort Options are not displayed as expected");
            CollectionAssert.AreEqual(expectedDocumentSortOptions, documentSortOptions);

            documentsListPage.DocumentSortDialog.CloseSortDialog();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC_16673 : Validate Checkin Error Scenario in Document Operations in Context Menu)): Step 5, 6, 7")]
        public void VerifyCheckInErrorScenario()
        {
            var checkedIn = CheckInStatus.CheckedIn.ToLower();
            var invoiceListPage = _outlook.Oc.InvoicesListPage;
            var myInvoicesList = invoiceListPage.ItemList;
            var invoiceSummaryPage = _outlook.Oc.InvoiceSummaryPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var myDocumentList = documentsListPage.ItemList;
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var checkInDocumentDialog = globalDocumentsPage.CheckInDocumentDialog;

            invoiceListPage.Open();
            Assert.GreaterOrEqual(myInvoicesList.GetCount(), 1);

            myInvoicesList.OpenRandom();
            invoiceSummaryPage.EntityTabs.Open("Documents");

            var wordFile = CreateDocument(OfficeApp.Word);
            DragAndDrop.FromFileSystem(wordFile, documentsListPage.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();

            invoiceSummaryPage.QuickSearch.SearchBy(wordFile.Name);
            var uploadedDocument = myDocumentList.GetMatterDocumentListItemFromText(wordFile.Name);
            var documentStatus = uploadedDocument.Status.ToLower();
            Assert.AreEqual(checkedIn, documentStatus);

            // Checkout
            var randomString = GetRandomText(10);
            uploadedDocument.FileOptions.CheckOut();
            _word = new Word(TestEnvironment);
            _word.Attach(wordFile.Name);
            Assert.True(_word.IsDocumentOpened);
            Assert.False(_word.IsReadOnly);

            _word.ReplaceTextWith(randomString);
            _word.SaveDocument();
            _word.Close();

            // Checkin
            uploadedDocument = myDocumentList.GetMatterDocumentListItemFromText(wordFile.Name);
            uploadedDocument.FileOptions.CheckIn();
            checkInDocumentDialog.Controls["Comments"].Set("Test Document : CheckIn Operation");
            checkInDocumentDialog.UploadDocument();

            // Search and Checkout again
            randomString = GetRandomText(10);
            invoiceSummaryPage.QuickSearch.SearchBy(wordFile.Name);
            uploadedDocument = myDocumentList.GetMatterDocumentListItemFromText(wordFile.Name);
            uploadedDocument.FileOptions.CheckOut();
            _word = new Word(TestEnvironment);
            _word.Attach(wordFile.Name);
            Assert.IsTrue(_word.IsDocumentOpened);
            Assert.IsFalse(_word.IsReadOnly);

            _word.ReplaceTextWith(randomString);
            var localFilePath = _word.GetActiveFilePath();
            _word.SaveDocument();
            _word.Close();

            var localFile = new FileInfo(localFilePath);
            var destinationFilePath = wordFile.FullName;
            wordFile.Delete();
            localFile.CopyTo(destinationFilePath);
            localFile.Delete();
            Assert.IsFalse(File.Exists(localFilePath));

            // Checkin
            uploadedDocument = myDocumentList.GetMatterDocumentListItemFromText(wordFile.Name);
            uploadedDocument.FileOptions.CheckIn();
            var checkInErrorDialog = globalDocumentsPage.CheckInErrorDialog;
            checkInErrorDialog.SelectFile(destinationFilePath);
            checkInDocumentDialog.Controls["Comments"].Set("Test Document : CheckIn successfully");
            checkInDocumentDialog.UploadDocument();

            // Clean Up
            uploadedDocument = myDocumentList.GetMatterDocumentListItemFromText(wordFile.Name);
            if (uploadedDocument.Status == CheckInStatus.CheckedOut)
            {
                uploadedDocument.FileOptions.CheckIn();
                checkInDocumentDialog.Controls["Comments"].Set("Clean up Operation");
                checkInDocumentDialog.UploadDocument();
            }
            uploadedDocument.Delete().Confirm();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC 16674 : Verify to Add a Document on Root and Folder lever and Perform Row level Operations")]
        public void DNDDocumentsOnMultipleLevel()
        {
            var invoicesListPage = _outlook.Oc.InvoicesListPage;
            var invoiceList = invoicesListPage.ItemList;
            var invoiceSummaryPage = _outlook.Oc.InvoiceSummaryPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var documentList = documentsListPage.ItemList;
            var documentSummaryPage = _outlook.Oc.DocumentSummaryPage;
            var checkInDialog = documentSummaryPage.CheckInDocumentDialog;
            var folderName = GetLongDateString();

            invoicesListPage.Open();
            var document = CreateDocument(OfficeApp.Notepad);
            var documentName = document.Name;
            var invoiceItem = invoiceList.GetInvoiceListItemByIndex(0);
            DragAndDrop.FromFileSystem(document, invoiceItem.DropPoint);
            checkInDialog.Controls["Comments"].Set("Document DND from FileSystem to OC-Document Summary");
            documentsListPage.AddDocumentDialog.UploadDocument();

            invoiceList.OpenFirst();
            invoicesListPage.Tabs.Open("Documents");
            documentsListPage.QuickSearch.SearchBy(documentName);
            var uploadedDocument = documentList.GetMatterDocumentListItemFromText(documentName);
            Assert.AreEqual(documentName, uploadedDocument.Name);

            //Cleanup
            uploadedDocument.Delete().Confirm();

            //Drag and Drop document on invoice Summary
            invoicesListPage.Open();
            invoiceList.OpenFirst();
            DragAndDrop.FromFileSystem(document, invoiceSummaryPage.DropPoint.GetElement());
            documentSummaryPage.AddDocumentDialog.UploadDocument();
            invoicesListPage.Tabs.Open("Documents");
            documentsListPage.QuickSearch.SearchBy(document.Name);
            uploadedDocument = documentList.GetMatterDocumentListItemFromText(document.Name);
            Assert.AreEqual(document.Name, uploadedDocument.Name);

            //Cleanup
            uploadedDocument.Delete().Confirm();

            //Drag and Drop document on folder level
            invoiceList.OpenAddFolderDialog();
            documentsListPage.AddFolderDialog.Controls["Name"].Set(folderName);
            documentsListPage.AddFolderDialog.Save();
            documentsListPage.QuickSearch.SearchBy(folderName);

            var invoiceListItem = invoiceList.GetInvoiceDocumentListItemByIndex(0);
            DragAndDrop.FromFileSystem(document, invoiceListItem.DropPoint);
            documentsListPage.AddDocumentDialog.UploadDocument();
            documentList.OpenFirst();
            documentsListPage.QuickSearch.SearchBy(document.Name);
            uploadedDocument = documentList.GetMatterDocumentListItemFromText(document.Name);
            Assert.AreEqual(document.Name, uploadedDocument.Name);

            //Cleanup
            documentsListPage.BreadcrumbsControl.NavigateToTheRoot();
            documentsListPage.QuickSearch.SearchBy(folderName);
            invoiceListItem = invoiceList.GetInvoiceDocumentListItemByIndex(0);
            invoiceListItem.Delete().Confirm();


        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC 16674 : Verify to Drag and Drop multiple documents on Invoice and Summary List. Verify to Drag and Drop Unsupported documents on Invoice and Summary List")]
        public void DNDMultipleAndUnsupportedDocuments()
        {
            var invoicesListPage = _outlook.Oc.InvoicesListPage;
            var invoiceList = invoicesListPage.ItemList;
            var invoiceSummaryPage = _outlook.Oc.InvoiceSummaryPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var documentList = documentsListPage.ItemList;

            invoicesListPage.Open();
            var notepadDocument = CreateDocument(OfficeApp.Notepad, GetRandomText(1024 * 3));
            var worddocument = CreateDocument(OfficeApp.Word, GetRandomText(1024 * 3));
            var path = Windows.GetWorkingTempFolder();
            var testDocuments = new ArrayList();
            testDocuments.Add(notepadDocument.Name);
            testDocuments.Add(worddocument.Name);
            Assert.NotNull(path);

            // Dnd Multiple Documents on Invoice List
            var invoice = invoiceList.GetInvoiceListItemByIndex(0);
            DragAndDrop.AllFilesInFolderDndOC(path, invoice.DropPoint);

            //Verify Uploaded Multiple Docs are Present and Delete the Docs
            invoiceList.OpenFirst();
            invoicesListPage.Tabs.Open("Documents");
            foreach (string testDocumentName in testDocuments)
            {
                documentsListPage.QuickSearch.SearchBy(testDocumentName);
                var uploadedDocument = documentList.GetMatterDocumentListItemFromText(testDocumentName);
                Assert.AreEqual(testDocumentName, uploadedDocument.Name);
                uploadedDocument.Delete().Confirm();
            }

            // Dnd Multiple Documents on Invoice Summary List
            DragAndDrop.AllFilesInFolderDndOC(path, invoiceSummaryPage.DropPoint.GetElement());


            //Verify Uploaded Multiple Docs are Present and Delete the Docs
            foreach (string testDocumentName in testDocuments)
            {
                documentsListPage.QuickSearch.SearchBy(testDocumentName);
                var uploadedDocument = documentList.GetMatterDocumentListItemFromText(testDocumentName);
                Assert.AreEqual(testDocumentName, uploadedDocument.Name);
                uploadedDocument.Delete().Confirm();
            }

            var unsupportedFile = CreateDocument(OfficeApp.Unsupported);
            invoicesListPage.Open();

            // Dnd unsuppored Documents on Invoice List
            invoice = invoiceList.GetInvoiceListItemByIndex(0);
            DragAndDrop.FromFileSystem(unsupportedFile, invoice.DropPoint);
            documentsListPage.AddDocumentDialog.UploadDocument();
            var ErrorMessage = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(1, ErrorMessage.Length);
            Assert.Contains(UnsupportedFileErrorMessage, ErrorMessage, "unsupportedFileMessage doesn't match");
            _outlook.Oc.CloseAllToastMessages();

            invoiceList.OpenFirst();

            // Dnd unsuppored Documents on Invoice List
            DragAndDrop.FromFileSystem(unsupportedFile, invoiceSummaryPage.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();
            ErrorMessage = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(1, ErrorMessage.Length);
            Assert.Contains(UnsupportedFileErrorMessage, ErrorMessage, "unsupportedFileMessage doesn't match");
            _outlook.Oc.CloseAllToastMessages();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("Test Case 16670 Verify Filter Options in Invoice documents list(UI Verification)")]
        public void UIVerifyFilterOptionsInInvoiceDocumentsList()
        {
            var invoicesListPage = _outlook.Oc.InvoicesListPage;
            var myInvoicesList = invoicesListPage.ItemList;
            var invoiceSummaryPage = _outlook.Oc.InvoiceSummaryPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var documentList = documentsListPage.ItemList;
            string[] expectedListOptions = { "Filter List", "Save Current View", "Set Current View as Default", "Restore Defaults", "Saved Views" };
            string[] expectedFilterDialogLabels = { "File Name (contains)", "Name (contains)", "Status", "Created By", "Updated By", "Created At (between)", "Comments (contains)", "Updated At (between)", "Content (contains)" };
            string[] expectedFilterDialogButtons = { "Apply", "Restore Defaults", "Cancel" };
            string[] expectedSaveCurrentViewDialogLabels = { "Create New", "Update Existing" };
            string[] expectedSaveCurrentViewDialogButtons = { "Save", "Cancel" };
            invoicesListPage.Open();
            var myInvoices = myInvoicesList.GetAllInvoiceListItems();

            Assert.Greater(myInvoices.Count, 0, "Invoices list is not loaded or has no items");
            Assert.IsFalse(_outlook.Oc.IsErrorDisplayed(), "Invoices list is loaded with error message");

            myInvoicesList.OpenRandom();
            invoiceSummaryPage.EntityTabs.Open("Documents");

            Assert.IsFalse(documentList.IsFilterIconVisible, "Filter Icon is visible");
            Assert.IsTrue(documentList.IsListOptionsDisplayed, "List Option Icon is not visible");

            // 2
            var listOptions = documentList.GetListOptionsMenu();
            foreach (var listOption in listOptions)
            {
                Assert.IsTrue(expectedListOptions.Contains(listOption));
            }
            // 3
            documentList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            var invoiceDocumentFilterDialog = documentsListPage.InvoiceDocumentListFilterDialog;
            var labelTexts = invoiceDocumentFilterDialog.GetAllLabelTexts();
            var dialogButtons = invoiceDocumentFilterDialog.GetDialogButtons();
            foreach (var label in labelTexts)
            {
                Assert.IsTrue(expectedFilterDialogLabels.Contains(label));
            }
            foreach (var button in dialogButtons)
            {
                Assert.IsTrue(expectedFilterDialogButtons.Contains(button));
            }
            invoiceDocumentFilterDialog.Cancel();

            // 4
            documentList.OpenListOptionsMenu().SaveCurrentView();
            var saveCurrentViewDialoglabels = documentsListPage.SaveCurrentViewDialog.GetAllLabelTexts();
            foreach (var label in expectedSaveCurrentViewDialogLabels)
            {
                Assert.IsTrue(saveCurrentViewDialoglabels.Contains(label));
            }
            var saveCurrentViewDialogButtons = documentsListPage.SaveCurrentViewDialog.GetDialogButtons();
            foreach (var buttons in saveCurrentViewDialogButtons)
            {
                Assert.IsTrue(expectedSaveCurrentViewDialogButtons.Contains(buttons));
            }
            documentsListPage.SaveCurrentViewDialog.Controls["Create New"].Set(LengthyViewName);
            documentsListPage.SaveCurrentViewDialog.Save();
            documentList.OpenListOptionsMenu().RestoreDefaults();

            // 5
            Assert.DoesNotThrow(() => documentList.OpenListOptionsMenu().HoverMouseOnSavedViewMenuOption());
            documentList.RandomOverlayClick();
            // 6
            Assert.DoesNotThrow(() => documentList.OpenListOptionsMenu().HoverMouseOnSavedViewByName(LengthyViewName));
            var viewNameTooltip = documentList.GetSavedViewNameToolTip(LengthyViewName);
            Assert.AreEqual(LengthyViewName, viewNameTooltip, "Mismatch in Tooltip and viewname");

            documentList.RandomOverlayClick();

            // Clean up : Delete Existing view
            documentList.OpenListOptionsMenu().RestoreDefaults();
            documentList.OpenListOptionsMenu().RemoveSavedView(LengthyViewName);
        }

        [TearDown]
        public void TearDown()
        {
            SaveScreenShotsAndLogs(_outlook);
            _outlook?.Destroy();
            _word?.Destroy();
        }
    }
}
