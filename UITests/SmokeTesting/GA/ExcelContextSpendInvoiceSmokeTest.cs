using NUnit.Framework;
using UITests.PageModel;
using UITests.PageModel.OfficeApps;
using static UITests.Constants;
using static UITests.TestHelpers;

namespace UITests.SmokeTesting.GA
{
    [TestFixture]
    internal class ExcelContextSpendInvoiceSmokeTest : UITestBase
    {
        private Excel _excel;

        [SetUp]
        public void SetUp()
        {
            _excel = new Excel(TestEnvironment);
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC_16674 : Verify to Add a Document on Root and Folder lever and Perform Row level Operations)")]
        public void QuickFileNewDocumentOnInvoiceSummaryAndDocumentFolder()
        {
            var quickFile = CreateDocument(OfficeApp.Excel);
            _excel.OpenDocumentFromExplorer(quickFile.FullName);
            _excel.CloseDocument();
            _excel.OpenNewExcel();
            _excel.AttachToOc();
            _excel.Oc.BasicSettingsPage.LogInAsAttorneyUser();

            var randomString = GetRandomText(10);
            var checkedIn = CheckInStatus.CheckedIn.ToLower();

            var invoiceListPage = _excel.Oc.InvoicesListPage;
            var invoiceList = invoiceListPage.ItemList;
            var invoiceSummaryPage = _excel.Oc.InvoiceSummaryPage;
            var documentsListPage = _excel.Oc.DocumentsListPage;
            var documentList = documentsListPage.ItemList;
            var globalDocumentsPage = _excel.Oc.GlobalDocumentsPage;
            var checkInDocumentDialog = globalDocumentsPage.CheckInDocumentDialog;
            var folderName = GetLongDateString();

            invoiceListPage.Open();
            invoiceList.OpenRandom();

            //Quickfile on invoice summary
            invoiceSummaryPage.QuickFile(false);
            invoiceSummaryPage.Dialog.SaveAndUpload(randomString, true);
            invoiceSummaryPage.Dialog.UploadDocument();
            invoiceSummaryPage.Tabs.Open("Documents");
            documentsListPage.QuickSearch.SearchBy(randomString);
            var uploadedDocument = documentList.GetInvoiceDocumentListItemFromText(randomString);
            var documentStatus = uploadedDocument.Status.ToLower();
            Assert.AreEqual(checkedIn, documentStatus);
            _excel.Attach(uploadedDocument.Name);
            _excel.CloseDocument();

            uploadedDocument = documentList.GetInvoiceDocumentListItemFromText(randomString);
            uploadedDocument.FileOptions.CheckOut();
            _excel = new Excel(TestEnvironment);
            _excel.Attach(uploadedDocument.Name);

            //Make some changes in excel and quickfile
            _excel.ReplaceTextWith(randomString);
            _excel.ClickTab();
            invoiceSummaryPage.QuickFile(false);
            invoiceSummaryPage.Dialog.SaveAs("New" + randomString, false);
            invoiceSummaryPage.Dialog.UploadDocument();

            //Create a New Folder and Quickfile on Folder
            documentList.OpenAddFolderDialog();
            documentsListPage.AddFolderDialog.Controls["Name"].Set(folderName);
            documentsListPage.AddFolderDialog.Save();
            documentsListPage.QuickSearch.SearchBy(folderName);
            var folder = documentList.GetInvoiceDocumentListItemFromText(folderName);
            folder.QuickFile();
            invoiceSummaryPage.Dialog.UploadDocument();
            invoiceList.Open();
            documentsListPage.QuickSearch.SearchBy(randomString);
            uploadedDocument = documentList.GetInvoiceDocumentListItemFromText(randomString);
            StringAssert.Contains(randomString, uploadedDocument.Name, "Document is not present");

            //Cleanup
            //Delete Created Folder
            invoiceSummaryPage.Tabs.Open("Documents");
            documentsListPage.QuickSearch.SearchBy(folderName);
            folder = invoiceList.GetInvoiceDocumentListItemFromText(folderName);
            folder.Delete().Confirm();

            //Delete files
            documentsListPage.QuickSearch.SearchBy(randomString);
            var testDocuments = documentList.GetAllInvoiceDocumentListItems();

            foreach (var testDocument in testDocuments)
            {
                documentsListPage.QuickSearch.SearchBy(testDocument.Name);
                uploadedDocument = documentList.GetInvoiceDocumentListItemFromText(testDocument.Name);
                if (uploadedDocument.Status == CheckInStatus.CheckedOut)
                {
                    uploadedDocument.FileOptions.CheckIn();
                    checkInDocumentDialog.Controls["Comments"].Set("Clean up Operation");
                    checkInDocumentDialog.UploadDocument();
                }
                uploadedDocument = documentList.GetInvoiceDocumentListItemFromText(testDocument.Name);
                uploadedDocument.Delete().Confirm();
            }
        }

        [TearDown]
        public void TearDown()
        {
            SaveScreenShotsAndLogs(_excel);
            _excel.Close();
            _excel.Destroy();
        }
    }
}
