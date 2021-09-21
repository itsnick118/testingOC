using System;
using System.IO;
using IntegratedDriver;
using NUnit.Framework;
using UITests.PageModel;
using UITests.PageModel.OfficeApps;
using static UITests.Constants;
using static UITests.TestHelpers;

namespace UITests.SmokeTesting.GA
{
    [TestFixture]
    public class WordContextSpendInvoiceSmokeTest : UITestBase
    {
        private Word _word;

        [SetUp]
        public void SetUp()
        {
            _word = new Word(TestEnvironment);
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC_16732 : Validate Document Operations in Context Menu (CheckIn/CheckOut/Discard CheckOut)")]
        public void VerifyQuickFileDocument()
        {
            var checkedIn = CheckInStatus.CheckedIn.ToLower();

            var quickFile = CreateDocument(OfficeApp.Word);
            _word.OpenDocumentFromExplorer(quickFile.FullName);
            _word.AttachToOc();
            _word.Oc.BasicSettingsPage.LogInAsAttorneyUser();
            var invoiceListPage = _word.Oc.InvoicesListPage;
            var myInvoicesList = invoiceListPage.ItemList;
            var invoiceSummaryPage = _word.Oc.InvoiceSummaryPage;
            var documentsListPage = _word.Oc.DocumentsListPage;
            invoiceListPage.Open();
            Assert.GreaterOrEqual(myInvoicesList.GetCount(), 1);

            myInvoicesList.OpenRandom();
            invoiceSummaryPage.EntityTabs.Open("Documents");
            invoiceSummaryPage.QuickFile();
            invoiceSummaryPage.Dialog.UploadDocument();
            invoiceSummaryPage.QuickSearch.SearchBy(quickFile.Name);
            var uploadedDocument = documentsListPage.ItemList.GetMatterDocumentListItemFromText(quickFile.Name);
            Assert.AreEqual(quickFile.Name, uploadedDocument.DocumentFileName);
            Assert.AreEqual(checkedIn, uploadedDocument.Status.ToLower());
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC_16673 : Validate Checkin Error Scenario in Document Quickfile)): Step 11, 12, 13")]
        public void VerifyCheckinScenarioinDocumentQuickFile()
        {
            var checkedIn = CheckInStatus.CheckedIn.ToLower();

            var quickFile = CreateDocument(OfficeApp.Word);
            _word.OpenDocumentFromExplorer(quickFile.FullName);
            _word.AttachToOc();
            _word.Oc.BasicSettingsPage.LogInAsAttorneyUser();

            var inVoiceListPage = _word.Oc.InvoicesListPage;
            var myInvoiceList = inVoiceListPage.ItemList;
            var invoiceSummaryPage = _word.Oc.InvoiceSummaryPage;
            var documentsListPage = _word.Oc.DocumentsListPage;
            var myDocumentList = documentsListPage.ItemList;
            var randomString = AutomatedComment + " " + GetRandomText(5);
            var globalDocumentsPage = _word.Oc.GlobalDocumentsPage;
            var checkInDocumentDialog = globalDocumentsPage.CheckInDocumentDialog;
            var guId = Guid.NewGuid();
            var newFileName = "New_" + quickFile.Name;
            inVoiceListPage.Open();
            Assert.GreaterOrEqual(myInvoiceList.GetCount(), 1);

            myInvoiceList.OpenFirst();
            invoiceSummaryPage.EntityTabs.Open("Documents");
            invoiceSummaryPage.QuickFile();
            invoiceSummaryPage.Dialog.UploadDocument();

            // Performing Quickfile for the first time

            invoiceSummaryPage.QuickSearch.SearchBy(quickFile.Name);
            var uploadedDocument = myDocumentList.GetMatterDocumentListItemFromText(quickFile.Name);
            Assert.AreEqual(quickFile.Name, uploadedDocument.DocumentFileName);
            Assert.AreEqual(checkedIn, uploadedDocument.Status.ToLower());
            _word.CloseDocument();

            // Checkout for modification
            uploadedDocument = myInvoiceList.GetMatterDocumentListItemFromText(quickFile.Name);
            uploadedDocument.FileOptions.CheckOut();
            _word = new Word(TestEnvironment);
            _word.Attach(quickFile.Name);
            Assert.True(_word.IsDocumentOpened);
            Assert.False(_word.IsReadOnly);
            _word.ReplaceTextWith(randomString);

            //Code to verify Save option

            invoiceSummaryPage.QuickFile(false);
            invoiceSummaryPage.Dialog.Save();
            invoiceSummaryPage.Dialog.Overwrite();
            invoiceSummaryPage.Dialog.UploadDocument();

            _word.CloseDocument();

            // Code to verify successful Checkin
            uploadedDocument = myDocumentList.GetMatterDocumentListItemFromText(quickFile.Name);
            var fileInfo = uploadedDocument.Download($"{guId}.tmp");
            var localFilePath = fileInfo.FullName;
            var localFileContent = _word.ReadWordContent(localFilePath);
            Assert.AreEqual(randomString.ToLower(), localFileContent.ToLower());

            //Code to verify checkin with 'Save As' option

            randomString = AutomatedComment + " " + GetRandomText(5);
            uploadedDocument = myInvoiceList.GetMatterDocumentListItemFromText(quickFile.Name);
            uploadedDocument.FileOptions.CheckOut();
            invoiceSummaryPage.Dialog.Overwrite();
            _word = new Word(TestEnvironment);
            _word.Attach(quickFile.Name);
            Assert.True(_word.IsDocumentOpened);
            Assert.False(_word.IsReadOnly);
            _word.ReplaceTextWith(randomString);

            invoiceSummaryPage.QuickFile(false);
            invoiceSummaryPage.Dialog.SaveAs(newFileName, false);

            invoiceSummaryPage.Dialog.UploadDocument();

            // Code to verify successful Checkin
            uploadedDocument = myDocumentList.GetMatterDocumentListItemFromText(newFileName);
            fileInfo = uploadedDocument.Download($"{guId}.docx");

            localFilePath = fileInfo.FullName;
            localFileContent = _word.ReadWordContent(localFilePath);
            Assert.AreEqual(randomString.ToLower(), localFileContent.ToLower());
            _word.SaveDocument();
            _word.CloseDocument();

            // Checkout for modification for Save with a new name
            uploadedDocument = myDocumentList.GetMatterDocumentListItemFromText(newFileName);
            uploadedDocument.FileOptions.CheckOut();
            invoiceSummaryPage.Dialog.Overwrite();
            _word = new Word(TestEnvironment);
            _word.Attach(quickFile.Name);
            Assert.True(_word.IsDocumentOpened);
            Assert.False(_word.IsReadOnly);
            _word.ReplaceTextWith(randomString);

            //Code to verify Save with a new name option

            invoiceSummaryPage.QuickFile(false);
            invoiceSummaryPage.Dialog.Save();
            invoiceSummaryPage.Dialog.SaveAs("NewNameOf_" + quickFile.Name, false, true);

            invoiceSummaryPage.Dialog.UploadDocument();

            // Code to verify successful Checkin
            uploadedDocument = myDocumentList.GetMatterDocumentListItemFromText("NewNameOf_" + quickFile.Name);
            fileInfo = uploadedDocument.Download($"{guId}.docx");
            localFilePath = fileInfo.FullName;
            localFileContent = _word.ReadWordContent(localFilePath);
            Assert.AreEqual(randomString.ToLower(), localFileContent.ToLower());

            //clean up
            _word.SaveDocument();
            _word.CloseDocument();

            Windows.ClearWorkingTempFolder();
            documentsListPage.QuickSearch.SearchBy(quickFile.Name);
            var testDocuments = myDocumentList.GetAllMatterDocumentListItems();

            foreach (var testDocument in testDocuments)
            {
                documentsListPage.QuickSearch.SearchBy(testDocument.Name);
                uploadedDocument = myDocumentList.GetMatterDocumentListItemFromText(testDocument.Name);
                if (uploadedDocument.Status == CheckInStatus.CheckedOut)
                {
                    uploadedDocument.FileOptions.CheckIn();
                    checkInDocumentDialog.Controls["Comments"].Set("Clean up Operation");
                    checkInDocumentDialog.UploadDocument();
                }
                uploadedDocument = myDocumentList.GetMatterDocumentListItemFromText(testDocument.Name);
                uploadedDocument.Delete().Confirm();
            }
            Windows.ClearWorkingTempFolder();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC_16674 : Verify to Add a Document on Root and Folder lever and Perform Row level Operations)")]
        public void QuickFileNewDocumentOnInvoiceList()
        {
            var quickFile = CreateDocument(OfficeApp.Word);
            _word.OpenDocumentFromExplorer(quickFile.FullName);
            _word.CloseDocument();
            _word.OpenNewWord();
            _word.AttachToOc();
            _word.Oc.BasicSettingsPage.LogInAsAttorneyUser();

            var wordDocName = GetRandomText(10) + ".docx";
            var invoiceListPage = _word.Oc.InvoicesListPage;
            var invoiceList = invoiceListPage.ItemList;
            var invoiceSummaryPage = _word.Oc.InvoiceSummaryPage;
            var documentsListPage = _word.Oc.DocumentsListPage;
            var documentList = documentsListPage.ItemList;
            var localFilePath = Path.Combine(Windows.GetWorkingTempFolder().FullName, wordDocName);

            invoiceListPage.Open();

            //QuickFile and Save the Doc to Local
            var invoice = invoiceList.GetInvoiceListItemByIndex(0);
            invoice.QuickFile();
            invoiceSummaryPage.Dialog.SaveAndUpload(wordDocName, true);
            Assert.IsFalse(File.Exists(localFilePath));
            invoiceSummaryPage.Dialog.UploadDocument();
            invoiceList.OpenFirst();
            invoiceSummaryPage.Tabs.Open("Documents");
            documentsListPage.QuickSearch.SearchBy(wordDocName);
            var uploadedDocument = documentList.GetInvoiceDocumentListItemFromText(wordDocName);
            Assert.AreEqual(wordDocName, uploadedDocument.Name);

            //Clean up
            _word.Attach(wordDocName);
            _word.CloseDocument();
            uploadedDocument.Delete().Confirm();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC 16674 : Verify to Drag and Drop multiple documents on Invoice and Summary List. Verify to Drag and Drop Unsupported documents on Invoice and Summary List")]
        public void ListPageOperationsViewDownloadDelete()
        {
            var quickFile = CreateDocument(OfficeApp.Word);
            _word.OpenDocumentFromExplorer(quickFile.FullName);
            _word.AttachToOc();
            _word.Oc.BasicSettingsPage.LogInAsAttorneyUser();

            var invoiceListPage = _word.Oc.InvoicesListPage;
            var invoiceList = invoiceListPage.ItemList;
            var invoiceSummaryPage = _word.Oc.InvoiceSummaryPage;
            var documentsListPage = _word.Oc.DocumentsListPage;
            var documentList = documentsListPage.ItemList;
            var globalDocumentsPage = _word.Oc.GlobalDocumentsPage;
            var checkInDocumentDialog = globalDocumentsPage.CheckInDocumentDialog;
            var checkedIn = CheckInStatus.CheckedIn.ToLower();

            invoiceListPage.Open();

            //Upload Document to Invoice List
            var wordDocument = CreateDocument(OfficeApp.Word);
            var docName = wordDocument.Name;
            var invoice = invoiceList.GetInvoiceListItemByIndex(0);
            DragAndDrop.FromFileSystem(wordDocument, invoice.DropPoint);
            documentsListPage.AddDocumentDialog.UploadDocument();

            //Navigate to the Uploaded Document
            invoiceList.OpenFirst();
            invoiceListPage.Tabs.Open("Documents");
            documentsListPage.QuickSearch.SearchBy(docName);
            var uploadedDocument = documentList.GetMatterDocumentListItemFromText(docName);

            var documentStatus = uploadedDocument.Status.ToLower();
            Assert.AreEqual(checkedIn, documentStatus);
            _word.CloseDocument();

            //Open the Uploaded Document
            uploadedDocument.Open();
            _word = new Word(TestEnvironment);
            _word.Attach(docName);
            Assert.IsNotNull(_word.GetReadOnlyLabel());
            uploadedDocument = documentList.GetMatterDocumentListItemFromText(docName);
            var fileOptionList = uploadedDocument.FileOptions.OpenFileOptionsAndGetOptions();
            Assert.IsTrue(fileOptionList[0].Enabled, "CheckedOut Button is found Disable");
            Assert.IsFalse(fileOptionList[1].Enabled, "CheckedIn Button is found Enable");
            Assert.IsFalse(fileOptionList[2].Enabled, "Discard CheckedOut Button is found Enable");

            //Download Document
            uploadedDocument = documentList.GetMatterDocumentListItemFromText(docName);
            var fileInfo = uploadedDocument.Download($"{Guid.NewGuid()}.docx");
            var localFilePath = fileInfo.FullName;
            var localFileContent = _word.ReadWordContent(localFilePath);
            Assert.AreEqual(Constants.InitialDefaultContent, localFileContent);

            //Rename Document
            uploadedDocument.Rename();
            var renameDocName = GetRandomText(6) + ".doc";
            documentsListPage.RenameDocumentDialog.Controls["Name"].Set(renameDocName);
            documentsListPage.RenameDocumentDialog.Controls["Document File Name"].Set(renameDocName);
            documentsListPage.AddFolderDialog.Save();

            documentsListPage.QuickSearch.SearchBy(renameDocName);
            var renamedDoc = documentList.GetMatterDocumentListItemFromText(renameDocName);
            Assert.IsNotNull(renamedDoc);
            _word.CloseDocument();

            // clean up
            uploadedDocument = documentList.GetMatterDocumentListItemFromText(renameDocName);
            uploadedDocument.Delete().Confirm();
            Assert.IsNull(documentList.GetMatterDocumentListItemFromText(renameDocName, false));

            //Upload New Document
            DragAndDrop.FromFileSystem(wordDocument, invoiceSummaryPage.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();
            documentsListPage.QuickSearch.SearchBy(docName);
            uploadedDocument = documentList.GetMatterDocumentListItemFromText(docName);
            documentStatus = uploadedDocument.Status.ToLower();
            Assert.AreEqual(checkedIn, documentStatus);
            _word.CloseDocument();

            //CheckOut Document
            uploadedDocument.FileOptions.CheckOut();
            uploadedDocument = documentList.GetMatterDocumentListItemFromText(docName);
            documentStatus = uploadedDocument.Status.ToLower();
            var checkedOut = CheckInStatus.CheckedOut.ToLower();
            Assert.AreEqual(checkedOut, documentStatus);
            _word.CloseDocument();


            //Open the Uploaded Document
            uploadedDocument.Open();
            _word = new Word(TestEnvironment);
            _word.Attach(docName);
            uploadedDocument = documentList.GetMatterDocumentListItemFromText(docName);
            fileOptionList = uploadedDocument.FileOptions.OpenFileOptionsAndGetOptions();
            Assert.IsFalse(fileOptionList[0].Enabled, "CheckedOut Button is found Enable");
            Assert.IsTrue(fileOptionList[1].Enabled, "CheckedIn Button is found Disable");
            Assert.IsTrue(fileOptionList[2].Enabled, "Discard CheckedOut Button is found Disable");
            documentsListPage.CheckInDocumentDialog.Cancel();

            //Download Document
            fileInfo = uploadedDocument.Download($"{Guid.NewGuid()}.docx");
            localFilePath = fileInfo.FullName;
            localFileContent = _word.ReadWordContent(localFilePath);
            Assert.AreEqual(Constants.InitialDefaultContent, localFileContent);

            //Rename Document
            _word.AttachToOc();
            uploadedDocument.Rename();
            var unsupportedFileMessage = _word.Oc.GetAllToastMessages();
            Assert.AreEqual(1, unsupportedFileMessage.Length);
            Assert.Contains(CheckedOutDocumentRenameErrorMessage, unsupportedFileMessage);
            _word.Oc.CloseAllToastMessages();

            // clean up
            uploadedDocument = documentList.GetMatterDocumentListItemFromText(docName);
            uploadedDocument.Delete().Confirm();
            var errorMessage = _word.Oc.GetAllToastMessages();
            Assert.AreEqual(1, errorMessage.Length);
            Assert.Contains(CheckedOutDocumentDeleteMessage, errorMessage);
            uploadedDocument.FileOptions.CheckIn();
            checkInDocumentDialog.Controls["Comments"].Set("Clean up Operation");
            checkInDocumentDialog.UploadDocument();
            Assert.IsNull(documentList.GetMatterDocumentListItemFromText(renameDocName, false));
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC 16675 :  Verify Invoice Document Summary, Breadcrumb & Version History)- Step 1 to 4")]
        public void VerifyInvoiceDocumentSummaryDragAndDrop()
        {
            var checkedIn = CheckInStatus.CheckedIn.ToLower();
            var quickFile = CreateDocument(OfficeApp.Word);
            _word.OpenDocumentFromExplorer(quickFile.FullName);
            _word.AttachToOc();
            _word.Oc.BasicSettingsPage.LogInAsAttorneyUser();

            var inVoiceListPage = _word.Oc.InvoicesListPage;
            var myInvoiceList = inVoiceListPage.ItemList;
            var invoiceSummaryPage = _word.Oc.InvoiceSummaryPage;
            var documentSummaryPage = _word.Oc.DocumentSummaryPage;
            var documentsListPage = _word.Oc.DocumentsListPage;
            var myDocumentList = documentsListPage.ItemList;

            inVoiceListPage.Open();
            Assert.GreaterOrEqual(myInvoiceList.GetCount(), 1);

            // Performing Quickfile for the first time
            myInvoiceList.OpenFirst();
            invoiceSummaryPage.EntityTabs.Open("Documents");
            invoiceSummaryPage.QuickFile();
            invoiceSummaryPage.Dialog.UploadDocument();

            // Drag and Drop another file : Step 3
            var wordFile = CreateDocument(OfficeApp.Word);
            DragAndDrop.FromFileSystem(wordFile, documentsListPage.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();

            var uploadedDocument = myDocumentList.GetInvoiceDocumentListItemFromText(wordFile.Name);
            Assert.AreEqual(wordFile.Name, uploadedDocument.DocumentFileName);
            Assert.AreEqual(checkedIn, uploadedDocument.Status.ToLower());

            // Drag and Drop - Step 4
            var similarWordFile = CreateDocument(OfficeApp.Word);
            uploadedDocument.NavigateToSummary();
            DragAndDrop.FromFileSystem(similarWordFile, documentsListPage.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.Proceed();
            documentsListPage.AddDocumentDialog.UploadDocument();
            documentSummaryPage.NavigateToParent();
            uploadedDocument = myDocumentList.GetInvoiceDocumentListItemFromText(wordFile.Name);
            Assert.AreEqual(wordFile.Name, uploadedDocument.DocumentFileName);
            Assert.AreEqual(checkedIn, uploadedDocument.Status.ToLower());

            // Clean up
            uploadedDocument.Delete().Confirm();
            Assert.IsNull(myDocumentList.GetInvoiceDocumentListItemFromText(wordFile.Name, false));
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("TC 16675 :  Verify Invoice Document Summary, Breadcrumb & Version History): Step 5")]
        public void VerifyOpenInvoiceDocumentFromVersionHistory()
        {
            var quickFile = CreateDocument(OfficeApp.Word);
            _word.OpenDocumentFromExplorer(quickFile.FullName);
            _word.AttachToOc();
            _word.Oc.BasicSettingsPage.LogInAsAttorneyUser();

            var inVoiceListPage = _word.Oc.InvoicesListPage;
            var myInvoiceList = inVoiceListPage.ItemList;
            var invoiceSummaryPage = _word.Oc.InvoiceSummaryPage;
            var documentSummaryPage = _word.Oc.DocumentSummaryPage;
            var documentsListPage = _word.Oc.DocumentsListPage;
            var myDocumentList = documentsListPage.ItemList;
            var globalDocumentsPage = _word.Oc.GlobalDocumentsPage;
            var checkInDocumentDialog = globalDocumentsPage.CheckInDocumentDialog;
            inVoiceListPage.Open();
            Assert.GreaterOrEqual(myInvoiceList.GetCount(), 1);

            // Performing Quickfile for the first time
            myInvoiceList.OpenFirst();
            invoiceSummaryPage.EntityTabs.Open("Documents");
            invoiceSummaryPage.QuickFile();
            invoiceSummaryPage.Dialog.UploadDocument();
            _word.CloseDocument();

            // Verify Read only label - Step 5
            invoiceSummaryPage.QuickSearch.SearchBy(quickFile.Name);
            var uploadedDocument = myDocumentList.GetInvoiceDocumentListItemFromText(quickFile.Name);

            //check out document
            uploadedDocument.FileOptions.CheckOut();
            const string EditedContent = "Content is edited by automated test.";
            _word.ReplaceTextWith(EditedContent);
            _word.SaveDocument();

            invoiceSummaryPage.QuickSearch.SearchBy(quickFile.Name);
            uploadedDocument = myDocumentList.GetInvoiceDocumentListItemFromText(quickFile.Name);

            //check in document
            uploadedDocument.FileOptions.CheckIn();
            invoiceSummaryPage.Dialog.UploadDocument();
            uploadedDocument = myDocumentList.GetInvoiceDocumentListItemFromText(quickFile.Name);

            uploadedDocument.NavigateToSummary();
            var file = myDocumentList.GetVersionHistoryListItemByIndex(1);
            file.Open();

            Assert.IsNull(_word.GetReadOnlyLabel());
            _word.CloseDocument();

            documentSummaryPage.NavigateToParent();
            _word.Oc.WaitForLoadComplete();
            invoiceSummaryPage.QuickSearch.SearchBy(quickFile.Name);
            uploadedDocument = myDocumentList.GetInvoiceDocumentListItemFromText(quickFile.Name);
            uploadedDocument.FileOptions.CheckOut();
            _word.CloseDocument();
            uploadedDocument = myDocumentList.GetInvoiceDocumentListItemFromText(quickFile.Name);
            uploadedDocument.NavigateToSummary();
            file = myDocumentList.GetVersionHistoryListItemByIndex(1);
            file.Open();
            Assert.IsNull(_word.GetReadOnlyLabel());
            documentSummaryPage.NavigateToParent();
            _word.Oc.WaitForLoadComplete();
            invoiceSummaryPage.QuickSearch.SearchBy(quickFile.Name);
            uploadedDocument = myDocumentList.GetInvoiceDocumentListItemFromText(quickFile.Name);
            uploadedDocument.FileOptions.CheckIn();
            invoiceSummaryPage.Dialog.UploadDocument();

            invoiceSummaryPage.QuickSearch.SearchBy(quickFile.Name);
            uploadedDocument = myDocumentList.GetInvoiceDocumentListItemFromText(quickFile.Name);
            uploadedDocument.NavigateToSummary();
            documentSummaryPage.SummaryPanel.Toggle();
            file = myDocumentList.GetVersionHistoryListItemByIndex(0);
            file.Open();

            Assert.IsNotNull(_word.GetReadOnlyLabel());
            documentSummaryPage.NavigateToParent();
            _word.Oc.WaitForLoadComplete();

            // Cleanup
            documentsListPage.QuickSearch.SearchBy(quickFile.Name);
            var testDocuments = myDocumentList.GetAllInvoiceDocumentListItems();

            foreach (var testDocument in testDocuments)
            {
                documentsListPage.QuickSearch.SearchBy(testDocument.Name);
                uploadedDocument = myDocumentList.GetInvoiceDocumentListItemFromText(testDocument.Name);
                if (uploadedDocument.Status == CheckInStatus.CheckedOut)
                {
                    uploadedDocument.FileOptions.CheckIn();
                    checkInDocumentDialog.Controls["Comments"].Set("Clean up Operation");
                    checkInDocumentDialog.UploadDocument();
                }
                uploadedDocument = myDocumentList.GetInvoiceDocumentListItemFromText(testDocument.Name);
                uploadedDocument.Delete().Confirm();
            }
            Windows.ClearWorkingTempFolder();
        }

        [TearDown]
        public void TearDown()
        {
            SaveScreenShotsAndLogs(_word);
            _word.Close();
            _word.Destroy();
        }
    }
}
