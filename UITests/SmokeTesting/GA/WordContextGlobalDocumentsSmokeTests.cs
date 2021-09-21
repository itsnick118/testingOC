using IntegratedDriver;
using NUnit.Framework;
using UITests.PageModel;
using UITests.PageModel.OfficeApps;
using static UITests.Constants;
using static UITests.TestHelpers;

namespace UITests.SmokeTesting.GA
{
    [TestFixture]
    public class WordContextGlobalDocumentsSmokeTests : UITestBase
    {
        private Word _word;

        [SetUp]
        public void SetUp()
        {
            _word = new Word(TestEnvironment);
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(GlobalDocumentsTestCategory)]
        [Description("Test case reference: Quick file a document at document summary - 'SAVE' ( quick file)")]
        public void QuickFileInDocumentSummaryAndSave()
        {
            var wordFile = CreateDocument(OfficeApp.Word);

            _word.OpenDocumentFromExplorer(wordFile.FullName);
            _word.AttachToOc();
            _word.Oc.BasicSettingsPage.LogIn();
            _word.CloseDocument();

            var mattersListPage = _word.Oc.MattersListPage;
            var matterDetails = _word.Oc.MatterDetailsPage;
            var globalDocuments = _word.Oc.GlobalDocumentsPage;
            var documentSummary = _word.Oc.DocumentSummaryPage;
            var checkInDialog = documentSummary.CheckInDocumentDialog;

            // Upload a document
            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetails.Tabs.Open("Documents");

            DragAndDrop.FromFileSystem(wordFile, matterDetails.DropPoint.GetElement());
            documentSummary.CheckInDocumentDialog.UploadDocument();

            // Verify scenario
            globalDocuments.Open();
            globalDocuments.QuickSearch.SearchBy(wordFile.Name);

            var doc = globalDocuments.ItemList.GetGlobalDocumentListItemFromText(wordFile.Name);
            doc.FileOptions.CheckOut();
            doc = globalDocuments.ItemList.GetGlobalDocumentListItemFromText(wordFile.Name);
            doc.NavigateToSummary();
            Assert.That(documentSummary.ItemList.GetCount(), Is.EqualTo(1), "There are more or less than 1 version of the document.");

            _word.ReplaceTextWith(GetRandomText(10));
            documentSummary.QuickFile();
            checkInDialog.Save();
            checkInDialog.UploadDocument();

            Assert.Warn("No visual notification about operations in progress http://mingle/projects/growth/cards/20524");
            documentSummary.WaitForStatusChangeTo(CheckInStatus.CheckedIn);
            Assert.That(documentSummary.ItemList.GetCount(), Is.EqualTo(2), "There are more or less than 2 versions of the document.");

            // Cleanup
            _word.Oc.Header.NavigateBack();
            globalDocuments.QuickSearch.SearchBy(wordFile.Name);
            doc = globalDocuments.ItemList.GetGlobalDocumentListItemFromText(wordFile.Name);
            doc.Delete().Confirm();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Category(GlobalDocumentsTestCategory)]
        [Description("Test case reference: Verify to view a document from list page (Checked in)")]
        public void ViewDocumentFromListPage()
        {
            var wordFile = CreateDocument(OfficeApp.Word);

            _word.OpenDocumentFromExplorer(wordFile.FullName);
            _word.AttachToOc();
            _word.Oc.BasicSettingsPage.LogIn();
            _word.CloseDocument();

            var mattersListPage = _word.Oc.MattersListPage;
            var matterDetails = _word.Oc.MatterDetailsPage;
            var globalDocuments = _word.Oc.GlobalDocumentsPage;
            var documentSummary = _word.Oc.DocumentSummaryPage;
            var checkInDialog = documentSummary.CheckInDocumentDialog;

            // Upload a document
            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetails.Tabs.Open("Documents");

            DragAndDrop.FromFileSystem(wordFile, matterDetails.DropPoint.GetElement());
            documentSummary.CheckInDocumentDialog.UploadDocument();

            // Verify scenario
            globalDocuments.Open();
            globalDocuments.OpenAllDocumentsList();
            globalDocuments.QuickSearch.SearchBy(wordFile.Name);
            var doc = globalDocuments.ItemList.GetGlobalDocumentListItemFromText(wordFile.Name);
            doc.Open();
            Assert.That(_word.GetReadOnlyLabel(), Is.Not.Null, "Read Only label is not displayed.");
            Assert.That(_word.IsReadOnly, Is.True);

            _word.CheckOut();
            Assert.That(_word.GetReadOnlyLabel(), Is.Null, "Read Only label is displayed on Check Out.");

            _word.ReplaceTextWith(GetRandomText(10));
            _word.SaveDocument();

            globalDocuments.QuickSearch.SearchBy(wordFile.Name);

            doc = globalDocuments.ItemList.GetGlobalDocumentListItemFromText(wordFile.Name);
            doc.FileOptions.CheckIn();
            checkInDialog.UploadDocument();

            doc = globalDocuments.ItemList.GetGlobalDocumentListItemFromText(wordFile.Name);
            doc.NavigateToSummary();
            Assert.That(documentSummary.ItemList.GetCount(), Is.EqualTo(2), "There are more or less than 2 versions of the document.");
            Assert.That(_word.IsDocumentOpened, Is.False);

            // Cleanup
            _word.Oc.Header.NavigateBack();
            globalDocuments.QuickSearch.SearchBy(wordFile.Name);
            doc = globalDocuments.ItemList.GetGlobalDocumentListItemFromText(wordFile.Name);
            doc.Delete().Confirm();
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
