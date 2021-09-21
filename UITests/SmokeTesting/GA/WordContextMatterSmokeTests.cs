using System.Linq;
using IntegratedDriver;
using NUnit.Framework;
using UITests.PageModel.OfficeApps;
using static UITests.Constants;
using static UITests.TestHelpers;

namespace UITests.SmokeTesting.GA
{
    [TestFixture]
    public class WordContextMatterSmokeTests : UITestBase
    {
        private Word _word;

        [SetUp]
        public void SetUp()
        {
            _word = new Word(TestEnvironment);
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("Test case reference: Quick file in document summary")]
        public void QuickFileInDocumentSummary()
        {
            const string UploadComment = "Upload comment";
            const double SecondVersionNumber = 2.0d;

            var dndFileInfo = CreateDocument(OfficeApp.Word);
            var quickFileInfo = CreateDocument(OfficeApp.Word);

            _word.OpenDocumentFromExplorer(quickFileInfo.FullName);
            _word.AttachToOc();
            _word.Oc.BasicSettingsPage.LogIn();

            var documentsListPage = _word.Oc.DocumentsListPage;
            var documentSummary = _word.Oc.DocumentSummaryPage;
            var matterDetailsPage = _word.Oc.MatterDetailsPage;
            var mattersListPage = _word.Oc.MattersListPage;
            var documentsList = documentsListPage.ItemList;
            var mattersList = mattersListPage.ItemList;
            var versionsList = documentSummary.ItemList;

            mattersListPage.Open();
            mattersList.OpenFirst();
            matterDetailsPage.Tabs.Open("Documents");

            DragAndDrop.FromFileSystem(dndFileInfo, matterDetailsPage.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();
            var dndDocument = documentsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            dndDocument.NavigateToSummary();

            // Verify document summary and versions history
            var versions = versionsList.GetCount();
            Assert.AreEqual(versions, 1, "Uploaded document has more or less than 1 version history record.");

            var summaryInfo = documentSummary.GetDocumentSummaryInfo();
            Assert.IsNotEmpty(summaryInfo, "Document Summary fields are not retrieved or empty.");

            foreach (var field in summaryInfo)
            {
                Assert.IsNotNull(field.Text, "Document Summary field is empty.");
            }

            // Verify quick file
            documentSummary.QuickFile();

            var expectedDialogText = UploadDocumentMessage(quickFileInfo.Name, dndFileInfo.Name);
            var actualDialogText = documentSummary.AddDocumentDialog.Text;

            Assert.AreEqual(expectedDialogText, actualDialogText, "Upload warning message is not correct.");

            documentSummary.AddDocumentDialog.Cancel();
            documentSummary.QuickFile();
            documentSummary.AddDocumentDialog.Proceed();
            documentSummary.AddDocumentDialog.Controls["Comments"].Set(UploadComment);
            documentSummary.AddDocumentDialog.UploadDocument();

            // Verify versions history list and sort order
            var versionsUpdated = versionsList.GetAllVersionHistoryListItems().Select(x => x.Version).ToList();
            var descendingListTemplate = versionsUpdated.OrderByDescending(x => x).ToList();
            Assert.AreEqual(versionsUpdated.Count, 2, "Updated document has more or less than 2 version history records.");
            Assert.AreEqual(versionsUpdated, descendingListTemplate);

            // Verify last version list item
            var lastVersion = documentsList.GetVersionHistoryListItemByIndex(0);
            Assert.AreEqual(lastVersion.Version, SecondVersionNumber);
            Assert.AreEqual(lastVersion.CreatedBy, _word.CurrentUserDisplayName);
            Assert.AreEqual(lastVersion.Comments, UploadComment);
            Assert.IsTrue(lastVersion.IsDownloadIconVisible());
            Assert.IsNotEmpty(lastVersion.Size);
            Assert.IsNotNull(lastVersion.UploadedAt);

            // Cleanup
            _word.Oc.Header.NavigateBack();
            dndDocument = documentsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            dndDocument.Delete().Confirm();
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
