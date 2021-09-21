using System;
using System.IO;
using IntegratedDriver;
using NUnit.Framework;
using UITests.PageModel;
using UITests.PageModel.OfficeApps;
using UITests.PageModel.Shared;
using static UITests.Constants;
using static UITests.TestHelpers;

namespace UITests.RegressionTesting
{
    [TestFixture]
    public class MultipleOfficeAppRegressionTests : UITestBase
    {
        private OfficeApplication _officeApp;
        private Outlook _outlook;
        private FileInfo _quickFileInfo;

        [SetUp]
        public void SetUp()
        {
        }

        public void LaunchAndLogin(OfficeApp officeApp)
        {
            if (officeApp != OfficeApp.Outlook)
            {
                switch (officeApp)
                {
                    case OfficeApp.Word:
                        _officeApp = new Word(TestEnvironment);
                        break;

                    case OfficeApp.Excel:
                        _officeApp = new Excel(TestEnvironment);
                        break;

                    case OfficeApp.Powerpoint:
                        _officeApp = new Powerpoint(TestEnvironment);
                        break;
                }
                _quickFileInfo = CreateDocument(officeApp);
                _officeApp.OpenDocumentFromExplorer(_quickFileInfo.FullName);
                _officeApp.AttachToOc();
                _officeApp.Oc.BasicSettingsPage.LogInAsStandardUser();
            }
            else if (officeApp == OfficeApp.Outlook)
            {
                _outlook = new Outlook(TestEnvironment);
                _outlook.Launch();
                _outlook.Oc.BasicSettingsPage.LogInAsStandardUser();
            }
            else
            {
                throw new NotSupportedException(@"Provided Office application not yet supported.");
            }
        }

        [Test]
        [Category(RegressionTestCategory)]
        [TestCase(OfficeApp.Word)]
        [TestCase(OfficeApp.Excel)]
        [TestCase(OfficeApp.Powerpoint)]
        [Description("TC : Upload a new version of an existing document from document summary through Dnd and Quick file.")]
        public void UploadDocFromDocSummary(OfficeApp officeApp)
        {
            LaunchAndLogin(officeApp);

            var globalDocumentsPage = _officeApp.Oc.GlobalDocumentsPage;
            var globalDocumentsList = globalDocumentsPage.ItemList;
            var mattersListPage = _officeApp.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetailsPage = _officeApp.Oc.MatterDetailsPage;
            var matterDetailsList = matterDetailsPage.ItemList;
            var documentsListPage = _officeApp.Oc.DocumentsListPage;
            var checkInDialog = globalDocumentsPage.CheckInDocumentDialog;
            var documentSummaryPage = _officeApp.Oc.DocumentSummaryPage;
            var versionHistoryList = documentSummaryPage.ItemList;

            mattersListPage.Open();
            mattersList.OpenRandom();
            matterDetailsPage.Tabs.Open("Documents");

            // Quick File Document
            matterDetailsPage.QuickFile();
            checkInDialog.Controls["Comments"].Set("Version 1 : QuickFile Initial Document Upload");
            checkInDialog.UploadDocument();
            documentsListPage.QuickSearch.SearchBy(_quickFileInfo.Name);
            var uploadedDocument = matterDetailsList.GetMatterDocumentListItemFromText(_quickFileInfo.Name);
            Assert.IsNotNull(uploadedDocument);
            Assert.AreEqual(1, matterDetailsList.GetCount());

            _officeApp.CloseDocument();

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenRecentDocumentsList();

            globalDocumentsPage.QuickSearch.SearchBy(_quickFileInfo.Name);
            var addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(_quickFileInfo.Name);
            Assert.NotNull(addedDocument);
            Assert.AreEqual(CheckInStatus.CheckedIn.ToLower(), addedDocument.Status.ToLower());

            addedDocument.FileOptions.CheckOut();

            _officeApp.SaveDocument();

            addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(_quickFileInfo.Name);
            Assert.NotNull(addedDocument);
            Assert.AreEqual(CheckInStatus.CheckedOut.ToLower(), addedDocument.Status.ToLower());

            addedDocument.NavigateToSummary();

            checkInDialog = documentSummaryPage.CheckInDocumentDialog;
            documentSummaryPage.SummaryPanel.Toggle();

            documentSummaryPage.QuickFile();
            checkInDialog.Controls["Comments"].Set("version 2 : QuickFile Upload on DocumentSummary Page");
            checkInDialog.UploadDocument();
            _officeApp.CloseDocument();

            DragAndDrop.FromFileSystem(_quickFileInfo, documentSummaryPage.DropPoint.GetElement());
            checkInDialog.Controls["Comments"].Set("version 3 : DragAndDrop from file System on DocumentSummary Page");
            checkInDialog.UploadDocument();

            var documentVersions = versionHistoryList.GetAllListItems();
            Assert.AreEqual(3, documentVersions.Count);

            // Clean up
            globalDocumentsPage.Open();
            globalDocumentsPage.OpenRecentDocumentsList();
            globalDocumentsPage.QuickSearch.SearchBy(_quickFileInfo.Name);
            addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(_quickFileInfo.Name);
            addedDocument.Delete().Confirm();
        }

        [TearDown]
        public void TearDown()
        {
            SaveScreenShotsAndLogs(_officeApp);
            _officeApp?.Destroy();
            if (_outlook != null)
            {
                SaveScreenShotsAndLogs(_outlook);
                _outlook?.Destroy();
            }
        }
    }
}
