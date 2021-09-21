using System;
using System.Collections;
using System.IO;
using System.Linq;
using IntegratedDriver;
using NUnit.Framework;
using UITests.PageModel;
using UITests.PageModel.OfficeApps;
using UITests.PageModel.Shared;
using static IntegratedDriver.Constants;
using static UITests.Constants;
using static UITests.TestHelpers;

namespace UITests.RegressionTesting
{
    [TestFixture]
    public class MatterDocumentsRegressionTests : UITestBase
    {
        private Outlook _outlook;
        private Word _word;
        private Excel _excel;
        private Powerpoint _powerpoint;
        private Notepad _notepad;

        [SetUp]
        public void SetUp()
        {
            _outlook = new Outlook(TestEnvironment);
            _outlook.Launch();
            _outlook.Oc.BasicSettingsPage.LogInAsStandardUser();
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_13691 :Verify UI elements of Matter Documents page")]
        public void UIElementsMatterDocumentPage()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var documentType = OfficeApp.Word;
            var dndFileInfo = CreateDocument(documentType);

            mattersListPage.Open();
            mattersList.OpenRandom();
            matterDetails.Tabs.Open("Documents");

            //dnd word file to matter document
            DragAndDrop.FromFileSystem(dndFileInfo, matterDetails.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();
            var toastMessage = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(UploadSuccessMessage, toastMessage[0]);
            _outlook.Oc.CloseAllToastMessages();
            var matterDetailsList = matterDetails.ItemList;
            var uploadedDocument = matterDetailsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);

            //assert document properties.
            Assert.IsNotNull(uploadedDocument);
            Assert.AreEqual(dndFileInfo.Name, uploadedDocument.Name);
            Assert.AreEqual(dndFileInfo.Name, uploadedDocument.DocumentFileName);
            var expectedFileSize = TestHelpers.ConvertBytesToKb(dndFileInfo.Length, 1);
            Assert.AreEqual(expectedFileSize + " KB", uploadedDocument.DocumentSize);
            Assert.AreEqual(TestEnvironment.StandarUserName, uploadedDocument.LastModifiedBy);
            Assert.AreEqual(CheckInStatus.CheckedIn, uploadedDocument.Status);
            Assert.IsTrue(uploadedDocument.IsFileOfType(documentType));

            //cleanup
            uploadedDocument.Delete().Confirm();
            uploadedDocument = matterDetailsList.GetMatterDocumentListItemFromText(dndFileInfo.Name, false);
            Assert.IsNull(uploadedDocument);

            //Validate if no records found
            var itemList = documentsListPage.ItemList;
            if (itemList.GetCount() == 0)
            {
                Assert.AreEqual(NoRecordsFound, itemList.ListEmptyMessage());
                Assert.AreEqual(0, itemList.GetCount());
            }
            else
            {
                var folderName = GetRandomText(6);
                var addFolderDialog = documentsListPage.AddFolderDialog;
                mattersListPage.ItemList.OpenAddFolderDialog();
                addFolderDialog.Controls["Name"].Set(folderName);
                addFolderDialog.Save();
                var createdFolder = matterDetailsList.GetMatterDocumentListItemFromText(folderName, false);
                createdFolder.Open();
                itemList = documentsListPage.ItemList;
                Assert.AreEqual(0, itemList.GetCount());
                Assert.AreEqual(NoRecordsFound, itemList.ListEmptyMessage());

                //clean up
                _outlook.Oc.Header.NavigateBack();
                createdFolder = matterDetailsList.GetMatterDocumentListItemFromText(folderName, false);
                createdFolder.Delete().Confirm();
            }
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_17150 :Download document from Matter document list page")]
        public void DownloadDocumentMatterDocumentPage()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var dndFileInfo = CreateDocument(OfficeApp.Word);

            mattersListPage.Open();
            mattersList.OpenRandom();
            matterDetails.Tabs.Open("Documents");

            //Dnd word file to matter document
            DragAndDrop.FromFileSystem(dndFileInfo, matterDetails.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();
            var toastMessage = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(UploadSuccessMessage, toastMessage[0]);
            _outlook.Oc.CloseAllToastMessages();
            var matterDetailsList = matterDetails.ItemList;
            var uploadedDocument = matterDetailsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            Assert.IsNotNull(uploadedDocument);
            dndFileInfo.Delete();

            //Quick search uploaded file
            documentsListPage.QuickSearch.SearchBy(dndFileInfo.Name);
            uploadedDocument = matterDetailsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            Assert.IsNotNull(uploadedDocument);
            Assert.AreEqual(1, matterDetailsList.GetCount());

            //Download and validate document.
            var localFile = uploadedDocument.Download(dndFileInfo.Name);
            _word = new Word(TestEnvironment);
            _word.OpenDocumentFromExplorer(localFile.FullName);
            Assert.IsTrue(_word.IsDocumentOpened);
            _word.Close();
            Assert.AreEqual(InitialDefaultContent, _word.ReadWordContent(localFile.FullName));

            //Clean up
            localFile.Delete();
            uploadedDocument.Delete().Confirm();
            uploadedDocument = matterDetailsList.GetMatterDocumentListItemFromText(dndFileInfo.Name, false);
            Assert.IsNull(uploadedDocument);
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_17151 :View Document Summary")]
        public void ViewDocumentSummary()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var documentSummaryPage = _outlook.Oc.DocumentSummaryPage;
            var matterDetailsList = matterDetails.ItemList;

            mattersListPage.Open();
            mattersList.OpenRandom();
            matterDetails.Tabs.Open("Documents");

            // upload document to matter
            var dndFileInfo = CreateDocument(OfficeApp.Word);
            DragAndDrop.FromFileSystem(dndFileInfo, matterDetails.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();
            var uploadedDocument = matterDetailsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            Assert.IsNotNull(uploadedDocument);

            // quick search for document
            documentsListPage.QuickSearch.SearchBy(dndFileInfo.Name);
            uploadedDocument = matterDetailsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            Assert.IsNotNull(uploadedDocument);

            //_outlook.Oc.ReloadOc();
            // verify document summary
            uploadedDocument.NavigateToSummary();
            var summaryInfo = documentSummaryPage.GetDocumentSummaryInfo();
            Assert.IsNotEmpty(summaryInfo, "Document Summary fields are not retrieved or empty.");

            foreach (var field in summaryInfo)
            {
                Assert.IsNotEmpty(field.Text);
            }

            documentSummaryPage.SummaryPanel.Toggle();

            summaryInfo = documentSummaryPage.GetDocumentSummaryInfo();
            foreach (var field in summaryInfo)
            {
                Assert.IsFalse(field.Displayed, "Summary Info is displayed on toggle");
            }

            var documentNewVersion = CreateDocument(OfficeApp.Word);
            DragAndDrop.FromFileSystem(documentNewVersion, documentSummaryPage.DropPoint.GetElement());
            documentSummaryPage.AddDocumentDialog.UploadDocument();
            documentSummaryPage.CheckInDocumentDialog.UploadDocument();

            var versionsList = documentSummaryPage.ItemList.GetAllVersionHistoryListItems().Select(x => x.Version).ToList();
            var descendingVersionsList = versionsList.OrderByDescending(x => x).ToList();

            Assert.AreEqual(versionsList, descendingVersionsList);

            // cleanup
            documentSummaryPage.NavigateToParentMatter();
            uploadedDocument = matterDetailsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            uploadedDocument.Delete().Confirm();
            Assert.IsNull(matterDetailsList.GetMatterDocumentListItemFromText(dndFileInfo.Name, false));
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_17144 :Rename Document")]
        public void RenameDocument()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var matterDetailsList = matterDetails.ItemList;

            mattersListPage.Open();
            mattersList.OpenRandom();
            matterDetails.Tabs.Open("Documents");

            // upload document to matter
            var dndFileInfo = CreateDocument(OfficeApp.Word);
            DragAndDrop.FromFileSystem(dndFileInfo, matterDetails.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();
            var uploadedDocument = matterDetailsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            Assert.IsNotNull(uploadedDocument);

            // rename document and change extension
            uploadedDocument.Rename();
            var renameDocName = GetRandomText(6) + ".doc";
            documentsListPage.RenameDocumentDialog.Controls["Name"].Set(renameDocName);
            documentsListPage.RenameDocumentDialog.Controls["Document File Name"].Set(renameDocName);
            documentsListPage.AddFolderDialog.Save();

            documentsListPage.QuickSearch.SearchBy(".doc");
            var renamedDoc = documentsListPage.ItemList.GetMatterDocumentListItemFromText(renameDocName);
            Assert.IsNotNull(renamedDoc);
            documentsListPage.QuickSearch.Close();

            // clean up
            uploadedDocument = matterDetailsList.GetMatterDocumentListItemFromText(renameDocName);
            uploadedDocument.Delete().Confirm();
            Assert.IsNull(matterDetailsList.GetMatterDocumentListItemFromText(renameDocName, false));

            // rename while uploading document
            DragAndDrop.FromFileSystem(dndFileInfo, matterDetails.DropPoint.GetElement());
            var newName = SampleLongName + ".ppt";
            documentsListPage.AddDocumentDialog.Controls["Document Name"].Set(newName);
            documentsListPage.AddDocumentDialog.Controls["File Name"].Set(newName);
            documentsListPage.AddDocumentDialog.UploadDocument();

            uploadedDocument = matterDetailsList.GetMatterDocumentListItemFromText(newName);
            Assert.IsNotNull(uploadedDocument);

            // rename document name with special characters
            uploadedDocument.Rename();
            renameDocName = "DocumentNameCheck" + SpecialCharset;
            documentsListPage.RenameDocumentDialog.Controls["Name"].Set(renameDocName);
            documentsListPage.RenameDocumentDialog.Save();

            renamedDoc = documentsListPage.ItemList.GetMatterDocumentListItemFromText(renameDocName);
            Assert.IsNotNull(renamedDoc);

            // rename document file name with unsupported special characters
            uploadedDocument = matterDetailsList.GetMatterDocumentListItemFromText(renameDocName);
            uploadedDocument.Rename();
            documentsListPage.RenameDocumentDialog.Controls["Document File Name"].Set(renameDocName);
            documentsListPage.RenameDocumentDialog.Save();
            var toastMessage = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(SpecialCharacterErrorMessage, toastMessage[0]);
            _outlook.Oc.CloseAllToastMessages();

            // clean up
            uploadedDocument = matterDetailsList.GetMatterDocumentListItemFromText(renameDocName);
            uploadedDocument.Delete().Confirm();
            Assert.IsNull(matterDetailsList.GetMatterDocumentListItemFromText(renameDocName, false));
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_17146 : Delete Document")]
        public void DeleteDocument()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var matterDetailsList = matterDetails.ItemList;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var documentSummaryPage = _outlook.Oc.DocumentSummaryPage;

            // favorite a matter
            mattersListPage.Open();
            mattersListPage.SetNthMatterAsFavorite(0);
            mattersListPage.OpenFavoritesList();

            // open favorited matter
            mattersList.OpenFirst();
            matterDetails.Tabs.Open("Documents");

            // upload document to matter
            var dndFileInfo = CreateDocument(OfficeApp.Notepad);
            DragAndDrop.FromFileSystem(dndFileInfo, matterDetails.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();
            documentsListPage.QuickSearch.SearchBy(dndFileInfo.Name);
            var uploadedDocument = matterDetailsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            Assert.IsNotNull(uploadedDocument);

            // check out document
            uploadedDocument.FileOptions.CheckOut();
            var notepad = new Notepad(dndFileInfo.Name);
            notepad.Close();

            // delete checked out document
            uploadedDocument = matterDetailsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            uploadedDocument.Delete();
            var expectedDialogText = DeleteDocumentMessage(uploadedDocument.Name);
            var actualDialogText = documentSummaryPage.DeleteDocumentDialog.Text;
            Assert.AreEqual(expectedDialogText, actualDialogText);
            documentSummaryPage.DeleteDocumentDialog.Confirm();
            var toastMessage = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(CheckedOutDocumentDeleteMessage, toastMessage[0]);
            _outlook.Oc.CloseAllToastMessages();

            // discard check out and delete
            uploadedDocument = matterDetailsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            uploadedDocument.FileOptions.DiscardCheckOutAndRemoveLocalCopy();
            uploadedDocument = matterDetailsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            uploadedDocument.Delete();
            expectedDialogText = DeleteDocumentMessage(uploadedDocument.Name);
            actualDialogText = documentSummaryPage.DeleteDocumentDialog.Text;
            Assert.AreEqual(expectedDialogText, actualDialogText);

            // delete document
            documentSummaryPage.DeleteDocumentDialog.Confirm();
            Assert.IsNull(matterDetailsList.GetMatterDocumentListItemFromText(dndFileInfo.Name, false));

            // clean up
            mattersListPage.Open();
            mattersListPage.ClearFavorites(1);
            Assert.AreEqual(mattersList.GetCount(), 0, "Favorite list has matters after removing");
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_17155 : Verify OC support for UTF-8 and non-unicode characters in documents")]
        public void DocumentCharacterSupport()
        {
            const string DateTimeFormat = "M-dd-yyyy h-mm-ss tt";

            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var documentList = documentsListPage.ItemList;

            mattersListPage.Open();
            mattersList.OpenRandom();
            matterDetails.Tabs.Open("Documents");

            // upload document with utf-8 characters in name and body
            var testDocument = CreateDocument(OfficeApp.Notepad, Utf8content, Utf8CharSet2 + Utf8CharSet3);
            DragAndDrop.FromFileSystem(testDocument, matterDetails.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();
            var toastMessage = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(UploadSuccessMessage, toastMessage[0]);
            _outlook.Oc.CloseAllToastMessages();
            var uploadedDocument = documentList.GetMatterDocumentListItemFromText(testDocument.Name);
            Assert.IsNotNull(uploadedDocument);

            // quick search
            documentsListPage.QuickSearch.SearchBy(Utf8CharSet2);
            Assert.That(documentList.GetAllMatterDocumentListItems(), Has.All.Property(nameof(MatterDocumentListItem.Name)).Contains(Utf8CharSet2),
                $"Document list has items not containing Name - {Utf8CharSet2}");
            Assert.That(documentList.GetAllMatterDocumentListItems(), Has.All.Property(nameof(MatterDocumentListItem.DocumentFileName)).Contains(Utf8CharSet2),
                $"Document list has items not containing File Name - {Utf8CharSet2}");
            documentsListPage.QuickSearch.SearchBy(Utf8CharSet3);
            Assert.That(documentList.GetAllMatterDocumentListItems(), Has.All.Property(nameof(MatterDocumentListItem.Name)).Contains(Utf8CharSet3),
                $"Document list has items not containing Name - {Utf8CharSet3}");
            Assert.That(documentList.GetAllMatterDocumentListItems(), Has.All.Property(nameof(MatterDocumentListItem.DocumentFileName)).Contains(Utf8CharSet3),
                $"Document list has items not containing File Name - {Utf8CharSet3}");

            documentsListPage.QuickSearch.Close();

            // create folder with utf-8 characters
            documentList.OpenAddFolderDialog();
            var dateTime = DateTime.Now.ToString(DateTimeFormat);
            var testFolderName = dateTime + Utf8CharSet1;
            documentsListPage.AddFolderDialog.Controls["Name"].Set(testFolderName);
            documentsListPage.AddFolderDialog.Save();

            var testFolder = documentList.GetMatterDocumentListItemFromText(testFolderName);
            Assert.IsNotNull(testFolder);

            testFolder.Open();
            DragAndDrop.FromFileSystem(testDocument, matterDetails.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();
            uploadedDocument = documentList.GetMatterDocumentListItemFromText(testDocument.Name);
            Assert.IsNotNull(uploadedDocument);

            // clean up
            matterDetails.Tabs.Open("Documents");
            foreach (var name in new[] { testFolderName, testDocument.Name })
            {
                documentsListPage.QuickSearch.SearchBy(name);
                documentList.GetMatterDocumentListItemFromText(name).Delete().Confirm();
                documentsListPage.QuickSearch.Close();
                Assert.IsNull(documentList.GetMatterDocumentListItemFromText(name, false));
            }
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_17154 : Matter Documents folders")]
        public void MatterDocumentsFolders()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var documentsList = documentsListPage.ItemList;
            var addFolderDialog = documentsListPage.AddFolderDialog;
            var breadCrumbsControl = documentsListPage.BreadcrumbsControl;
            var dndFileInfo = CreateDocument(OfficeApp.Word);

            mattersList.OpenRandom();
            matterDetails.Tabs.Open("Documents");

            //Validate add folder or document, sort, quick search icon visible.
            Assert.IsTrue(documentsList.IsQuickSearchIconDisplayed);
            Assert.IsTrue(documentsList.IsSortIconVisible);
            Assert.IsTrue(documentsList.IsAddFolderButtonVisible);

            //Add folder.
            var folderName = GetRandomText(6);
            mattersList.OpenAddFolderDialog();
            addFolderDialog.Controls["Name"].Set(folderName);
            addFolderDialog.Save();
            var createdFolder = documentsList.GetMatterDocumentListItemFromText(folderName);
            Assert.IsNotNull(createdFolder);
            createdFolder.Open();

            //Add sub folder with special character.
            var specialCharacters = SpecialCharsetInFolderCreation.ToCharArray();
            foreach (var specialCharacter in specialCharacters)
            {
                mattersList.OpenAddFolderDialog();
                var subFolderName = $"{folderName}{specialCharacter}";
                addFolderDialog.Controls["Name"].Set(subFolderName);
                addFolderDialog.Save();
                var toasterMessage = _outlook.Oc.GetAllToastMessages();
                Assert.AreEqual(SpecialCharacterErrorMessage, toasterMessage[0]);
                _outlook.Oc.CloseAllToastMessages();
            }

            //Add sub folder with long folder name.
            mattersList.OpenAddFolderDialog();
            var longFolderName = GetRandomText(255);
            addFolderDialog.Controls["Name"].Set(longFolderName);
            addFolderDialog.Save();
            var createdSubFolder = documentsList.GetMatterDocumentListItemFromText(longFolderName);
            Assert.IsNotNull(createdSubFolder);
            createdSubFolder.Delete().Confirm();
            createdSubFolder = documentsList.GetMatterDocumentListItemFromText(longFolderName, false);
            Assert.IsNull(createdSubFolder);

            //Rename folder.
            var renamedFolderName = GetRandomText(6);
            breadCrumbsControl.NavigateToTheRoot();
            createdFolder = documentsList.GetMatterDocumentListItemFromText(folderName, false);
            createdFolder.Rename();
            addFolderDialog.Controls["Name"].Set(renamedFolderName);
            addFolderDialog.Update();
            var renamedFolder = documentsList.GetMatterDocumentListItemFromText(renamedFolderName);
            createdFolder = documentsList.GetMatterDocumentListItemFromText(folderName, false);
            Assert.IsNotNull(renamedFolder);
            Assert.IsNull(createdFolder);

            //clean up
            renamedFolder = documentsList.GetMatterDocumentListItemFromText(renamedFolderName, false);
            renamedFolder.Delete().Confirm();
            renamedFolder = documentsList.GetMatterDocumentListItemFromText(renamedFolderName, false);
            Assert.IsNull(renamedFolder);

            //Raname folder having checked out document.
            folderName = GetRandomText(6);
            mattersList.OpenAddFolderDialog();
            addFolderDialog = documentsListPage.AddFolderDialog;
            addFolderDialog.Controls["Name"].Set(folderName);
            addFolderDialog.Save();
            createdFolder = documentsList.GetMatterDocumentListItemFromText(folderName);
            Assert.IsNotNull(createdFolder);
            createdFolder.Open();
            DragAndDrop.FromFileSystem(dndFileInfo, matterDetails.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();
            var toastMessage = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(UploadSuccessMessage, toastMessage[0]);
            _outlook.Oc.CloseAllToastMessages();
            var uploadedDocument = documentsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            Assert.IsNotNull(uploadedDocument);

            //check out document
            uploadedDocument.FileOptions.CheckOut();
            _word = new Word(TestEnvironment);
            _word.Attach(dndFileInfo.Name);
            _word.ReplaceTextWith(InitialDefaultContent);
            _word.SaveDocument();
            _word.Close();
            uploadedDocument = documentsList.GetMatterDocumentListItemFromText(dndFileInfo.Name, false);
            var documentStatus = uploadedDocument.Status.ToLower();
            Assert.AreEqual(CheckInStatus.CheckedOut.ToLower(), documentStatus);
            breadCrumbsControl.NavigateToTheRoot();
            createdFolder = documentsList.GetMatterDocumentListItemFromText(folderName, false);
            createdFolder.Rename();
            addFolderDialog.Controls["Name"].Set(GetRandomText(6));
            addFolderDialog.Update();
            toastMessage = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(RenameFolderHasCheckedOutDocumentErrorMessage, toastMessage[0]);
            _outlook.Oc.CloseAllToastMessages();
            createdFolder = documentsList.GetMatterDocumentListItemFromText(folderName, false);
            Assert.IsNotNull(createdFolder);

            //Delete folder having checked out document.
            createdFolder.Delete().Confirm();
            toastMessage = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(DeleteFolderHasCheckedOutDocumentErrorMessage, toastMessage[0]);
            _outlook.Oc.CloseAllToastMessages();
            createdFolder = documentsList.GetMatterDocumentListItemFromText(folderName, false);
            Assert.IsNotNull(createdFolder);

            //clean up
            createdFolder.Open();
            uploadedDocument = documentsList.GetMatterDocumentListItemFromText(dndFileInfo.Name, false);
            uploadedDocument.FileOptions.DiscardCheckOutAndRemoveLocalCopy();
            uploadedDocument = documentsList.GetMatterDocumentListItemFromText(dndFileInfo.Name, false);
            documentStatus = uploadedDocument.Status.ToLower();
            Assert.AreEqual(CheckInStatus.CheckedIn.ToLower(), documentStatus);
            breadCrumbsControl.NavigateToTheRoot();
            createdFolder = documentsList.GetMatterDocumentListItemFromText(folderName, false);
            createdFolder.Delete().Confirm();
            createdFolder = documentsList.GetMatterDocumentListItemFromText(folderName, false);
            Assert.IsNull(createdFolder);

            //Navigate from bread crumbs
            var firstFolderName = GetRandomText(6);
            var secondFolderName = GetRandomText(6);
            var thirdFolderName = GetRandomText(6);

            mattersList.OpenAddFolderDialog();
            addFolderDialog = documentsListPage.AddFolderDialog;
            addFolderDialog.Controls["Name"].Set(firstFolderName);
            addFolderDialog.Save();
            var firstCreatedFolder = documentsList.GetMatterDocumentListItemFromText(firstFolderName);
            Assert.IsNotNull(firstCreatedFolder);
            firstCreatedFolder.Open();

            mattersList.OpenAddFolderDialog();
            addFolderDialog = documentsListPage.AddFolderDialog;
            addFolderDialog.Controls["Name"].Set(secondFolderName);
            addFolderDialog.Save();
            var secondCreatedFolder = documentsList.GetMatterDocumentListItemFromText(secondFolderName);
            Assert.IsNotNull(secondCreatedFolder);
            secondCreatedFolder.Open();

            mattersList.OpenAddFolderDialog();
            addFolderDialog = documentsListPage.AddFolderDialog;
            addFolderDialog.Controls["Name"].Set(thirdFolderName);
            addFolderDialog.Save();
            var thirdCreatedFolder = documentsList.GetMatterDocumentListItemFromText(thirdFolderName);
            Assert.IsNotNull(thirdCreatedFolder);

            breadCrumbsControl.NavigateToFolder(firstFolderName);
            secondCreatedFolder = documentsList.GetMatterDocumentListItemFromText(secondFolderName, false);
            Assert.IsNotNull(secondCreatedFolder);
            Assert.AreEqual(1, documentsList.GetCount());

            //clean up
            breadCrumbsControl.NavigateToTheRoot();
            firstCreatedFolder = documentsList.GetMatterDocumentListItemFromText(firstFolderName, false);
            firstCreatedFolder.Delete().Confirm();
            firstCreatedFolder = documentsList.GetMatterDocumentListItemFromText(firstFolderName, false);
            Assert.IsNull(firstCreatedFolder);
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_17148 : View document from document list")]
        public void ViewDocuments()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var documentsList = documentsListPage.ItemList;
            var addFolderDialog = documentsListPage.AddFolderDialog;
            var settingsPage = _outlook.Oc.SettingsPage;
            var breadCrumbsControl = documentsListPage.BreadcrumbsControl;

            var folderName = GetRandomText(4);
            mattersList.OpenRandom();
            matterDetails.Tabs.Open("Documents");
            var matterName = matterDetails.MatterName;

            // Create new folder
            mattersList.OpenAddFolderDialog();
            addFolderDialog.Controls["Name"].Set(folderName);
            addFolderDialog.Save();

            // Add new document
            var createdFolder = documentsList.GetMatterDocumentListItemFromText(folderName);
            Assert.IsNotNull(createdFolder);
            createdFolder.Open();
            var dndFileInfo = CreateDocument(OfficeApp.Word);
            DragAndDrop.FromFileSystem(dndFileInfo, matterDetails.DropPoint.GetElement());
            documentsListPage.AddDocumentDialog.UploadDocument();

            var toastMessage = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(UploadSuccessMessage, toastMessage[0]);
            _outlook.Oc.CloseAllToastMessages();
            var uploadedDocument = documentsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            Assert.IsNotNull(uploadedDocument);

            // Verify read only banner message for checked in document
            uploadedDocument.Open();
            var fileName = uploadedDocument.Name;
            _word = new Word(TestEnvironment);
            _word.Attach(fileName);
            Assert.IsNotNull(_word.GetReadOnlyLabel());
            _word.Close();

            // Check out a document and verify no banner message shown
            uploadedDocument.FileOptions.CheckOut();
            _word = new Word(TestEnvironment);
            _word.Attach(fileName);
            Assert.IsNull(_word.GetReadOnlyLabel());
            _word.Close();

            // Log out sbrown from office companion
            _outlook.Oc.OpenSettings();
            settingsPage.OpenConfiguration();
            settingsPage.LogOut().Confirm();

            // Log in office companion as dmaxwell
            _outlook.Oc.BasicSettingsPage.LogInAsAttorneyUser();
            mattersListPage.Open();
            var matter = mattersList.GetMatterListItemFromText(matterName);
            matter.Open();
            matterDetails.Tabs.Open("Documents");

            // View document which was checked out by other users
            createdFolder = documentsList.GetMatterDocumentListItemFromText(folderName);
            createdFolder.Open();
            uploadedDocument = documentsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            uploadedDocument.Open();
            fileName = uploadedDocument.Name;
            _word = new Word(TestEnvironment);
            _word.Attach(fileName);
            Assert.IsNotNull(_word.GetReadOnlyLabelForCheckedOutDocument());
            _word.Close();

            // Log out dmaxwell from office companion
            _outlook.Oc.OpenSettings();
            settingsPage.OpenConfiguration();
            settingsPage.LogOut().Confirm();

            // Log in office companion as sbrown
            _outlook.Oc.BasicSettingsPage.LogInAsStandardUser();
            mattersListPage.Open();
            matter = mattersList.GetMatterListItemFromText(matterName);
            matter.Open();
            matterDetails.Tabs.Open("Documents");

            // discard check out
            createdFolder = documentsList.GetMatterDocumentListItemFromText(folderName);
            createdFolder.Open();
            uploadedDocument = documentsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            uploadedDocument.FileOptions.DiscardCheckOutAndRemoveLocalCopy();

            // Clean up
            uploadedDocument = documentsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            uploadedDocument.Delete().Confirm();
            uploadedDocument = documentsList.GetMatterDocumentListItemFromText(dndFileInfo.Name);
            Assert.IsNull(uploadedDocument);
            breadCrumbsControl.NavigateToTheRoot();
            createdFolder = documentsList.GetMatterDocumentListItemFromText(folderName, false);
            createdFolder.Delete().Confirm();
            createdFolder = documentsList.GetMatterDocumentListItemFromText(folderName);
            Assert.IsNull(createdFolder);
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Editing Documents - CheckIn")]
        public void CheckInCheckOutDocument()
        {
            var checkedIn = CheckInStatus.CheckedIn.ToLower();
            var checkedOut = CheckInStatus.CheckedOut.ToLower();

            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var documentsListPage = _outlook.Oc.DocumentsListPage;
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var globalDocumentList = globalDocumentsPage.ItemList;
            var checkInDocumentDialog = globalDocumentsPage.CheckInDocumentDialog;

            var testDocuments = new ArrayList();

            foreach (OfficeApp officeAppName in Enum.GetValues(typeof(OfficeApp)))
            {
                if (officeAppName.ToString() != "None" && officeAppName.ToString() != "Outlook")
                {
                    mattersListPage.Open();
                    mattersList.OpenFirst();
                    matterDetails.Tabs.Open("Documents");
                    var testDocument = CreateDocument(officeAppName);
                    var testDocumentName = testDocument.Name;
                    DragAndDrop.FromFileSystem(testDocument, documentsListPage.DropPoint.GetElement());
                    documentsListPage.AddDocumentDialog.UploadDocument();

                    globalDocumentsPage.Open();
                    globalDocumentsPage.OpenRecentDocumentsList();
                    globalDocumentsPage.QuickSearch.SearchBy(testDocumentName);
                    var document = globalDocumentList.GetGlobalDocumentListItemByIndex(0);

                    var documentStatus = document.Status.ToLower();
                    Assert.AreEqual(checkedIn, documentStatus);

                    // Checkout from GDL page
                    document.FileOptions.CheckOut();

                    globalDocumentsPage.OpenCheckedOutDocumentsList();
                    globalDocumentsPage.QuickSearch.SearchBy(testDocumentName);
                    document = globalDocumentList.GetGlobalDocumentListItemByIndex(0);

                    // status should be checked Out
                    documentStatus = document.Status.ToLower();
                    Assert.AreEqual(checkedOut, documentStatus);

                    // Adding all test Documents here for clean up later
                    testDocuments.Add(testDocumentName);

                    switch (officeAppName.ToString())
                    {
                        case "Word":
                            _word = new Word(TestEnvironment);
                            _word.Attach(testDocumentName);
                            _word?.Close();
                            break;

                        case "Excel":
                            _excel = new Excel(TestEnvironment);
                            _excel.Attach(testDocumentName);
                            _excel?.Close();

                            break;

                        case "Powerpoint":
                            _powerpoint = new Powerpoint(TestEnvironment);
                            _powerpoint.Attach(testDocumentName);
                            _powerpoint?.Close();
                            break;

                        default:
                            _notepad = new Notepad(document.Name);
                            _notepad?.Close();
                            break;
                    }
                }
            }

            foreach (var document in testDocuments)
            {
                globalDocumentsPage.OpenRecentDocumentsList();
                globalDocumentsPage.QuickSearch.SearchBy(document.ToString());
                var currentTestDocument = globalDocumentList.GetGlobalDocumentListItemByIndex(0);
                // Check In from GDL
                currentTestDocument.FileOptions.CheckIn();
                checkInDocumentDialog.Controls["Comments"].Set("Test Document : CheckIn Operation");
                checkInDocumentDialog.UploadDocument();

                // clean up
                currentTestDocument.Delete().Confirm();
            }
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Editing Documents - DnDDocument")]
        public void DragAndDropDocument()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailPage = _outlook.Oc.MatterDetailsPage;
            var documentListPage = _outlook.Oc.DocumentsListPage;
            var documentList = documentListPage.ItemList;
            var documentSummaryPage = _outlook.Oc.DocumentSummaryPage;
            var checkInDialog = documentSummaryPage.CheckInDocumentDialog;
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;

            var globalDocumentsList = globalDocumentsPage.ItemList;
            var versionHistoryList = documentSummaryPage.ItemList;

            var testEmail = _outlook.AddTestEmailsToFolder(1, FileSize.VerySmall, true, OfficeApp.Word).First();
            var filename = new FileInfo(testEmail.Value).Name;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetailPage.Tabs.Open("Documents");

            _outlook.OpenTestEmailFolder();
            _outlook.SelectNthItem(0);

            var attachment = _outlook.GetAttachmentFromReadingPane(filename);
            DragAndDrop.FromElementToElement(attachment, matterDetailPage.DropPoint.GetElement());
            checkInDialog.Controls["Comments"].Set("1. Document DND from Outlook to OC-Matter Documents");
            documentListPage.AddDocumentDialog.UploadDocument();

            var uploadedDocument = documentList.GetMatterDocumentListItemFromText(filename);
            Assert.IsNotNull(uploadedDocument);
            var documentStatus = uploadedDocument.Status.ToLower();
            Assert.AreEqual(CheckInStatus.CheckedIn.ToLower(), documentStatus);

            uploadedDocument.NavigateToSummary();
            documentSummaryPage.SummaryPanel.Toggle();

            testEmail = _outlook.AddTestEmailsToFolder(1, FileSize.VerySmall, true, OfficeApp.Word).First();
            var testFileName = new FileInfo(testEmail.Value).Name;
            _outlook.OpenTestEmailFolder();
            _outlook.SelectNthItem(0);

            attachment = _outlook.GetAttachmentFromReadingPane(testFileName);
            DragAndDrop.FromElementToElement(attachment, documentSummaryPage.DropPoint.GetElement());
            documentListPage.AddDocumentDialog.Proceed();
            checkInDialog.Controls["Comments"].Set("2. Document DND from Outlook to OC-Document Summary");
            documentListPage.AddDocumentDialog.UploadDocument();
            var documentVersions = versionHistoryList.GetAllVersionHistoryListItems();
            Assert.AreEqual(documentVersions.Count, 2);

            var testdndDocument = CreateDocument(OfficeApp.Word);
            DragAndDrop.FromFileSystem(testdndDocument, documentSummaryPage.DropPoint.GetElement());
            documentListPage.AddDocumentDialog.Proceed();
            checkInDialog.Controls["Comments"].Set("3. Document DND from FileSystem to OC-Document Summary");
            documentListPage.AddDocumentDialog.UploadDocument();
            documentVersions = versionHistoryList.GetAllVersionHistoryListItems();
            Assert.AreEqual(documentVersions.Count, 3);

            // clean up test document
            globalDocumentsPage.Open();
            globalDocumentsPage.OpenRecentDocumentsList();
            globalDocumentsPage.QuickSearch.SearchBy(filename);
            var testDocument = globalDocumentsList.GetGlobalDocumentListItemByIndex(0);
            testDocument.Delete().Confirm();
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("TC : Drag and drop files from OC to Windows file system")]
        public void DndFromOcToFileSystem()
        {
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var globalDocumentsList = globalDocumentsPage.ItemList;
            var documentsFilterDialog = globalDocumentsPage.GlobalDocumentsListFilterDialog;
            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;

            mattersListPage.Open();
            mattersList.OpenRandom();
            matterDetailsPage.Tabs.Open("Documents");

            var wordDocument = CreateDocument(OfficeApp.Word);
            var excelDocument = CreateDocument(OfficeApp.Excel);
            var pptDocument = CreateDocument(OfficeApp.Powerpoint);
            var notepadDocument = CreateDocument(OfficeApp.Notepad);

            string[] fileExtensions = { ".doc", ".xlsx", ".ppt", ".txt" };
            var testDocuments = new ArrayList();
            var path = Windows.GetWorkingTempFolder();
            Assert.NotNull(path);

            // Dnd Multiple Documents
            DragAndDrop.AllFilesInFolderDndOC(path, matterDetailsPage.DropPoint.GetElement());

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenRecentDocumentsList();
            _outlook.Oc.CloseAllToastMessages();
            Windows.ClearWorkingTempFolder();
            globalDocumentsPage.OpenRecentDocumentsList();

            var addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(wordDocument.Name);
            Assert.IsNotNull(addedDocument);
            addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(excelDocument.Name);
            Assert.IsNotNull(addedDocument);
            addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(pptDocument.Name);
            Assert.IsNotNull(addedDocument);
            addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(notepadDocument.Name);
            Assert.IsNotNull(addedDocument);

            for (var i = 0; i < fileExtensions.Length; i++)
            {
                globalDocumentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
                documentsFilterDialog.Controls["Status"].Set(CheckInStatus.CheckedIn);
                documentsFilterDialog.Controls["Name"].Set(fileExtensions[i]);
                documentsFilterDialog.Apply();

                var testDocument = globalDocumentsList.GetGlobalDocumentListItemByIndex(0);
                Assert.NotNull(testDocument);
                testDocuments.Add(testDocument.Name);

                // Dnd from OC to FileSystem
                DragAndDrop.ToFileSystem(testDocument.PrimaryElement, path);
            }
            var filesAddedByDnd = Directory.EnumerateFiles(path.ToString());
            var filesAddedByDndCount = filesAddedByDnd.ToList().Count;
            Assert.AreEqual(testDocuments.Count, filesAddedByDndCount, (testDocuments.Count - filesAddedByDndCount) + " files are not successfully DND into FileSystem");

            // clean up
            globalDocumentsList.OpenListOptionsMenu().RestoreDefaults();
            foreach (string testDocumentName in testDocuments)
            {
                globalDocumentsList.GetGlobalDocumentListItemFromText(testDocumentName).Delete().Confirm();
            }
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("TC : Upload Multiple Documents through DND")]
        public void MultipleDocumentsUpload()
        {
            var globalDocumentsPage = _outlook.Oc.GlobalDocumentsPage;
            var globalDocumentsList = globalDocumentsPage.ItemList;
            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;

            mattersListPage.Open();
            mattersList.OpenRandom();
            matterDetailsPage.Tabs.Open("Documents");
            var testDocuments = new ArrayList();
            var wordDocument = CreateDocument(OfficeApp.Word);
            testDocuments.Add(wordDocument);
            var excelDocument = CreateDocument(OfficeApp.Excel);
            testDocuments.Add(excelDocument);
            var pptDocument = CreateDocument(OfficeApp.Powerpoint);
            testDocuments.Add(pptDocument);
            var notepadDocument = CreateDocument(OfficeApp.Notepad);
            testDocuments.Add(notepadDocument);

            var path = Windows.GetWorkingTempFolder();
            Assert.NotNull(path);

            // Dnd Multiple Documents
            DragAndDrop.AllFilesInFolderDndOC(path, matterDetailsPage.DropPoint.GetElement());

            globalDocumentsPage.Open();
            globalDocumentsPage.OpenRecentDocumentsList();

            _outlook.Oc.CloseAllToastMessages();

            globalDocumentsPage.OpenRecentDocumentsList();

            foreach (FileInfo testDocument in testDocuments)
            {
                globalDocumentsPage.QuickSearch.SearchBy(testDocument.Name);
                var addedDocument = globalDocumentsList.GetGlobalDocumentListItemFromText(testDocument.Name);
                Assert.IsNotNull(addedDocument, "Document is not properly uploaded to OC");
                Assert.AreEqual(TestEnvironment.StandarUserName, addedDocument.CreatedByFullName);
                Assert.AreEqual(CheckInStatus.CheckedIn, addedDocument.Status);
                // Clean up
                addedDocument.Delete().Confirm();
            }
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("TC-17152 Part 1 : View and download Multiple document versions from Document summary")]
        public void ViewDownloadMultipleVersionFromSummary()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var documentSummaryPage = _outlook.Oc.DocumentSummaryPage;
            var checkInDialog = documentSummaryPage.CheckInDocumentDialog;
            var documentListPage = _outlook.Oc.DocumentsListPage;
            var documentList = documentListPage.ItemList;
            var versionHistoryList = documentSummaryPage.ItemList;
            var documentSummary = _outlook.Oc.DocumentSummaryPage;

            const int DocumentVersionsCreated = 2;
            const string DocumentContent = "This is a test document with version : ";

            var testDocument = CreateDocument(OfficeApp.Word, DocumentContent + "1");

            mattersListPage.Open();
            mattersList.OpenRandom();
            matterDetailsPage.Tabs.Open("Documents");

            var folderName = GetRandomText(5);
            documentList.OpenAddFolderDialog();
            documentListPage.AddFolderDialog.Controls["Name"].Set(folderName);
            documentListPage.AddFolderDialog.Save();

            var testFolder = documentList.GetMatterDocumentListItemFromText(folderName);
            Assert.IsNotNull(testFolder);
            testFolder.Open();

            DragAndDrop.FromFileSystem(testDocument, matterDetailsPage.DropPoint.GetElement());
            checkInDialog.Controls["Comments"].Set($"{AutomatedComment} : Version 1");
            documentListPage.AddDocumentDialog.UploadDocument();

            documentListPage.QuickSearch.SearchBy(testDocument.Name);
            var uploadedDocument = documentList.GetMatterDocumentListItemFromText(testDocument.Name);
            Assert.IsNotNull(uploadedDocument);

            // Creating multiple versions of uploaded document
            for (var i = 0; i < DocumentVersionsCreated; i++)
            {
                documentListPage.QuickSearch.SearchBy(testDocument.Name);
                uploadedDocument = documentList.GetMatterDocumentListItemFromText(testDocument.Name);
                uploadedDocument.FileOptions.CheckOut();

                _word = new Word(TestEnvironment);
                _word.Attach(testDocument.Name);

                _word.ReplaceTextWith($"{DocumentContent}{i + 2}");
                _word.SaveDocument();
                _word.Close();

                uploadedDocument = documentList.GetMatterDocumentListItemFromText(testDocument.Name);
                uploadedDocument.FileOptions.CheckIn();
                checkInDialog.Controls["Comments"].Set($"{AutomatedComment} : Version {i + 2}");
                documentListPage.AddDocumentDialog.UploadDocument();

                matterDetailsPage.Tabs.Open("Documents");

                testFolder = documentList.GetMatterDocumentListItemFromText(folderName);
                Assert.IsNotNull(testFolder);
                testFolder.Open();
            }

            uploadedDocument = documentList.GetMatterDocumentListItemFromText(testDocument.Name);

            uploadedDocument.NavigateToSummary();
            documentSummaryPage.SummaryPanel.Toggle();

            var documentVersions = versionHistoryList.GetAllListItems();
            Assert.AreEqual(DocumentVersionsCreated + 1, documentVersions.Count);

            // View and validate content of all Versions of Checked In Document
            var currentVersionOfDocument = versionHistoryList.GetListItemByIndex(0);
            Assert.IsNotNull(currentVersionOfDocument);

            for (var i = documentVersions.Count; i > 0; i--)
            {
                currentVersionOfDocument = versionHistoryList.GetListItemByIndex(documentVersions.Count - i);
                currentVersionOfDocument.Open();

                _word = new Word(TestEnvironment);
                _word.Attach(testDocument.Name);

                var documentContent = _word.ReadActiveFileContent();
                Assert.AreEqual(DocumentContent + i, documentContent);

                _word.Close();
            }

            Windows.ClearWorkingTempFolder();

            // To download and validate the latest version
            var documentAllVersions = versionHistoryList.GetAllVersionHistoryListItems();
            Assert.IsNotNull(documentAllVersions[0]);

            var localFile = documentAllVersions[0].Download(testDocument.Name);

            _word = new Word(TestEnvironment);

            _word.OpenDocumentFromExplorer(localFile.FullName);
            Assert.IsTrue(_word.IsDocumentOpened);

            _word.Close();

            Assert.AreEqual(DocumentContent + documentAllVersions[0].Version, _word.ReadWordContent(localFile.FullName));

            documentSummary.NavigateToParentMatter();

            // Clean up
            testFolder = documentList.GetMatterDocumentListItemFromText(folderName);
            testFolder.Delete().Confirm();
        }

        [TearDown]
        public void TearDown()
        {
            SaveScreenShotsAndLogs(_outlook);
            SaveScreenShotsAndLogs(_word);
            _outlook?.Destroy();
            _word?.Destroy();
            _excel?.Destroy();
            _powerpoint?.Destroy();
        }
    }
}
