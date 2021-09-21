using IntegratedDriver;
using NUnit.Framework;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using UITests.PageModel;
using UITests.PageModel.Passport;
using UITests.PageModel.Shared;
using static UITests.Constants;
using static UITests.TestHelpers;

namespace UITests.SmokeTesting.GA
{
    [TestFixture]
    public class SmokeTestsWithPersistenceEnabled : UITestBase
    {
        private Outlook _outlook;
        private MattersListPage _mattersListPage;
        private PassportPreferencesPage _passportPreferencesPage;

        [SetUp]
        public void SetUp()
        {
            _outlook = new Outlook(TestEnvironment);
            _outlook.Launch();
            _outlook.Oc.BasicSettingsPage.LogIn();

            _mattersListPage = _outlook.Oc.MattersListPage;
            _mattersListPage.ItemList.AddMatter();
            _passportPreferencesPage = _outlook.Oc.PassportPreferencesPage;
            _passportPreferencesPage.SetPersistentListPagesTo(true);
            _passportPreferencesPage.CloseWindowHandleSwitchToOc();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("Test case reference: Filters & Smoke Test Views")]
        public void MatterFilters()
        {
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var mattersFilter = _mattersListPage.MatterListFilterDialog;
            var mattersList = _mattersListPage.ItemList;

            //Select random Matter & Open it.
            mattersList.OpenRandom();

            //Fetch Matter Properties.
            var matterName = matterDetails.MatterName;
            var matterNumber = matterDetails.MatterNumber;
            var matterType = matterDetails.MatterType;
            var matterStatus = matterDetails.Status;
            var matterPIC = matterDetails.PrimaryInternalContact;
            var matterPABU = matterDetails.PracticeAreaBusinessUnit;

            _outlook.Oc.Header.NavigateBack();
            Assert.That(mattersList.ListOptionsToolTip, Is.EqualTo(ListOptionsToolTip), "List options tooltip differs from expected");

            // Apply filter - Matter Name
            mattersList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            mattersFilter.Controls["Matter Name"].Set(matterName);
            mattersFilter.Apply();
            Assert.That(mattersList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(mattersList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(mattersList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            mattersList.OpenRandom();
            var filteredMatterName = matterDetails.MatterName;
            StringAssert.Contains(filteredMatterName, matterName);

            // Save View
            _outlook.Oc.Header.NavigateBack();
            mattersList.OpenListOptionsMenu().SaveCurrentView();
            _mattersListPage.SaveCurrentViewDialog.Controls["Create New"].Set(ViewName);
            _mattersListPage.SaveCurrentViewDialog.Save();
            var savedViewMattersCount = mattersList.GetCount();

            // Clear Filters
            mattersList.OpenListOptionsMenu().RestoreDefaults();
            Assert.That(mattersList.IsFilterIconVisible, Is.False, "Filter Icon is visible");

            // Apply Saved View
            mattersList.OpenListOptionsMenu().ApplySavedView(ViewName);
            Assert.AreEqual(savedViewMattersCount, mattersList.GetCount());

            // Apply filter - Matter Number
            mattersList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            mattersFilter.Controls["Matter Number"].Set(matterNumber);
            mattersFilter.Apply();
            Assert.That(mattersList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(mattersList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(mattersList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            mattersList.OpenRandom();
            var filteredMatterNumber = matterDetails.MatterNumber;
            Assert.AreEqual(filteredMatterNumber, matterNumber);

            // Set view as default
            _outlook.Oc.Header.NavigateBack();
            mattersList.OpenListOptionsMenu().SetCurrentViewAsDefault();

            // Verify current view is set as default
            mattersList.OpenListOptionsMenu().RestoreDefaults();
            Assert.That(mattersList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(mattersList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");

            // Clear user default
            mattersList.OpenListOptionsMenu().ClearUserDefault();
            Assert.That(mattersList.IsFilterIconVisible, Is.False, "Filter Icon is visible");

            // Apply filter - Matter Type
            mattersList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            mattersFilter.Controls["Matter Type"].Set(matterType);
            mattersFilter.Apply();
            Assert.That(mattersList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(mattersList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(mattersList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            mattersList.OpenRandom();
            var filteredMatterType = matterDetails.MatterType;
            Assert.AreEqual(filteredMatterType, matterType);

            // Verify filter persist after navigation
            _outlook.Oc.Header.NavigateBack();
            Assert.That(mattersList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(mattersList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");

            // Apply filter - Matter Status
            mattersList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            mattersFilter.RestoreDefaults();
            Assert.That(mattersList.IsFilterIconVisible, Is.False, "Filter Icon is visible");

            mattersList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            mattersFilter.Controls["Matter Status"].Set(matterStatus);
            mattersFilter.Apply();
            Assert.That(mattersList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(mattersList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(mattersList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            mattersList.OpenRandom();
            var filteredMatterStatus = matterDetails.Status;
            Assert.AreEqual(filteredMatterStatus, matterStatus);

            // Apply filter - Matter Primary Internal Contact
            _outlook.Oc.Header.NavigateBack();
            mattersList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            mattersFilter.RestoreDefaults();
            Assert.That(mattersList.IsFilterIconVisible, Is.False, "Filter Icon is visible");

            mattersList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            mattersFilter.Controls["Primary Internal Contact"].Set(matterPIC);
            mattersFilter.Apply();
            Assert.That(mattersList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(mattersList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(mattersList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            mattersList.Open(mattersList.GetCount() - 1);
            var filteredMatterPIC = matterDetails.PrimaryInternalContact;
            Assert.AreEqual(filteredMatterPIC, matterPIC);

            _outlook.Oc.Header.NavigateBack();
            mattersList.OpenListOptionsMenu().RestoreDefaults();
            Assert.That(mattersList.IsFilterIconVisible, Is.False, "Filter Icon is visible");

            // Apply filter - Matter PABU
            mattersList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            mattersFilter.Controls["Practice Area - Business Unit"].Set(matterPABU);
            mattersFilter.Apply();
            Assert.That(mattersList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(mattersList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(mattersList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            mattersList.OpenRandom();
            var filteredMatterPABU = matterDetails.PracticeAreaBusinessUnit;
            Assert.AreEqual(filteredMatterPABU, matterPABU);

            _outlook.Oc.Header.NavigateBack();
            mattersList.OpenListOptionsMenu().RestoreDefaults();
            Assert.That(mattersList.IsFilterIconVisible, Is.False, "Filter Icon is visible");
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("Test case reference: Filters & Smoke Test Views")]
        public void MyMattersFilters()
        {
            var myMattersListPage = _outlook.Oc.MattersListPage;
            var myMattersList = _outlook.Oc.MattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var myMattersFilter = _mattersListPage.MatterListFilterDialog;

            myMattersListPage.OpenMyMattersList();

            //My Matters tab default behaviour
            Assert.That(myMattersList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(myMattersList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(myMattersList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");
            myMattersList.OpenListOptionsMenu().RestoreDefaults();

            //Select random Matter & Open it.
            myMattersList.OpenRandom();

            //Fetch Matter Properties.
            var myMatterName = matterDetails.MatterName;
            var myMatterNumber = matterDetails.MatterNumber;
            var myMatterType = matterDetails.MatterType;
            var myMatterStatus = matterDetails.Status;
            var myMatterPIC = matterDetails.PrimaryInternalContact;
            var myMatterPABU = matterDetails.PracticeAreaBusinessUnit;

            _outlook.Oc.Header.NavigateBack();
            Assert.That(myMattersList.ListOptionsToolTip, Is.EqualTo(ListOptionsToolTip), "List options tooltip differs from expected");

            // Apply filter - Matter Name
            myMattersList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            myMattersFilter.Controls["Matter Name"].Set(myMatterName);
            myMattersFilter.Apply();
            Assert.That(myMattersList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(myMattersList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(myMattersList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            myMattersList.OpenRandom();
            var filteredMatterName = matterDetails.MatterName;
            StringAssert.Contains(myMatterName, filteredMatterName);

            // Save View
            _outlook.Oc.Header.NavigateBack();
            myMattersList.OpenListOptionsMenu().SaveCurrentView();
            myMattersListPage.SaveCurrentViewDialog.Controls["Create New"].Set(ViewName);
            myMattersListPage.SaveCurrentViewDialog.Save();
            var savedViewMattersCount = myMattersList.GetCount();

            // Clear Filters
            myMattersList.OpenListOptionsMenu().RestoreDefaults();
            Assert.That(myMattersList.IsFilterIconVisible, Is.False, "Filter Icon is visible");

            // Apply Saved View
            myMattersList.OpenListOptionsMenu().ApplySavedView(ViewName);
            Assert.AreEqual(savedViewMattersCount, myMattersList.GetCount());

            // Apply filter - Matter Number
            myMattersList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            myMattersFilter.Controls["Matter Number"].Set(myMatterNumber);
            myMattersFilter.Apply();
            Assert.That(myMattersList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(myMattersList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(myMattersList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            myMattersList.OpenRandom();
            var filteredMatterNumber = matterDetails.MatterNumber;
            Assert.AreEqual(filteredMatterNumber, myMatterNumber);

            // Set view as default
            _outlook.Oc.Header.NavigateBack();
            myMattersList.OpenListOptionsMenu().SetCurrentViewAsDefault();

            // Verify current view is set as default
            myMattersList.OpenListOptionsMenu().RestoreDefaults();
            Assert.That(myMattersList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(myMattersList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");

            // Clear user default
            myMattersList.OpenListOptionsMenu().ClearUserDefault();
            Assert.That(myMattersList.IsFilterIconVisible, Is.False, "Filter Icon is visible");

            // Apply filter - Matter Type
            myMattersList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            myMattersFilter.Controls["Matter Type"].Set(myMatterType);
            myMattersFilter.Apply();
            Assert.That(myMattersList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(myMattersList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(myMattersList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            myMattersList.OpenRandom();
            var filteredMatterType = matterDetails.MatterType;
            Assert.AreEqual(filteredMatterType, myMatterType);

            // Verify filter persist after navigation
            _outlook.Oc.Header.NavigateBack();
            Assert.That(myMattersList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(myMattersList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");

            // Apply filter - Matter Status
            myMattersList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            myMattersFilter.RestoreDefaults();
            Assert.That(myMattersList.IsFilterIconVisible, Is.False, "Filter Icon is visible");

            myMattersList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            myMattersFilter.Controls["Matter Status"].Set(myMatterStatus);
            myMattersFilter.Apply();
            Assert.That(myMattersList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(myMattersList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(myMattersList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            myMattersList.OpenRandom();
            var filteredMatterStatus = matterDetails.Status;
            Assert.AreEqual(filteredMatterStatus, myMatterStatus);

            // Apply filter - Matter Primary Internal Contact
            _outlook.Oc.Header.NavigateBack();
            myMattersList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            myMattersFilter.RestoreDefaults();
            Assert.That(myMattersList.IsFilterIconVisible, Is.False, "Filter Icon is visible");

            myMattersList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            myMattersFilter.Controls["Primary Internal Contact"].Set(myMatterPIC);
            myMattersFilter.Apply();
            Assert.That(myMattersList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(myMattersList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(myMattersList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            myMattersList.OpenRandom();
            var filteredMatterPIC = matterDetails.PrimaryInternalContact;
            Assert.AreEqual(filteredMatterPIC, myMatterPIC);

            _outlook.Oc.Header.NavigateBack();
            myMattersList.OpenListOptionsMenu().RestoreDefaults();
            Assert.That(myMattersList.IsFilterIconVisible, Is.False, "Filter Icon is visible");

            // Apply filter - Matter PABU
            myMattersList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            myMattersFilter.Controls["Practice Area - Business Unit"].Set(myMatterPABU);
            myMattersFilter.Apply();
            Assert.That(myMattersList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(myMattersList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(myMattersList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            myMattersList.OpenRandom();
            var filteredMatterPABU = matterDetails.PracticeAreaBusinessUnit;
            Assert.AreEqual(filteredMatterPABU, myMatterPABU);

            _outlook.Oc.Header.NavigateBack();
            myMattersList.OpenListOptionsMenu().RestoreDefaults();
            Assert.That(myMattersList.IsFilterIconVisible, Is.False, "Filter Icon is visible");
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("Test case reference: Filters & Smoke Test Views")]
        public void MatterEmailsFilters()
        {
            var mattersList = _mattersListPage.ItemList;
            var subject1 = _outlook.AddTestEmailsToFolder(1, FileSize.VerySmall, true, OfficeApp.Word).First().Key;
            _outlook.OpenTestEmailFolder();

            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var emailsListPage = _outlook.Oc.EmailListPage;
            var emailsList = emailsListPage.ItemList;
            var emailsFilterDialog = emailsListPage.EmailsListFilterDialog;

            _outlook.SelectNthItem(0);

            //Select random Matter & Open it.
            mattersList.OpenRandom();
            matterDetails.Tabs.Open("Emails");

            // quick file test emails
            matterDetails.QuickFile();
            var filedEmail1 = emailsList.GetEmailListItemFromText(subject1);
            var subject2 = _outlook.AddTestEmailsToFolder(1).First().Key;
            _outlook.SelectNthItem(0);
            matterDetails.QuickFile();
            var filedEmail2 = emailsList.GetEmailListItemFromText(subject2);

            //Select random email
            var randomInt = GetRandomNumber(1);

            //Fetch Email Properties.
            var senderName = GetRandomSubstring(filedEmail1.From);
            var emailSubject = GetRandomSubstring(randomInt == 0 ? filedEmail1.Subject : filedEmail2.Subject);
            var emailBody = GetRandomSubstring(randomInt == 0 ? filedEmail1.EmailBody : filedEmail2.EmailBody);
            var hasAttachment = randomInt == 0 ? filedEmail1.HasAttachment : filedEmail2.HasAttachment;
            var receivedDate = filedEmail1.ReceivedTime.Date;

            Assert.That(emailsListPage.ItemList.ListOptionsToolTip, Is.EqualTo(ListOptionsToolTip), "List options tooltip differs from expected");

            // Apply filter - Sender Name
            emailsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            emailsFilterDialog.Controls["Sender Name"].Set(senderName);
            emailsFilterDialog.Apply();
            Assert.That(emailsList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(emailsList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(emailsList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            var filteredList = emailsList.GetAllEmailListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by Sender Name - {senderName}");
            Assert.That(filteredList, Has.All.Property(nameof(EmailListItem.From)).Contains(senderName).IgnoreCase, $"Filtered list has items not containing Sender Name - {senderName}");

            // Save View
            emailsList.OpenListOptionsMenu().SaveCurrentView();
            emailsListPage.SaveCurrentViewDialog.Controls["Create New"].Set(ViewName);
            emailsListPage.SaveCurrentViewDialog.Save();
            var savedViewEmailsCount = emailsList.GetCount();

            // Clear Filters
            emailsList.OpenListOptionsMenu().RestoreDefaults();
            Assert.That(emailsList.IsFilterIconVisible, Is.False, "Filter Icon visible");

            emailsList.OpenListOptionsMenu().ApplySavedView(ViewName);
            Assert.AreEqual(savedViewEmailsCount, emailsList.GetCount());

            // Apply filter - Subject
            emailsList.OpenListOptionsMenu().RestoreDefaults();
            emailsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            emailsFilterDialog.Controls["Subject"].Set(emailSubject);
            emailsFilterDialog.Apply();
            Assert.That(emailsList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(emailsList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(emailsList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            filteredList = emailsList.GetAllEmailListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by Subject - {emailSubject}");
            Assert.That(filteredList, Has.All.Property(nameof(EmailListItem.Subject)).Contains(emailSubject).IgnoreCase, $"Filtered list has items not containing Subject - {emailSubject}");

            // Set view as default
            emailsList.OpenListOptionsMenu().SetCurrentViewAsDefault();

            // Verify current view is set as default
            emailsList.OpenListOptionsMenu().RestoreDefaults();
            Assert.That(emailsList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(emailsList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");

            // Clear user default
            emailsList.OpenListOptionsMenu().ClearUserDefault();
            Assert.That(emailsList.IsFilterIconVisible, Is.False, "Filter Icon is visible");

            // Apply filter - Email Body
            emailsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            emailsFilterDialog.Controls["Email Body"].Set(emailBody);
            emailsFilterDialog.Apply();
            Assert.That(emailsList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(emailsList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(emailsList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            filteredList = emailsList.GetAllEmailListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by Email Body - {emailBody}");
            Assert.That(filteredList, Has.All.Property(nameof(EmailListItem.EmailBody)).Contains(emailBody).IgnoreCase, $"Filtered list has items not containing Email Body - {emailBody}");

            // Apply filter - Has Attachment
            emailsList.OpenListOptionsMenu().RestoreDefaults();
            emailsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            emailsFilterDialog.Controls["Has Attachment"].Set(hasAttachment);
            emailsFilterDialog.Apply();
            Assert.That(emailsList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(emailsList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(emailsList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            filteredList = emailsList.GetAllEmailListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by Has Attachment - {hasAttachment}");
            Assert.That(filteredList.Where(x => !x.IsFolder), Has.All.Property(nameof(EmailListItem.HasAttachment)).EqualTo(hasAttachment), $"Filtered list has items other than Has Attachment - {hasAttachment}");

            //Apply Filter - Received Date
            emailsList.OpenListOptionsMenu().RestoreDefaults();
            var dateTo = receivedDate.AddDays(3);
            emailsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            emailsFilterDialog.Controls["Received Date"].Set(FormatDateRange(receivedDate, dateTo));
            emailsFilterDialog.Apply();
            Assert.That(emailsList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(emailsList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(emailsList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            var endDate = dateTo.AddSeconds(-1).AddDays(1);
            filteredList = emailsList.GetAllEmailListItems();
            Assert.That(filteredList, Is.Not.Empty, $"(No records found on applying filter by Received At : {receivedDate} - {dateTo}");
            Assert.That(filteredList, Has.All.Property(nameof(EmailListItem.ReceivedTime)).InRange(receivedDate, endDate), $"Filtered list has items out of Received At : {receivedDate} - {endDate}");

            // Verify filter persist after navigation
            matterDetails.Tabs.Open("Documents");
            _outlook.Oc.Header.NavigateBack();
            Assert.That(emailsList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(emailsList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");

            // clean up
            emailsList.OpenListOptionsMenu().RestoreDefaults();
            filedEmail1 = emailsList.GetEmailListItemFromText(subject1);
            filedEmail1.Delete().Confirm();
            filedEmail2 = emailsList.GetEmailListItemFromText(subject2);
            filedEmail2.Delete().Confirm();
        }

        [Test]
        [Category(SmokeTestCategory)]
        [Description("Test case reference: Filters & Smoke Test Views")]
        public void MatterDocumentsFilters()
        {
            const int uniqueDocumentsCount = 2;
            IDictionary<string, string> emails = new Dictionary<string, string>();

            var mattersList = _mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var documentListPage = _outlook.Oc.DocumentsListPage;
            var documentsList = documentListPage.ItemList;

            //Select random Matter & Open it.
            mattersList.OpenRandom();
            matterDetails.Tabs.Open("Documents");

            var docs = documentsList.GetAllMatterDocumentListItems();

            if (docs.Where(x => !x.IsFolder).GroupBy(x => x.Name).Count() < uniqueDocumentsCount)
            {
                // Upload documents
                emails = _outlook.AddTestEmailsToFolder(uniqueDocumentsCount, FileSize.VerySmall, true);
                _outlook.OpenTestEmailFolder();
                _outlook.TurnOnReadingPane();

                for (var i = 0; i < emails.Count; i++)
                {
                    _outlook.SelectNthItem(i);
                    var filename = new FileInfo(emails.ElementAt(i).Value).Name;
                    var attachment = _outlook.GetAttachmentFromReadingPane(filename);
                    DragAndDrop.FromElementToElement(attachment, documentListPage.DropPoint.GetElement());
                    documentListPage.AddDocumentDialog.UploadDocument();

                    var uploadedDocument = documentsList.GetMatterDocumentListItemFromText(filename);
                    Assert.IsNotNull(uploadedDocument);
                }
            }

            //Select random Document
            var unfilteredCount = documentsList.GetCount();
            var randomDoc = GetRandomNumber(unfilteredCount - 1);
            var selectedDocument = documentsList.GetMatterDocumentListItemByIndex(randomDoc);

            //Fetch Document Properties
            var fileName = GetRandomSubstring(selectedDocument.DocumentFileName);
            var name = GetRandomSubstring(selectedDocument.Name);
            var status = selectedDocument.Status;
            var updatedBy = selectedDocument.LastModifiedBy;
            var dateUpdated = selectedDocument.UpdatedAt.Date;

            //Apply Filter - FileName
            documentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            var documentsFilterDialog = documentListPage.MatterDocumentsListFilterDialog;
            documentsFilterDialog.Controls["Name"].Set(fileName);
            documentsFilterDialog.Apply();
            Assert.That(documentsList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(documentsList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(documentsList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            var filteredList = documentsList.GetAllMatterDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by FileName - {fileName}");
            Assert.That(filteredList, Has.All.Property(nameof(MatterDocumentListItem.DocumentFileName)).Contains(fileName).IgnoreCase, $"Filtered list has items not containing Name - {fileName}");

            // Save View
            documentsList.OpenListOptionsMenu().SaveCurrentView();
            documentListPage.SaveCurrentViewDialog.Controls["Create New"].Set(ViewName);
            documentListPage.SaveCurrentViewDialog.Save();

            // Clear Filters
            documentsList.OpenListOptionsMenu().RestoreDefaults();
            Assert.That(documentsList.IsFilterIconVisible, Is.False, "Filter Icon is visible");

            // Apply Saved View
            documentsList.OpenListOptionsMenu().ApplySavedView(ViewName);
            Assert.AreEqual(filteredList.Count, documentsList.GetCount());

            // Set view as default
            documentsList.OpenListOptionsMenu().SetCurrentViewAsDefault();

            // Verify current view is set as default
            documentsList.OpenListOptionsMenu().RestoreDefaults();
            Assert.That(documentsList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(documentsList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");

            // Clear user default
            documentsList.OpenListOptionsMenu().ClearUserDefault();
            Assert.That(documentsList.IsFilterIconVisible, Is.False, "Filter Icon is visible");

            //Apply Filter - Name
            documentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Name"].Set(name);
            documentsFilterDialog.Apply();
            Assert.That(documentsList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(documentsList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(documentsList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            filteredList = documentsList.GetAllMatterDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by Name - {name}");
            Assert.That(filteredList, Has.All.Property(nameof(MatterDocumentListItem.Name)).Contains(name).IgnoreCase, $"Filtered list has items not containing Name - {name}");

            //Apply Filter - Status
            documentsList.OpenListOptionsMenu().RestoreDefaults();
            documentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Status"].Set(status);
            documentsFilterDialog.Apply();
            Assert.That(documentsList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(documentsList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(documentsList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            filteredList = documentsList.GetAllMatterDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by Status - {status}");
            Assert.That(filteredList, Has.All.Property(nameof(MatterDocumentListItem.Status)).EqualTo(status));

            //Apply Filter - Updated By
            documentsList.OpenListOptionsMenu().RestoreDefaults();
            documentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Updated By"].Set(updatedBy);
            documentsFilterDialog.Apply();
            Assert.That(documentsList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(documentsList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(documentsList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            filteredList = documentsList.GetAllMatterDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"No records found on applying filter by Updated By - {updatedBy}");
            Assert.That(filteredList.Where(x => !x.IsFolder), Has.All.Property(nameof(MatterDocumentListItem.LastModifiedBy)).EqualTo(updatedBy), $"Filtered list has items with update by user other than- {updatedBy}");

            //Apply Filter - Updated At
            documentsList.OpenListOptionsMenu().RestoreDefaults();
            var dateTo = dateUpdated.AddDays(3);
            documentsList.OpenListOptionsMenu().OpenCreateListFilterDialog();
            documentsFilterDialog.Controls["Updated At"].Set(FormatDateRange(dateUpdated, dateTo));
            documentsFilterDialog.Apply();
            Assert.That(documentsList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(documentsList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");
            Assert.That(documentsList.FilterIconToolTip, Is.EqualTo(FilterIconToolTip), "Filter Icon tooltip differs from expected");

            var endDate = dateTo.AddSeconds(-1).AddDays(1);
            filteredList = documentsList.GetAllMatterDocumentListItems();
            Assert.That(filteredList, Is.Not.Empty, $"(No records found on applying filter by Updated At : {dateUpdated} - {dateTo}");
            Assert.That(filteredList.Where(x => !x.IsFolder), Has.All.Property(nameof(MatterDocumentListItem.UpdatedAt)).InRange(dateUpdated, endDate), $"Filtered list has items out of Updated At : {dateUpdated} - {endDate}");

            // Verify filter persist after navigation
            matterDetails.Tabs.Open("Emails");
            _outlook.Oc.Header.NavigateBack();
            Assert.That(documentsList.IsFilterIconVisible, Is.True, "Filter Icon is not visible");
            Assert.That(documentsList.GetFilterIconColor().Name, Is.EqualTo(BlueColorName), "Filter Icon is not expected color");

            // clean up
            for (var i = 0; i < emails.Count; i++)
            {
                var filename = new FileInfo(emails.ElementAt(i).Value).Name;
                var document = documentsList.GetMatterDocumentListItemFromText(filename);
                document.Delete().Confirm();
            }
        }

        [TearDown]
        public void TearDown()
        {
            SaveScreenShotsAndLogs(_outlook);
           _outlook?.Destroy();
        }
    }
}
