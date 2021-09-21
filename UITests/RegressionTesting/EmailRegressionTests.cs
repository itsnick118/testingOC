using System;
using System.IO;
using System.Linq;
using IntegratedDriver;
using NUnit.Framework;
using UITests.PageModel;
using UITests.PageModel.Shared;
using static IntegratedDriver.Constants;
using static UITests.Constants;

namespace UITests.RegressionTesting
{
    [TestFixture]
    public class EmailRegressionTests : UITestBase
    {
        private Outlook _outlook;
        private OutlookEmailForm _outlookEmailForm;

        [SetUp]
        public void SetUp()
        {
            _outlook = new Outlook(TestEnvironment);
            _outlook.Launch();
            _outlook.Oc.BasicSettingsPage.LogInAsStandardUser();
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_16877 Verify Style Email Sub Grid Narrow View")]
        public void StyleEmailSubGrid()
        {
            const string dateTimeFormat = "M-dd-yyyy h-mm-ss tt";

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailPage = _outlook.Oc.MatterDetailsPage;
            var emailListPage = _outlook.Oc.EmailListPage;
            var emailList = emailListPage.ItemList;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenRandom();
            matterDetailPage.Tabs.Open("Emails");

            Assert.IsTrue(emailList.IsQuickSearchIconDisplayed);
            Assert.IsTrue(emailList.IsAddFolderButtonVisible, "Add Folder button is not displayed. Please check if enabled in Passport server.");
            Assert.IsTrue(emailList.IsSortIconVisible);
            Assert.IsTrue(emailList.IsListOptionsDisplayed);

            var subject = _outlook.AddTestEmailsToFolder(1).First().Key;
            _outlook.OpenTestEmailFolder();
            _outlook.SelectNthItem(0);
            matterDetailPage.QuickFile();

            // verify email
            var filedEmail = emailList.GetEmailListItemFromText(subject);
            Assert.IsNotNull(filedEmail);

            Assert.IsTrue(filedEmail.HasCheckBox);
            filedEmail.Select();
            Assert.IsTrue(emailListPage.IsDeleteEmailsButtonDisplayed);

            // verify folder
            var newFolderName = DateTime.Now.ToString(dateTimeFormat);
            emailList.OpenAddFolderDialog();
            emailListPage.AddFolderDialog.Controls["Name"].Set(newFolderName);
            emailListPage.AddFolderDialog.Save();
            var testFolder = emailList.GetEmailListItemFromText(newFolderName);
            Assert.IsNotNull(testFolder);

            Assert.IsTrue(testFolder.HasRenameButton);
            Assert.IsTrue(testFolder.HasQuickFileButton);
            Assert.IsTrue(testFolder.IsDeleteButtonVisible());

            Assert.AreEqual(emailList.GetFooterCount(), emailList.GetAllEmailListItems().Count);

            testFolder.Open();
            var breadcrumbsPath = emailListPage.BreadcrumbsControl.GetCurrentPath();
            Assert.IsTrue(breadcrumbsPath.EndsWith(newFolderName));

            emailListPage.BreadcrumbsControl.NavigateToTheRoot();

            // cleanup
            emailList.GetEmailListItemFromText(newFolderName).Delete().Confirm();
            emailList.GetEmailListItemFromText(subject).Delete().Confirm();
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_16883 Delete emails from a matter")]
        public void DeleteEmailsFromMatter()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailPage = _outlook.Oc.MatterDetailsPage;
            var emailListPage = _outlook.Oc.EmailListPage;
            var emailList = emailListPage.ItemList;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenRandom();

            var emails = _outlook.AddTestEmailsToFolder(3);
            string[] subjects = new string[3];

            _outlook.OpenTestEmailFolder();
            _outlook.SelectAllItems();

            for (var i = 0; i < emails.Count; i++)
            {
                subjects.SetValue(emails.ElementAt(i).Key, i);
            }

            matterDetailPage.Tabs.Open("Emails");
            var existingEmailsCount = emailList.GetAllEmailListItems().Where(x => !x.IsFolder).ToList().Count;

            // Upload emails
            matterDetailPage.QuickFile();
            Assert.AreEqual(existingEmailsCount + 3, emailList.GetAllEmailListItems().Where(x => !x.IsFolder).ToList().Count);

            var testEmails = emailList.GetAllEmailListItems().Where(x => !x.IsFolder && subjects.Any(x.Subject.Contains)).ToList();
            var firstEmail = testEmails[0];
            var firstEmailSubject = firstEmail.Subject;

            // select an email
            firstEmail.Select();
            Assert.IsTrue(emailListPage.IsDeleteEmailsButtonDisplayed);
            Assert.AreEqual(BlueColorName, firstEmail.GetCheckBoxColor().Name, "Selected checkbox does not have expected color");

            // unselect same email
            firstEmail = emailList.GetEmailListItemFromText(firstEmailSubject);
            firstEmail.Select();
            Assert.IsFalse(emailListPage.IsDeleteEmailsButtonDisplayed);
            Assert.AreEqual(BlackColorName, firstEmail.GetCheckBoxColor().Name, "Unselected checkbox does not have expected color");

            // select 2 emails
            firstEmail = emailList.GetEmailListItemFromText(firstEmailSubject);
            firstEmail.Select();

            var secondEmail = testEmails[1];
            var secondEmailSubject = secondEmail.Subject;
            secondEmail.Select();

            Assert.IsTrue(emailListPage.IsDeleteEmailsButtonDisplayed);
            Assert.AreEqual(BlueColorName, firstEmail.GetCheckBoxColor().Name, "Selected checkbox does not have expected color");
            Assert.AreEqual(BlueColorName, secondEmail.GetCheckBoxColor().Name, "Selected checkbox does not have expected color");

            // open delete email dialog and cancel
            firstEmail.Delete();
            Assert.IsTrue(emailListPage.DeleteEmailDialog.IsDisplayed());
            emailListPage.DeleteEmailDialog.Cancel();

            // selected emails are selected
            firstEmail = emailList.GetEmailListItemFromText(firstEmailSubject);
            secondEmail = emailList.GetEmailListItemFromText(secondEmailSubject);
            Assert.AreEqual(BlueColorName, firstEmail.GetCheckBoxColor().Name, "Selected checkbox does not have expected color");
            Assert.AreEqual(BlueColorName, secondEmail.GetCheckBoxColor().Name, "Selected checkbox does not have expected color");

            // select third email
            var thirdEmail = testEmails[2];
            var thirdEmailSubject = thirdEmail.Subject;
            thirdEmail.Select();
            Assert.AreEqual(BlueColorName, thirdEmail.GetCheckBoxColor().Name, "Selected checkbox does not have expected color");

            // delete multiple selected emails
            emailListPage.DeleteEmails();
            firstEmail = emailList.GetEmailListItemFromText(firstEmailSubject, false);
            Assert.IsNull(firstEmail);

            secondEmail = emailList.GetEmailListItemFromText(secondEmailSubject, false);
            Assert.IsNull(secondEmail);

            thirdEmail = emailList.GetEmailListItemFromText(thirdEmailSubject, false);
            Assert.IsNull(thirdEmail);
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test Case Reference : TC_16929 Verify upload of emails to folder")]
        public void UploadEmailsToFolder()
        {
            const string dateTimeFormat = "M-dd-yyyy h-mm-ss tt";

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var emailListPage = _outlook.Oc.EmailListPage;
            var emailList = emailListPage.ItemList;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetailsPage.Tabs.Open("Emails");

            emailList.OpenAddFolderDialog();
            var addFolderDialog = emailListPage.AddFolderDialog;
            var testFolderName = DateTime.Now.ToString(dateTimeFormat);
            addFolderDialog.Controls["Name"].Set(testFolderName);
            addFolderDialog.Save();

            // dnd email to test folder
            var subject1 = _outlook.AddTestEmailsToFolder(1).First().Key;
            _outlook.OpenTestEmailFolder();
            var testFolder = emailList.GetEmailListItemFromText(testFolderName);
            _outlook.DragAndDropEmailToOc(_outlook.GetNthItem(0), testFolder.DropPoint);

            // quick file emails to test folder
            var subject2 = _outlook.AddTestEmailsToFolder(1).First().Key;
            testFolder = emailList.GetEmailListItemFromText(testFolderName);
            testFolder.QuickFile();

            // verify uploaded emails in test folder
            testFolder = emailList.GetEmailListItemFromText(testFolderName);
            testFolder.Open();
            Assert.IsNotNull(emailList.GetEmailListItemFromText(subject1));
            Assert.IsNotNull(emailList.GetEmailListItemFromText(subject2));

            // delete emails in test folder
            foreach (var email in emailList.GetAllEmailListItems())
            {
                email.Select();
            }
            emailListPage.DeleteEmails();

            Assert.IsNull(emailList.GetEmailListItemFromText(subject1, false));
            Assert.IsNull(emailList.GetEmailListItemFromText(subject2, false));

            // create sub folder
            emailList.OpenAddFolderDialog();
            var subFolderName = DateTime.Now.AddMinutes(5).ToString(dateTimeFormat);
            addFolderDialog.Controls["Name"].Set(subFolderName);
            addFolderDialog.Save();

            // dnd email to sub folder
            var subject3 = _outlook.AddTestEmailsToFolder(1).First().Key;
            _outlook.OpenTestEmailFolder();
            var subFolder = emailList.GetEmailListItemFromText(subFolderName);
            _outlook.DragAndDropEmailToOc(_outlook.GetNthItem(0), subFolder.DropPoint);

            // quick file emails to sub folder
            var subject4 = _outlook.AddTestEmailsToFolder(1).First().Key;
            subFolder = emailList.GetEmailListItemFromText(subFolderName);
            subFolder.QuickFile();

            // verify uploaded emails in sub folder
            subFolder = emailList.GetEmailListItemFromText(subFolderName);
            subFolder.Open();
            Assert.IsNotNull(emailList.GetEmailListItemFromText(subject3));
            Assert.IsNotNull(emailList.GetEmailListItemFromText(subject4));

            // delete emails in sub folder
            foreach (var email in emailList.GetAllEmailListItems())
            {
                email.Select();
            }
            emailListPage.DeleteEmails();

            Assert.IsNull(emailList.GetEmailListItemFromText(subject3, false));
            Assert.IsNull(emailList.GetEmailListItemFromText(subject4, false));
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference : TC_16927 Verify OC support for UTF-8 and non-unicode characters in emails")]
        public void EmailCharacterSupport()
        {
            const string dateTimeFormat = "M-dd-yyyy h-mm-ss tt";

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var emailListPage = _outlook.Oc.EmailListPage;
            var emailList = emailListPage.ItemList;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetailsPage.Tabs.Open("Emails");

            var subject = _outlook.AddTestEmailsToFolder(1, subject: Utf8subject, fileContent: Utf8content).First().Key;
            _outlook.OpenTestEmailFolder();
            matterDetailsPage.QuickFile();

            var filedEmail = emailList.GetEmailListItemFromText(subject);
            Assert.IsNotNull(filedEmail);

            // quick search
            emailListPage.QuickSearch.SearchBy(Utf8CharSet1);
            Assert.That(emailList.GetAllEmailListItems(), Has.All.Property(nameof(EmailListItem.EmailBody)).Contains(Utf8CharSet1).IgnoreCase,
                $"Filtered list has items not containing Type - {Utf8CharSet1}");

            emailListPage.QuickSearch.SearchBy(Utf8CharSet2);
            Assert.That(emailList.GetAllEmailListItems(), Has.All.Property(nameof(EmailListItem.EmailBody)).Contains(Utf8CharSet2).IgnoreCase,
                $"Filtered list has items not containing Type - {Utf8CharSet2}");

            emailListPage.QuickSearch.SearchBy(Utf8CharSet3);
            Assert.That(emailList.GetAllEmailListItems(), Has.All.Property(nameof(EmailListItem.EmailBody)).Contains(Utf8CharSet3).IgnoreCase,
                $"Filtered list has items not containing Type - {Utf8CharSet3}");

            emailListPage.QuickSearch.SearchBy(Utf8CharSet4);
            Assert.That(emailList.GetAllEmailListItems(), Has.All.Property(nameof(EmailListItem.EmailBody)).Contains(Utf8CharSet4).IgnoreCase,
                $"Filtered list has items not containing Type - {Utf8CharSet4}");

            emailListPage.QuickSearch.SearchBy(Utf8CharSet5);
            Assert.That(emailList.GetAllEmailListItems(), Has.All.Property(nameof(EmailListItem.EmailBody)).Contains(Utf8CharSet5).IgnoreCase,
                $"Filtered list has items not containing Type - {Utf8CharSet5}");

            emailListPage.QuickSearch.SearchBy(Utf8CharSet6);
            Assert.That(emailList.GetAllEmailListItems(), Has.All.Property(nameof(EmailListItem.EmailBody)).Contains(Utf8CharSet6).IgnoreCase,
                $"Filtered list has items not containing Type - {Utf8CharSet6}");

            emailListPage.QuickSearch.SearchBy(Utf8CharSet7);
            Assert.That(emailList.GetAllEmailListItems(), Has.All.Property(nameof(EmailListItem.EmailBody)).Contains(Utf8CharSet7).IgnoreCase,
                $"Filtered list has items not containing Type - {Utf8CharSet7}");

            emailListPage.QuickSearch.SearchBy(Utf8CharSet8);
            Assert.That(emailList.GetAllEmailListItems(), Has.All.Property(nameof(EmailListItem.EmailBody)).Contains(Utf8CharSet8).IgnoreCase,
                $"Filtered list has items not containing Type - {Utf8CharSet8}");

            emailListPage.QuickSearch.Close();

            // create folder with utf8 characters
            emailList.OpenAddFolderDialog();
            var addFolderDialog = emailListPage.AddFolderDialog;
            var dateTime = DateTime.Now.ToString(dateTimeFormat);
            var testFolderName = dateTime + Utf8CharSet1;
            addFolderDialog.Controls["Name"].Set(testFolderName);
            addFolderDialog.Save();

            var testFolder = emailList.GetEmailListItemFromText(testFolderName);
            Assert.IsNotNull(testFolder);

            testFolder.Open();
            matterDetailsPage.QuickFile();
            filedEmail = emailList.GetEmailListItemFromText(subject);
            Assert.IsNotNull(filedEmail);

            filedEmail.Delete().Confirm();
            Assert.IsNull(emailList.GetEmailListItemFromText(subject, false));

            // create sub folder with utf8 characters
            emailList.OpenAddFolderDialog();
            dateTime = DateTime.Now.AddMinutes(5).ToString(dateTimeFormat);
            var subFolderName = dateTime + Utf8CharSet2;
            addFolderDialog.Controls["Name"].Set(subFolderName);
            addFolderDialog.Save();

            var testSubFolder = emailList.GetEmailListItemFromText(subFolderName);
            Assert.IsNotNull(testSubFolder);

            testSubFolder.Open();
            matterDetailsPage.QuickFile();
            filedEmail = emailList.GetEmailListItemFromText(subject);
            Assert.IsNotNull(filedEmail);

            filedEmail.Delete().Confirm();
            Assert.IsNull(emailList.GetEmailListItemFromText(subject, false));

            // cleanup
            matterDetailsPage.Tabs.Open("Emails");

            testFolder = emailList.GetEmailListItemFromText(testFolderName);
            testFolder.Delete().Confirm();
            Assert.IsNull(emailList.GetEmailListItemFromText(testFolderName, false));

            filedEmail = emailList.GetEmailListItemFromText(subject);
            filedEmail.Delete().Confirm();
            Assert.IsNull(emailList.GetEmailListItemFromText(subject, false));

            // verify email with non unicode characters
            subject = _outlook.AddTestEmailsToFolder(1, fileContent: NonUnicodeCharSet).First().Key;
            _outlook.OpenTestEmailFolder();
            matterDetailsPage.QuickFile();

            filedEmail = emailList.GetEmailListItemFromText(subject);
            Assert.IsNotNull(filedEmail);

            // cleanup
            filedEmail = emailList.GetEmailListItemFromText(subject);
            filedEmail.Delete().Confirm();
            Assert.IsNull(emailList.GetEmailListItemFromText(subject, false));
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description(
           "TC:16832 - To verify Email Quick Search")]
        public void EmailQuickSearch()
        {
            const string dateTimeFormat = "M-dd-yyyy h-mm-ss tt";

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var emailListPage = _outlook.Oc.EmailListPage;
            var emailList = emailListPage.ItemList;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenRandom();
            matterDetailsPage.Tabs.Open("Emails");

            var emailsToUpload = _outlook.AddTestEmailsToFolder(2, useDifferentTemplates: true);
            var subject = emailsToUpload.First().Key;
            _outlook.OpenTestEmailFolder();
            _outlook.SelectAllItems();
            matterDetailsPage.QuickFile();

            // verify email
            var filedEmail = emailList.GetEmailListItemFromText(subject);
            Assert.IsNotNull(filedEmail);

            //Quick search with subject
            emailListPage.QuickSearch.SearchBy(subject);
            var foundEmailsBySubject = emailList.GetAllEmailListItems();
            Assert.GreaterOrEqual(foundEmailsBySubject.Count, 1, "Email with unique subject not found.");
            var firstEmail = foundEmailsBySubject[0];

            //Quick search with sender/From
            emailListPage.QuickSearch.SearchBy(firstEmail.From);
            var foundEmailsBySender = emailList.GetAllEmailListItems();
            Assert.GreaterOrEqual(foundEmailsBySender.Count, 1, "Email with unique sender not found.");

            //Quick search with Email body
            Assert.IsNotNull(firstEmail.EmailBody);
            var emailSubBody = firstEmail.EmailBody.Substring(0, 15);
            Assert.IsNotNull(emailSubBody);
            emailListPage.QuickSearch.SearchBy(emailSubBody);
            var foundEmailsByEmailBody = emailList.GetAllEmailListItems();
            Assert.GreaterOrEqual(foundEmailsByEmailBody.Count, 1, "Email with unique Email Body not found..");
            emailListPage.QuickSearch.Close();

            //Quick search with Email Received date
            var receivedDate = firstEmail.ReceivedTime.ToString("yyyy-MM-dd");
            emailListPage.QuickSearch.SearchBy(receivedDate);
            var foundEmailsByReceivedDate = emailList.GetAllEmailListItems();
            Assert.GreaterOrEqual(foundEmailsByReceivedDate.Count, 1, "Email with unique Email Received not found.");
            emailListPage.QuickSearch.Close();

            // verify Add folder
            var newFolderName = DateTime.Now.ToString(dateTimeFormat);
            emailList.OpenAddFolderDialog();
            emailListPage.AddFolderDialog.Controls["Name"].Set(newFolderName);
            emailListPage.AddFolderDialog.Save();
            var testFolder = emailList.GetEmailListItemFromText(newFolderName);
            Assert.IsNotNull(testFolder, "Folder not created successfully");
            Assert.GreaterOrEqual(emailList.GetFooterCount(), emailList.GetAllEmailListItems().Count);

            //Upload Email to the Folder
            testFolder.Open();
            matterDetailsPage.QuickFile();
            var breadcrumbsPath = emailListPage.BreadcrumbsControl.GetCurrentPath();
            Assert.IsTrue(breadcrumbsPath.EndsWith(newFolderName));

            //Quick search with subject in the Folder
            emailListPage.QuickSearch.SearchBy(subject);
            var foundEmailBySubject = emailList.GetAllEmailListItems();
            Assert.GreaterOrEqual(foundEmailBySubject.Count, 1, "Email with unique subject not found.");
            emailListPage.QuickSearch.Close();

            //Verify Add sub folder
            var subFolderName = DateTime.Now.ToString(dateTimeFormat);
            emailList.OpenAddFolderDialog();
            emailListPage.AddFolderDialog.Controls["Name"].Set(subFolderName);
            emailListPage.AddFolderDialog.Save();
            Assert.IsNotNull(subFolderName, " Sub Folder not created successfully");
            var breadcrumbPath = emailListPage.BreadcrumbsControl.GetCurrentPath();
            Assert.IsTrue(breadcrumbPath.EndsWith(newFolderName));
            emailListPage.BreadcrumbsControl.NavigateToTheRoot();

            //Verify search with the folder name
            var rootFolder = emailList.GetEmailListItemFromText(newFolderName);
            Assert.IsNotNull(rootFolder.FolderName);
            emailListPage.QuickSearch.SearchBy(rootFolder.FolderName);
            var foundEmailsByFolderName = emailList.GetAllEmailListItems();
            Assert.GreaterOrEqual(foundEmailsByFolderName.Count, 1, "Email with unique folder not found.");
            emailListPage.QuickSearch.Close();

            //cleanup
            emailList.GetEmailListItemFromText(newFolderName).Delete().Confirm();
            emailList.GetEmailListItemFromText(subject).Delete().Confirm();
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_16882 DnD of matter emails from OC to outlook, file system and other matters")]
        public void DragAndDropEmailFromOc()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailPage = _outlook.Oc.MatterDetailsPage;
            var emailListPage = _outlook.Oc.EmailListPage;
            var emailList = emailListPage.ItemList;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenRandom();
            matterDetailPage.Tabs.Open("Emails");

            var subject = _outlook.AddTestEmailsToFolder(1).First().Key;
            _outlook.OpenTestEmailFolder();
            _outlook.SelectNthItem(0);
            matterDetailPage.QuickFile();

            // verify email
            var filedEmail = emailList.GetEmailListItemFromText(subject);
            Assert.IsNotNull(filedEmail);

            var outlookNewEmailForm = _outlook.OpenNewEmail();
            outlookNewEmailForm.Attach(NewEmailWindowTitle);

            DragAndDrop.FromElementToElement(filedEmail.DropPoint, outlookNewEmailForm.GetEmailPageElement());

            var attachment = outlookNewEmailForm.GetAttachment(filedEmail.Subject);
            Assert.IsNotNull(attachment);

            // drag and drop to file system
            var path = Windows.GetWorkingTempFolder();
            DragAndDrop.ToFileSystem(filedEmail.DropPoint, path);
            Assert.IsNotEmpty(path.GetFiles());

            // clean up
            filedEmail.Delete().Confirm();
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference : TC_16886 Upload emails to matters (Copy email when filing : Yes/No)")]
        public void UploadEmailsToMatter()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var emailListPage = _outlook.Oc.EmailListPage;
            var settingsPage = _outlook.Oc.SettingsPage;
            var emailList = emailListPage.ItemList;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenRandom();
            matterDetailsPage.Tabs.Open("Emails");

            // quick file emails to test folder
            var subject = _outlook.AddTestEmailsToFolder(1).First().Key;
            _outlook.OpenTestEmailFolder();
            matterDetailsPage.QuickFile();
            Assert.IsNotNull(emailList.GetEmailListItemFromText(subject));

            // verify copy emails when filing is checked in settings
            _outlook.Oc.OpenSettings();
            settingsPage.OpenMatterManagement();
            Assert.IsTrue(settingsPage.IsCopyEmailsWhenFiling);

            settingsPage.Cancel();

            // delete uploaded email
            emailList.GetEmailListItemFromText(subject).Delete().Confirm();
            Assert.IsNull(emailList.GetEmailListItemFromText(subject, false));

            // uncheck copy emails when filing is checked in settings
            _outlook.Oc.OpenSettings();
            settingsPage.OpenMatterManagement();
            settingsPage.SelectCopyEmailsWhenFiling();
            settingsPage.Apply();

            // upload email
            matterDetailsPage.QuickFile();
            Assert.IsNotNull(emailList.GetEmailListItemFromText(subject));

            // verify email is deleted from outlook folder
            Assert.IsNull(_outlook.GetNthItem(0, 5));

            // clean up
            emailList.GetEmailListItemFromText(subject).Delete().Confirm();
            Assert.IsNull(emailList.GetEmailListItemFromText(subject, false));

            // re select copy emails when filing
            _outlook.Oc.OpenSettings();
            settingsPage.OpenMatterManagement();
            settingsPage.SelectCopyEmailsWhenFiling();
            settingsPage.Apply();
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_16920 Duplicate email upload(Quick file) validations")]
        public void DuplicateQuickFileEmailUpload()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var emailListPage = _outlook.Oc.EmailListPage;
            var emailList = emailListPage.ItemList;

            mattersListPage.Open();
            var matter = mattersListPage.ItemList.GetMatterListItemByIndex(0);
            matter.Open();
            matterDetailsPage.Tabs.Open("Emails");

            var matterName = matterDetailsPage.MatterName;
            var emails = _outlook.AddTestEmailsToFolder(2, useDifferentTemplates: true);
            var subject1 = emails.Keys.ElementAt(1);
            var subject2 = emails.Keys.ElementAt(0);
            _outlook.OpenTestEmailFolder();
            _outlook.SelectNthItem(0);

            // verify Quick file Single email again.
            matterDetailsPage.QuickFile();
            emailList.GetEmailListItemFromText(subject1);
            Assert.NotNull(emailList.GetEmailListItemFromText(subject1), "List does not contain a email subject based on your search");

            // verify Quick file same email again for duplicate message.
            _outlook.SelectAllItems();
            matterDetailsPage.QuickFile();
            var emailDuplicateMessage = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(1, emailDuplicateMessage.Length);
            Assert.Contains(EmailDuplicateMessage(subject1, matterName), emailDuplicateMessage);
            _outlook.Oc.CloseAllToastMessages();

            //Verify second email is upload
            emailList.GetEmailListItemFromText(subject2);
            Assert.NotNull(emailList.GetEmailListItemFromText(subject2),
                "List does not contain a email subject based on your search");

            //Delete emails
            emailList.GetEmailListItemFromText(subject1).Delete().Confirm();
            Assert.IsNull(emailList.GetEmailListItemFromText(subject1, false));
            emailList.GetEmailListItemFromText(subject2).Delete().Confirm();
            Assert.IsNull(emailList.GetEmailListItemFromText(subject2, false));

            //Verify Quick file the same email to different matter.
            mattersListPage.Open();
            mattersListPage.ItemList.OpenRandom();
            matterDetailsPage.Tabs.Open("Emails");
            _outlook.SelectNthItem(0);
            matterDetailsPage.QuickFile();
            emailList.GetEmailListItemFromText(subject1);
            Assert.NotNull(emailList.GetEmailListItemFromText(subject1), "List does not contain a email subject based on your search");

            // Delete Email
            emailList.GetEmailListItemFromText(subject1).Delete().Confirm();
            Assert.IsNull(emailList.GetEmailListItemFromText(subject1, false));

            //Verify Quick file the same email to Favorite matter.
            mattersListPage.Open();
            mattersListPage.SetNthMatterAsFavorite(0);
            mattersListPage.OpenFavoritesList();
            mattersListPage.ItemList.OpenFirst();
            matterDetailsPage.Tabs.Open("Emails");
            _outlook.SelectNthItem(0);
            matterDetailsPage.QuickFile();
            emailList.GetEmailListItemFromText(subject1);
            Assert.NotNull(emailList.GetEmailListItemFromText(subject1), "List does not contain a email subject based on your search");

            //Verify Quick file the same email to Favorite matter for Duplicate error
            matterDetailsPage.QuickFile();
            var emailDuplicateMessages = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(1, emailDuplicateMessage.Length);
            Assert.Contains(EmailDuplicateMessage(subject1, matterName), emailDuplicateMessages);
            _outlook.Oc.CloseAllToastMessages();

            // Delete Email
            emailList.GetEmailListItemFromText(subject1).Delete().Confirm();
            Assert.IsNull(emailList.GetEmailListItemFromText(subject1, false));
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_16920 Duplicate email upload(Dnd) validations")]
        public void DuplicateDragAndDropEmailUpload()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var emailListPage = _outlook.Oc.EmailListPage;
            var emailList = emailListPage.ItemList;

            mattersListPage.Open();
            var matter = mattersListPage.ItemList.GetMatterListItemByIndex(0);
            matter.Open();
            matterDetailsPage.Tabs.Open("Emails");

            var matterName = matterDetailsPage.MatterName;
            var emails = _outlook.AddTestEmailsToFolder(2, useDifferentTemplates: true);
            var subject1 = emails.Keys.ElementAt(1);
            var subject2 = emails.Keys.ElementAt(0);
            _outlook.OpenTestEmailFolder();
            _outlook.SelectNthItem(0);

            //  Drag And Drop single email
            _outlook.DragAndDropEmailToOc(_outlook.GetNthItem(0), matterDetailsPage.DropPoint.GetElement());
            emailList.GetEmailListItemFromText(subject1);
            Assert.NotNull(emailList.GetEmailListItemFromText(subject1), "List does not contain a email subject based on your search");

            // Drag And Drop same email again for the same matter for duplicate message.
            _outlook.SelectAllItems();
            _outlook.DragAndDropEmailToOc(_outlook.GetNthItem(0), matterDetailsPage.DropPoint.GetElement());
            var emailDuplicateMessage = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(1, emailDuplicateMessage.Length);
            Assert.Contains(EmailDuplicateMessage(subject1, matterName), emailDuplicateMessage);
            _outlook.Oc.CloseAllToastMessages();

            //Verify second email is upload
            emailList.GetEmailListItemFromText(subject2);
            Assert.NotNull(emailList.GetEmailListItemFromText(subject2), "List does not contain a email subject based on your search");

            //Delete uploaded emails
            emailList.GetEmailListItemFromText(subject1).Delete().Confirm();
            Assert.IsNull(emailList.GetEmailListItemFromText(subject1, false));
            emailList.GetEmailListItemFromText(subject2).Delete().Confirm();
            Assert.IsNull(emailList.GetEmailListItemFromText(subject2, false));

            //Verify Drag And Drop the same email to different matter.
            mattersListPage.Open();
            mattersListPage.ItemList.OpenRandom();
            matterDetailsPage.Tabs.Open("Emails");
            _outlook.SelectNthItem(0);
            _outlook.DragAndDropEmailToOc(_outlook.GetNthItem(0), matterDetailsPage.DropPoint.GetElement());
            emailList.GetEmailListItemFromText(subject1);
            Assert.NotNull(emailList.GetEmailListItemFromText(subject1), "List does not contain a email subject based on your search");

            // Delete Email
            emailList.GetEmailListItemFromText(subject1).Delete().Confirm();
            Assert.IsNull(emailList.GetEmailListItemFromText(subject1, false));

            //Verify Drag and Drop the same email to Favorite matter.
            mattersListPage.Open();
            mattersListPage.SetNthMatterAsFavorite(0);
            mattersListPage.OpenFavoritesList();
            mattersListPage.ItemList.OpenFirst();
            matterDetailsPage.Tabs.Open("Emails");
            _outlook.SelectNthItem(0);
            _outlook.DragAndDropEmailToOc(_outlook.GetNthItem(0), matterDetailsPage.DropPoint.GetElement());
            emailList.GetEmailListItemFromText(subject1);
            Assert.NotNull(emailList.GetEmailListItemFromText(subject1), "List does not contain a email subject based on your search");

            //Verify Drag and Drop the same email to Favorite matter for Duplicate error
            _outlook.DragAndDropEmailToOc(_outlook.GetNthItem(0), matterDetailsPage.DropPoint.GetElement());
            var emailDuplicateMessages = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(1, emailDuplicateMessage.Length);
            Assert.Contains(EmailDuplicateMessage(subject1, matterName), emailDuplicateMessages);
            _outlook.Oc.CloseAllToastMessages();

            // Delete uploaded Email
            emailList.GetEmailListItemFromText(subject1).Delete().Confirm();
            Assert.IsNull(emailList.GetEmailListItemFromText(subject1, false));
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference : TC_16875 UploadQueue")]
        public void UploadQueue()
        {
            const string expectedTab = "emails";

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var emailListPage = _outlook.Oc.EmailListPage;
            var emailList = emailListPage.ItemList;
            var ocHeader = _outlook.Oc.Header;
            var ocUploadHistoryListPage = _outlook.Oc.UploadHistoryPage;
            var ocUploadHistoryList = ocUploadHistoryListPage.ItemList;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenRandom();
            matterDetailsPage.Tabs.Open("Emails");
            var matterName = matterDetailsPage.MatterName;

            //quick file single email
            var subject = _outlook.AddTestEmailsToFolder(1).First().Key;
            _outlook.OpenTestEmailFolder();
            _outlook.SelectNthItem(0);
            matterDetailsPage.QuickFile();
            Assert.IsNotNull(emailList.GetEmailListItemFromText(subject));

            //quick file single failed email
            matterDetailsPage.QuickFile();
            var emailDuplicateMessage = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(1, emailDuplicateMessage.Length);
            Assert.Contains(EmailDuplicateMessage(subject, matterName), emailDuplicateMessage);
            _outlook.Oc.CloseAllToastMessages();

            //Verify Upload queue
            ocHeader.OpenUploadQueue();
            ocHeader.OpenUploadHistory();

            Assert.AreEqual(2, ocUploadHistoryList.GetCount());
            ocUploadHistoryListPage.ClearUploadHistory();
            Assert.AreEqual(0, ocUploadHistoryList.GetCount());

            ocUploadHistoryListPage.CloseUploadHistory();
            var selectTab = matterDetailsPage.Tabs.GetActiveTab().ToLower();
            Assert.AreEqual(expectedTab, selectTab);

            //Delete Email
            emailList.GetEmailListItemFromText(subject).Delete().Confirm();

            //Before Count of emails
            var beforeCount = emailList.GetFooterCount();

            //Multiple emails upload by Quick file to matter list
            var emailsToUpload = _outlook.AddTestEmailsToFolder(3);
            string[] subjects = new string[3];

            for (var i = 0; i < emailsToUpload.Count; i++)
            {
                subjects.SetValue(emailsToUpload.ElementAt(i).Key, i);
            }

            matterDetailsPage.Tabs.Open("Emails");
            var existingEmailsCount = emailList.GetAllEmailListItems().Where(x => !x.IsFolder).ToList().Count;

            // Upload emails
            matterDetailsPage.QuickFile();
            Assert.AreEqual(existingEmailsCount + 3,
                emailList.GetAllEmailListItems().Where(x => !x.IsFolder).ToList().Count);

            var testEmails = emailList.GetAllEmailListItems()
                .Where(x => !x.IsFolder && subjects.Any(x.Subject.Contains)).ToList();
            var firstEmail = testEmails[0];
            var firstEmailSubject = firstEmail.Subject;
            var secondEmail = testEmails[1];
            var secondEmailSubject = secondEmail.Subject;
            var thirdEmail = testEmails[2];
            var thirdEmailSubject = thirdEmail.Subject;

            //After count of emails
            var afterCount = emailList.GetFooterCount();
            Assert.AreEqual(beforeCount + 3, afterCount, "Footer list count is not incremented upon adding a Emails");

            //Emails are uploaded in the Upload queue.
            ocHeader.OpenUploadQueue();
            ocHeader.OpenUploadHistory();

            Assert.AreEqual(3, ocUploadHistoryList.GetCount());
            ocUploadHistoryListPage.ClearUploadHistory();
            Assert.AreEqual(0, ocUploadHistoryList.GetCount());

            ocUploadHistoryListPage.CloseUploadHistory();
            var selectedTab = matterDetailsPage.Tabs.GetActiveTab().ToLower();
            Assert.AreEqual(expectedTab, selectedTab);

            // Delete emails
            foreach (var email in emailsToUpload)
            {
                var result = emailList.GetEmailListItemFromText(email.Key);
                result.Select();
            }

            emailListPage.DeleteEmails();
            firstEmail = emailList.GetEmailListItemFromText(firstEmailSubject, false);
            Assert.IsNull(firstEmail);

            secondEmail = emailList.GetEmailListItemFromText(secondEmailSubject, false);
            Assert.IsNull(secondEmail);

            thirdEmail = emailList.GetEmailListItemFromText(thirdEmailSubject, false);
            Assert.IsNull(thirdEmail);

            //Emails are uploaded by Drag and Drop to Favorite matter.
            mattersListPage.Open();
            mattersListPage.SetNthMatterAsFavorite(0);
            mattersListPage.OpenFavoritesList();
            mattersListPage.ItemList.OpenFirst();
            matterDetailsPage.Tabs.Open("Emails");

            //Before Count of emails
            var beforeCounts = emailList.GetFooterCount();

            //Emails are uploaded by Drag and Drop to Favorite matter.
            _outlook.SelectAllItems();
            _outlook.DragAndDropEmailToOc(_outlook.GetNthItem(0), matterDetailsPage.DropPoint.GetElement());

            //After count of emails
            var afterCounts = emailList.GetFooterCount();
            Assert.AreEqual(beforeCounts + 3, afterCounts,
                "Footer list count is not incremented upon adding a Emails");

            //Emails are uploaded in the Upload queue by Drag and Drop.
            ocHeader.OpenUploadQueue();
            ocHeader.OpenUploadHistory();

            Assert.AreEqual(3, ocUploadHistoryList.GetCount());
            ocUploadHistoryListPage.ClearUploadHistory();
            Assert.AreEqual(0, ocUploadHistoryList.GetCount());

            ocUploadHistoryListPage.CloseUploadHistory();
            Assert.AreEqual(expectedTab, selectedTab);

            //Delete Emails
            foreach (var email in emailsToUpload)
            {
                var result = emailList.GetEmailListItemFromText(email.Key);
                result.Select();
            }
            emailListPage.DeleteEmails();
            firstEmail = emailList.GetEmailListItemFromText(firstEmailSubject, false);
            Assert.IsNull(firstEmail);
            secondEmail = emailList.GetEmailListItemFromText(secondEmailSubject, false);
            Assert.IsNull(secondEmail);
            thirdEmail = emailList.GetEmailListItemFromText(thirdEmailSubject, false);
            Assert.IsNull(thirdEmail);
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_16973 Email Thread Linking")]
        public void EmailThreadLinking()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var emailListPage = _outlook.Oc.EmailListPage;
            var settingsPage = _outlook.Oc.SettingsPage;
            var emailList = emailListPage.ItemList;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenRandom();
            matterDetailsPage.Tabs.Open("Emails");

            // verify automatically upload emails is unchecked in settings
            _outlook.Oc.OpenSettings();
            settingsPage.OpenMatterManagement();
            Assert.IsFalse(settingsPage.IsAutomaticallyThreadLinking);

            settingsPage.Cancel();

            // quick file emails to test folder
            var subject = _outlook.AddTestEmailsToFolder(1).First().Key;
            _outlook.OpenTestEmailFolder();
            matterDetailsPage.QuickFile();
            Assert.IsNotNull(emailList.GetEmailListItemFromText(subject));

            // reply to email
            _outlookEmailForm = new OutlookEmailForm(TestEnvironment);
            _outlookEmailForm.Reply();
            _outlookEmailForm.Send();

            // verify email is not uploaded on reply
            Assert.IsNull(emailList.GetEmailListItemFromText("RE: " + subject, false));

            // select automatically upload emails option in settings
            _outlook.Oc.OpenSettings();
            settingsPage.OpenMatterManagement();
            settingsPage.SelectEmailThreadLinking();
            settingsPage.Apply();
            settingsPage.OpenMatterManagement();
            Assert.IsTrue(settingsPage.IsAutomaticallyThreadLinking);
            settingsPage.Cancel();

            // reply to email
            _outlookEmailForm = new OutlookEmailForm(TestEnvironment);
            _outlookEmailForm.Reply();
            _outlookEmailForm.Send();
            _outlook.Oc.WaitForQueueComplete();

            // verify emails are uploaded
            Assert.AreEqual(2, emailList.GetAllEmailListItems().Where(s => s.Subject.Equals("RE: " + subject)).Count());

            // clean up
            foreach (var email in emailList.GetAllEmailListItems().Where(s => s.Subject.Contains(subject)))
            {
                email.Select();
            }

            emailListPage.DeleteEmails();
            Assert.IsNull(emailList.GetEmailListItemFromText(subject, false));

            _outlook.Oc.OpenSettings();
            settingsPage.OpenMatterManagement();
            settingsPage.SelectEmailThreadLinking();
            settingsPage.Apply();
            settingsPage.OpenMatterManagement();
            Assert.IsFalse(settingsPage.IsAutomaticallyThreadLinking);
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference : TC_17012 Verify to DND emails or documents from OC to create emails attachment")]
        public void VerifyDndFromOcToOutlook()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailPage = _outlook.Oc.MatterDetailsPage;
            var emailListPage = _outlook.Oc.EmailListPage;
            var emailList = emailListPage.ItemList;
            var documentListPage = _outlook.Oc.DocumentsListPage;
            var documentList = documentListPage.ItemList;
            var testEmail = _outlook.AddTestEmailsToFolder(1, FileSize.VerySmall, true, OfficeApp.Word).First();
            var subject = testEmail.Key;
            var filename = new FileInfo(testEmail.Value).Name;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenRandom();
            matterDetailPage.Tabs.Open("Emails");

            _outlook.OpenTestEmailFolder();
            _outlook.SelectNthItem(0);
            matterDetailPage.QuickFile();

            var attachment = _outlook.GetAttachmentFromReadingPane(filename);
            DragAndDrop.FromElementToElement(attachment, matterDetailPage.DropPoint.GetElement());
            documentListPage.AddDocumentDialog.UploadDocument();

            // verify email
            var filedEmail = emailList.GetEmailListItemFromText(subject);
            Assert.IsNotNull(filedEmail);

            var newMailWindow = _outlook.OpenNewMailWindow();

            _outlook.ToggleTaskPane(newMailWindow);
            _outlook.Oc.SwitchToLastOcInstance();

            matterDetailPage.Tabs.Open("Emails");
            filedEmail = emailList.GetEmailListItemFromText(subject);
            DragAndDrop.FromElementToElement(filedEmail.DropPoint, _outlook.GetNewMailBodyElement(newMailWindow), false);

            // check if email attachment added correctly
            var attachedEmail = emailList.GetEmailListItemFromText(subject);
            Assert.IsNotNull(attachedEmail);

            matterDetailPage.Tabs.Open("Documents");
            var uploadedDocument = documentList.GetMatterDocumentListItemFromText(filename);
            Assert.IsNotNull(uploadedDocument);

            DragAndDrop.FromElementToElement(uploadedDocument.DropPoint, _outlook.GetNewMailBodyElement(newMailWindow), false);

            //  check if email attachment added correctly
            var attachedDocument = documentList.GetMatterDocumentListItemFromText(filename);
            Assert.IsNotNull(attachedDocument);

            // clean up test document
            uploadedDocument.Delete().Confirm();

            // clean up test email
            matterDetailPage.Tabs.Open("Emails");
            filedEmail = emailList.GetEmailListItemFromText(subject);
            filedEmail.Delete().Confirm();
        }

        [TearDown]
        public void TearDown()
        {
            SaveScreenShotsAndLogs(_outlook);
            _outlook?.Destroy();
        }
    }
}
