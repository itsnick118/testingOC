using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using static IntegratedDriver.Constants;
using static UITests.TestHelpers;

namespace UITests.PageModel
{
    public class EmailGenerator
    {
        private static string[] _emailTemplates = { "sender1.msg", "sender2.msg" };
        private static string _sourcePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, TestDataFolderName);

        public IDictionary<string, string> AddTestEmailsToFolder(int numEmails, FileSize fileSize = FileSize.VerySmall, bool addAttachment = false,
            OfficeApp docType = OfficeApp.Notepad, string subject = null, string fileContent = null, bool useDifferentTemplates = false)
        {
            var results = new Dictionary<string, string>();

            var outlook = new Application();
            var templateEmailFullPath = GetTemplateEmailFullPath();

            var testFolder = GetTestEmailFolder(outlook);
            RemoveAllItemsInFolder(testFolder);

            for (var i = 0; i < numEmails; i++)
            {
                if (useDifferentTemplates)
                {
                    var templateName = GetTemplateNameByIndex(i);
                    templateEmailFullPath = Path.Combine(_sourcePath, templateName);
                }

                var email = CreateEmailAndMoveToFolder(outlook, testFolder, templateEmailFullPath, addAttachment, fileContent, fileSize, docType, subject);
                if (email.Key != null)
                {
                    results.Add(email.Key, email.Value);
                }
            }

            Marshal.ReleaseComObject(testFolder);

            return results;
        }

        private static Folder GetTestEmailFolder(_Application application)
        {
            var inbox = application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            foreach (Folder folder in inbox.Folders)
            {
                if (folder.Name == TestEmailFolderName)
                {
                    return folder;
                }
            }

            return (Folder)inbox.Folders.Add(TestEmailFolderName);
        }

        private static void RemoveAllItemsInFolder(MAPIFolder folder)
        {
            while (folder.Items.Count > 0)
            {
                var item = folder.Items.GetFirst();
                item.Delete();
            }
        }

        private static KeyValuePair<string, string> CreateEmailAndMoveToFolder(_Application outlook, MAPIFolder folder,
            string templateEmailFullPath, bool addAttachment, string fileContent, FileSize fileSize, OfficeApp docType, string subject = null)
        {
            KeyValuePair<string, string> email;

            var mailItem = (MailItem)outlook.Session.OpenSharedItem(templateEmailFullPath);
            mailItem.Subject = subject ?? TestEmailPrefix + GetRandomText(16);
            email = new KeyValuePair<string, string>(mailItem.Subject, string.Empty);

            FileInfo attachment = null;

            if (addAttachment)
            {
                attachment = string.IsNullOrEmpty(fileContent)
                    ? CreateDocumentWithRandomText(fileSize, docType)
                    : CreateDocument(docType, fileContent);
                mailItem.Attachments.Add(attachment.FullName);
                mailItem.Body = "See attachment.";

                email = new KeyValuePair<string, string>(mailItem.Subject, attachment.FullName);
            }
            else
            {
                mailItem.Body = string.IsNullOrEmpty(fileContent) ? GetRandomText((int)fileSize) : fileContent;
            }

            mailItem.Move(folder);
            Marshal.ReleaseComObject(mailItem);

            attachment?.Delete();

            return email;
        }

        public static FileInfo GetTestEmailTemplate()
        {
            return new FileInfo(GetTemplateEmailFullPath());
        }

        private static string GetTemplateEmailFullPath(int templateIndex = 0)
        {
            return Path.Combine(_sourcePath, _emailTemplates[templateIndex]);
        }

        private string GetTemplateNameByIndex(int index)
        {
            // Iterate through available templates
            var templatesCount = _emailTemplates.Length;
            var realIndex = index >= templatesCount ? index % templatesCount : index;
            return _emailTemplates[realIndex];
        }
    }
}