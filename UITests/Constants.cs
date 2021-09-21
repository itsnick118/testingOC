using System;
using System.Collections.Generic;

namespace UITests
{
    public static class Constants
    {
        // registry constants
        public const string CpuInfoRegistryKey = "HARDWARE\\DESCRIPTION\\System\\CentralProcessor";

        public const string FirstCoreInfoRegistryKey = "HARDWARE\\DESCRIPTION\\System\\CentralProcessor\\0";
        public const string OfficeRegistryKey = "Outlook.Application\\CurVer";

        // iteration constants
        public const int ReloadIterations = 50;

        public const int MassEmailCount = 25;
        public const int InspectorWindowCount = 10;
        public const int LoginRetryLimit = 3;

        // time constants
        public const int ShortTimeOutSeconds = 10;

        public const int WarmupTimeMilliseconds = 10000;
        public const int CooldownTimeMilliseconds = 10000;
        public const int NormalTimeoutSeconds = 60;
        public const int LongTimeoutSeconds = 10 * 60;

        public static Dictionary<OfficeApp, string> ProcessName = new Dictionary<OfficeApp, string>
        {
            { OfficeApp.Outlook, "OUTLOOK" },
            { OfficeApp.Word, "WINWORD" },
            { OfficeApp.Powerpoint, "POWERPNT" },
            { OfficeApp.Excel, "EXCEL" }
        };

        // category constants
        public const string SmokeTestCategory = "Smoke";

        public const string RegressionTestCategory = "Regression";
        public const string DataDependentTestCategory = "Data-dependent";
        public const string GlobalDocumentsTestCategory = "Global Documents";

        // color constants
        public const string BlackColorName = "ff000000";

        public const string WhiteColorName = "ffffffff";
        public const string BlueColorName = "ff1976d2";
        public const string RedColorName = "fff44336";

        // tooltip constants
        public const string FilterIconToolTip = "Filter Applied";

        public const string ListOptionsToolTip = "List Options";

        // input constants
        public const string ViewName = "Automated view name";
        public const string LengthyViewName = "Automated view name(New Lengthy Version)";
        public const string AutomatedComment = "Automated comment";

        // info messages
        public const string CancelMessage = "Are you sure you want to close without saving your changes?";

        public const string DeleteMessage = "Are you sure you want to delete this item?";
        public const string ConfirmationMessageHeader = "Please Confirm";
        public const string FieldIsRequiredWarning = "Field is required.";
        public const string NarrativeDuplicateTestMessage = "The combination of Narrative Type, Date, and Description already exists.";
        public const string NarrativeErrorMessage = "Unable to create narrative. Email does not have a valid subject.";
        public const string FutureStartDateMessage = "startDate - Start date cannot be after today's date.";
        public const string OverlappingTimePeriodMessage = "Overlapping time period with existing record is detected.";
        public const string NoRecordsFound = "No records found";
        public const string UploadSuccessMessage = "Upload was successful.";
        public const string CheckedOutDocumentDeleteMessage = "Unable to delete document. Document is checked out. Please check in first and try again.";
        public const string CheckedOutDocumentRenameErrorMessage = "Unable to rename document. Document is checked out. Please check in first and try again.";
        public const string SpecialCharacterErrorMessage = "Folder and document file names can't contain the following special characters : < > : \" / \\ | ? * # ^";
        public const string RenameFolderHasCheckedOutDocumentErrorMessage = "This folder contains checked-out documents. Check-in all documents and then rename the folder.";
        public const string DeleteFolderHasCheckedOutDocumentErrorMessage = "This folder contains checked-out documents. Check-in all documents and then delete the folder.";
        public const string DocumentSummaryQuickFileMessage = "Check In was successful.";
        public const string UnsupportedFileErrorMessage = "Unsupported document type";

        public static string DeleteDocumentMessageForDifferentUser(string name)
        {
            return $"Unable to delete document. Document is checked out by {name}.";
        }

        public static string ViewDocumentWarningMessage(string name)
        {
            return $"Document is already checked out by {name}. The last checked in version will be attached, downloaded, or opened in read only mode.";
        }

        // edit dialog buttons
        public static readonly string[] editPopupButtons = new[] { "Reset", "Save", "Cancel" };

        // initial content of document
        public const string InitialDefaultContent = "Initial automatically created content sample.";

        public static string UploadDocumentMessage(string newFileName, string oldFileName)
        {
            return $"{newFileName} will be uploaded as a new version of {oldFileName}" + Environment.NewLine +
                   Environment.NewLine + " Warning: The File name of the new version is different than the original file name.";
        }

        public static string DeleteDocumentMessage(string fileName)
        {
            return $"Are you sure you want to delete '{fileName}' from the matter?" + Environment.NewLine +
                   "This action cannot be undone.";
        }

        public static string MessageonTryingToRemovePIC(string personName)
        {
            return $"Action cannot be performed. Please check permissions and confirm {personName} is not primary internal contact.";
        }

        public static string EmailDuplicateMessage(string subject, string matterName)
        {
            return $"A duplicate of email document with subject ‘{subject}’ is already associated to the matter ‘{matterName}’.";
        }
    }
}
