using UITests.PageModel.Shared.InputControls;

namespace UITests.PageModel.Configurations.GA
{
    public class Dialogs
    {
        public static InputControlList AddTaskDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new Dropdown(app, "Type", "eventType"),
                new InputField(app, "Name", "name"),
                new DateField(app, "Due Date", "dueDate"),
                new DateField(app, "Completed Date", "completedDate"),
                new TextArea(app, "Description", "description"),
                new AutoComplete(app, "Invitees/Assigned To", "people1", panelNumber: 1)
            };
        }

        public static InputControlList AddEventDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new Dropdown(app, "Type", "eventType"),
                new InputField(app, "Subject", "name1"),
                new Dropdown(app, "Category", "categoryType"),
                new DateField(app, "Start Date/Time", "startDateTime"),
                new DateField(app, "End Date/Time", "endDateTime"),
                new InputField(app, "Location", "location"),
                new TextArea(app, "Description", "description"),
                new AutoComplete(app, "Invitees/Assigned To", "people", panelNumber: 1)
            };
        }

        public static InputControlList AddNarrativeDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new Dropdown(app, "Narrative Type", "narrativeType"),
                new TextArea(app, "Description", "description"),
                new DateField(app, "Narrative Date", "narrativeDate"),
                new TextArea(app, "Narrative", "narrative")
            };
        }

        public static InputControlList AddPersonDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new Dropdown(app, "Person Type", "matterPersonType"),
                new AutoComplete(app, "Person", "person"),
                new AutoComplete(app, "Role/Involvement Type", "roleInvolvement"),
                new InputField(app, "Comments", "comments"),
                new DateField(app, "Start Date", "startDate"),
                new DateField(app, "End Date", "endDate"),
                new CheckBox(app, "Active", "isActive")
            };
        }

        public static InputControlList AddFolderDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new InputField(app, "Name", "name")
            };
        }

        public static InputControlList AddDocumentDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new InputField(app, "Document Name", "name"),
                new InputField(app, "File Name", "documentFileName"),
                new InputField(app, "Comments", "comments")
            };
        }

        public static InputControlList RenameDocumentDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new InputField(app, "Name", "name"),
                new InputField(app, "Document File Name", "documentFileName")
            };
        }

        public static InputControlList CheckInDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new InputField(app, "Document Name", "name"),
                new InputField(app, "File Name", "documentFileName"),
                new InputField(app, "Comments", "comments")
            };
        }

        public static InputControlList SaveCurrentViewDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new InputField(app, "Create New", "SavedSearchName"),
                new Dropdown(app, "Update Existing", "SavedSearch")
            };
        }

        public static InputControlList MatterListFilterDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new InputField(app, "Matter Name", "matterName"),
                new InputField(app, "Long Matter Name", "longMatterName"),
                new InputField(app, "Matter Number", "matterNumber"),
                new AutoComplete(app, "Matter Type", "matterType", false),
                new Dropdown(app, "Matter Status", "matterStatus"),
                new AutoComplete(app, "Primary Internal Contact", "primaryInternalContact", true,1),
                new AutoComplete(app, "Practice Area - Business Unit", "practiceAreaBusinessUnit", false, 1),
                new AutoComplete(app, "Matter People", "person", true, 1)
            };
        }

        public static InputControlList MatterDocumentsListFilterDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new InputField(app, "File Name", "documentFileName"),
                new InputField(app, "Name", "name"),
                new Dropdown(app, "Status", "status"),
                new InputField(app, "Updated By", "updatedByUser"),
                new DateField(app, "Updated At", "updatedAt")
            };
        }

        public static InputControlList EmailsListFilterDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new InputField(app, "Sender Name", "senderName"),
                new InputField(app, "Sender Email Address", "senderEmailAddress"),
                new InputField(app, "Subject", "subject"),
                new TextArea(app, "Email Body", "mailBody"),
                new DateField(app, "Received Date", "receivedTime"),
                new Dropdown(app, "Has Attachment", "attachmentPresent")
            };
        }

        public static InputControlList GlobalDocumentsListFilterDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new InputField(app, "File Name", "documentFileName"),
                new InputField(app, "Name", "name"),
                new Dropdown(app, "Status", "status"),
                new DateField(app, "Created At", "createdAt"),
                new DateField(app, "Updated At", "updatedAt"),
                new InputField(app, "Created By", "createdByUser"),
                new InputField(app, "Updated By", "updatedByUser"),
                new InputField(app, "Comment", "comments"),
                new TextArea(app, "Content", "content"),
                new AutoComplete(app, "Matter Type", "matterType", false, 1),
                new AutoComplete(app, "Practice Area - Business Unit", "practiceAreaBusinessUnit", false, 1),
                new InputField(app, "Matter Name", "matterName", 1),
                new AutoComplete(app, "Organizations", "organization", true, 1)
            };
        }

        public static InputControlList InvoiceDocumentsListFilterDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new InputField(app, "File Name", "documentFileName"),
                new InputField(app, "Name", "name"),
                new Dropdown(app, "Status", "status"),
                new DateField(app, "Created At", "createdAt"),
                new DateField(app, "Updated At", "updatedAt"),
                new InputField(app, "Created By", "createdByFullName"),
                new InputField(app, "Updated By", "lastModifiedByFullName"),
                new InputField(app, "Comments", "comments"),
                new TextArea(app, "Content", "content"),
            };
        }

        public static InputControlList InvoiceListFilterDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new InputField(app, "Invoice Number", "invoiceNumber"),
                new DateField(app, "Received Date", "receivedDate"),
                new InputField(app, "Organization Name", "name", 1),
                new InputField(app, "Vendor Id", "vendorId", 1),
                new InputField(app, "Matter Name", "matterName", 1),
                new InputField(app, "Matter Number", "matterNumber", 1)
            };
        }

        public static InputControlList ApproveInvoiceDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new TextArea(app, "Internal Comment", "internalComment"),
                new TextArea(app, "External Comment", "externalComment")
            };
        }

        public static InputControlList RejectInvoiceDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new AutoComplete(app, "Reject Reason Codes", "reasonTypes", false),
                new TextArea(app, "Internal Comment", "internalComment"),
                new TextArea(app, "External Comment", "externalComment")
            };
        }

        public static InputControlList AdjustLineItemDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new Dropdown(app, "Adjustment Type", "operationName"),
                new InputField(app, "Adjustment Value", "deltaValue"),
                new Dropdown(app, "Adjustment Reason", "reason"),
                new TextArea(app, "Adjustment Description", "externalComment")
            };
        }

        public static InputControlList HeaderAdjustmentDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new Dropdown(app, "Adjustment Type", "operationName"),
                new InputField(app, "Adjustment Value", "deltaValue"),
                new Dropdown(app, "Adjustment Reason", "reason"),
                new TextArea(app, "Adjustment Description", "externalComment")
            };
        }

        public static InputControlList RejectLineItemDialogControls(IAppInstance app)
        {
            return new InputControlList
            {
                new Dropdown(app, "Reason Code", "reason"),
                new TextArea(app, "Internal Comment", "internalComment"),
                new TextArea(app, "External Comment", "externalComment")
            };
        }
    }
}
