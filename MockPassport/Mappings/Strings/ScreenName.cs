namespace MockPassport.Mappings.Strings
{
    public class ScreenName
    {
        public static string AsParam(string screenName)
        {
            return ParamKey.ScreenName + "=" + screenName;
        }
        public const string AdjustmentLineItemList =
            "Adjustment Line Item List Screen - OC Spend Management";

        public const string DetailLineItemList =
            "DetailLineItemListScreen - OC Spend Management";

        public const string EmailDocumentList =
            "Email Document List Screen - Office";

        public const string EmailDocumentCmisList =
            "PassportCmisObject - Email Documents List Page - OC";

        public const string GlobalDocumentsCmisList =
            "PassportCmisObject - Document Management - OC";

        public const string InvoicesList = 
            "My Invoices - OC Spend Management";

        public const string MatterDocumentCmisList =
            "PassportCmisObject - Matter Documents List Page - OC";

        public const string MatterEventList = 
            "Matter Event List - OC Matter Management";

        public const string MatterList =
            "Matter List - Office";

        public const string MatterNarrativesList =
            "Narratives List - OC Matter Management";

        public const string MatterPersonList = 
            "Matter Person List Screen - Office";

        public const string MatterSummary =
            "View Matter Summary - OC";

        public const string ToBeAcknowledgedByMe =
            "To be Acknowledged by Me - Office";
    }
}