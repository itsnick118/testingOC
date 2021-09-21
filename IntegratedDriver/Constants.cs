using System.Text.RegularExpressions;

namespace IntegratedDriver
{
    public static class Constants
    {
        public const int FindElementTimeout = 15;
        public const int FileDownloadTimeout = 60;
        public const int RetryLimit = 100;
        public const int OcTransitionWidth = 609;

        public const string AddInTitle = "Wolters Kluwer  |  Office Companion";
        public const string NewFolderMenuItemName = "New Folder...";
        public const string InboxFolderName = "Inbox";
        public const string TestEmailFolderName = "MSUIA_test_emails";
        public const string TestDataFolderName = "TestData";
        public const string TestEmailPrefix = "Test email -- ";
        public const string NewEmailWindowTitle = "Untitled";
        public const string RecepientTo = "To";
        public const string NewEmailPageContent = "Page 1 content";
        public const string EmailSubject = "Subject";
        public const string ViewMoreOnMsExchange = "Click here to view more on Microsoft Exchange";

        public static Regex MailSubjectPattern = new Regex($@"{TestEmailPrefix}([a-zA-Z0-9]+),");

        public const string Utf8subject = @"UTF 8 ¡ ¢ £ ¤ ¥ ¦ § ¨ © ª « ¬  ® ¯ ° ± ² ³ ´ µ ¶ · ¸ ¹ º » ¼ ½ ¾ ¿ À Á Â Ã Ä Å Æ Ç È É Ê
            Ë Ì Í Î Ï Ð Ñ Ò Ó Ô Õ Ö × Ø Ù Ú Û Ü Ý Þ ß à á â ã ä å æ ç è é ê ë ì í î ï ð ñ ò ó ô õ ö ÷ ø ù ú û ü ý þ ÿ畑㵲応ｲ瑤㵨㠱";

        public const string Utf8content = @"Utf-8 symbols should be supported across all functionalities of office companion
    Use at least the following characters to Verify listed below areas:
    1) ¡ ¢ £ ¤ ¥ ¦ § ¨ © ª « ¬ ® ¯ ° ± ² ³ ´ µ ¶ · ¸ ¹ º » ¼ ½ ¾ ¿ À Á Â Ã Ä Å Æ Ç È É Ê Ë Ì Í Î Ï Ð Ñ Ò Ó Ô Õ Ö × Ø Ù Ú Û Ü Ý Þ ß à á â ã ä å æ ç è é ê ë ì í î ï ð ñ ò ó ô õ ö ÷ ø ù ú û ü ý þ ÿ
    2) 畑㵲応ｲ瑤㵨㠱
    3) ÄäÖöÜüß
    4) ÉéÀàÈèÙùÂâÊêÎîÔôÛûÇçËëÏïÜüŸÿ
    5) ÑñÓóÁáéí
    6) йцукен
    7) Special case &asd;
    8) Special case <div class='spinner spinner-large'></div>";

        public const string Utf8CharSet1 = "¡ ¢ £ ¤ ¥ ¦ § ¨ © ª « ¬ ® ¯ ° ± ² ³ ´ µ ¶ · ¸ ¹ º » ¼ ½ ¾ ¿ À Á Â Ã Ä Å Æ Ç È É Ê Ë Ì Í Î Ï Ð Ñ Ò Ó Ô Õ Ö × Ø Ù Ú Û Ü Ý Þ ß à á â ã ä å æ ç è é ê ë ì í î ï ð ñ ò ó ô õ ö ÷ ø ù ú û ü ý þ ÿ";
        public const string Utf8CharSet2 = "畑㵲応ｲ瑤㵨㠱";
        public const string Utf8CharSet3 = "ÄäÖöÜüß";
        public const string Utf8CharSet4 = "ÉéÀàÈèÙùÂâÊêÎîÔôÛûÇçËëÏïÜüŸÿ";
        public const string Utf8CharSet5 = "ÑñÓóÁáéí";
        public const string Utf8CharSet6 = "йцукен";
        public const string Utf8CharSet7 = "&asd;";
        public const string Utf8CharSet8 = "<div class='spinner spinner-large'></div>";
        public const string NonUnicodeCharSet = "èòf  dØ";
        public const string SampleLongName = "_IsaacBAsimov war ein ausgezboeichneter Science Fiction Author der die Basis fuer Roboter Entwicklung gelegt hat";
        public const string SpecialCharset = "<>:\"\\/? *#{}%~&^";
        public const string SpecialCharsetInFolderCreation = "<>:\"\\/|?*#^";
    }
}
