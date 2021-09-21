using System;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using IntegratedDriver;
using UITests.PageModel;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace UITests
{
    public class TestHelpers
    {
        private static readonly Random Random = new Random();
        private const string DateTimeDashFormat = "M-dd-yyyy h-mm-ss tt";
        private const string DateFormat = "MM'/'dd'/'yyyy";
        private const string DateTimeFormat = "MM/dd/yyyy hh:mm tt";
        private static CultureInfo usCulture = CultureInfo.CreateSpecificCulture("en-US");

        public static DateTime ParseDateTime(string unCulturedDateTime)
        {
            return DateTime.Parse(unCulturedDateTime, usCulture, DateTimeStyles.None);
        }

        public static string GetRandomText(int bytes)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            var resultString = new char[bytes];

            for (var i = 0; i < resultString.Length; i++)
            {
                resultString[i] = chars[Random.Next(chars.Length)];
            }

            return new string(resultString);
        }

        public static string GetRandomTextWithSpaces(int bytes)
        {
            var text = GetRandomText(bytes);
            if (bytes <= 2)
            {
                return text;
            }

            var split = bytes / 2;
            var chars = text.ToCharArray();
            chars[split] = ' ';

            return new string(chars);
        }

        public static int GetRandomNumber(int max, int min = 0) => Random.Next(min, max);

        public static string GetRandomSubstring(string text) => text.Substring(GetRandomNumber(text.Length)).Trim();

        public static string GetRandomFrom(string[] values) => values[GetRandomNumber(values.Length)];

        public static FileInfo CreateDocumentWithRandomText(FileSize size, OfficeApp officeApp)
        {
            var randomContent = GetRandomText((int)size);
            return CreateDocument(officeApp, randomContent);
        }

        public static decimal GetNumeral(string input)
        {
            return decimal.Parse(Regex.Replace(input, @"[^\d.]", ""));
        }

        public static FileInfo CreateDocument(OfficeApp officeApp, string content = Constants.InitialDefaultContent, string name = null)
        {
            switch (officeApp)
            {
                case OfficeApp.Word:
                    return CreateWordFile(content, name);

                case OfficeApp.Excel:
                    return CreateExcelFile(content, name);

                case OfficeApp.Powerpoint:
                    return CreatePresentationFile(name);

                case OfficeApp.Notepad:
                    return CreateFile(content, ".txt", name);

                case OfficeApp.Unsupported:
                    return CreateFile(content, ".unsupported", name);

                default:
                    Console.WriteLine(@"Provided Office application format not yet supported.");
                    return null;
            }
        }

        public static void SaveScreenShotsAndLogs(AddInHost app, string testName)
        {
            try
            {
                app.Oc.CaptureScreen(testName);
                app.Oc.SaveLogs(testName);
            }
            catch (Exception e)
            {
                Console.WriteLine($@"Failed to save screenshot/logs. Reason: {e.Message}");
            }
        }

        public static string FormatDate(DateTime date)
        {
            return date.ToString(DateFormat, usCulture);
        }

        public static string FormatDateTime(DateTime date)
        {
            return date.ToString(DateTimeFormat, usCulture).ToLower();
        }

        public static string FormatDateTime(string dateTime)
        {
            return FormatDateTime(Convert.ToDateTime(dateTime, usCulture));
        }

        public static string FormatDateRange(DateTime start, DateTime end)
        {
            return $"{start.ToString(DateFormat)} - {end.ToString(DateFormat)}";
        }

        public static string GetLongDateString()
        {
            return DateTime.Now.ToString(DateTimeDashFormat);
        }

        private static FileInfo CreateFile(string content, string extension, string name)
        {
            var tempFile = GetNewFilePath(extension, name);

            File.WriteAllText(tempFile, content);
            return new FileInfo(tempFile);
        }

        private static FileInfo CreateWordFile(string content, string name)
        {
            var tempFile = GetNewFilePath(".docx", name);

            using (var docx = WordprocessingDocument.Create(tempFile, WordprocessingDocumentType.Document))
            {
                var mainPart = docx.AddMainDocumentPart();
                mainPart.Document = new W.Document();
                var body = mainPart.Document.AppendChild(new W.Body());
                var paragraph = body.AppendChild(new W.Paragraph());
                var run = paragraph.AppendChild(new W.Run());
                run.AppendChild(new W.Text(content));
            }

            return new FileInfo(tempFile);
        }

        private static FileInfo CreateExcelFile(string content, string name, uint rowIndex = 1U, string cellRef = "A1")
        {
            var tempFile = GetNewFilePath(".xlsx", name);

            using (var xlsx = SpreadsheetDocument.Create(tempFile, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = xlsx.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                var sheets = xlsx.WorkbookPart.Workbook.AppendChild(new Sheets());

                var sheet = new Sheet()
                {
                    Id = xlsx.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Test Sheet"
                };

                sheets.AppendChild(sheet);
                var worksheet = new Worksheet();
                var sheetData = new SheetData();

                var row = new Row() { RowIndex = rowIndex, Spans = new ListValue<StringValue>() };
                var cell = new Cell()
                {
                    CellReference = cellRef,
                    DataType = CellValues.String,
                    CellValue = new CellValue(content)
                };

                row.AppendChild(cell);
                sheetData.AppendChild(row);
                worksheet.AppendChild(sheetData);
                worksheetPart.Worksheet = worksheet;
                workbookPart.Workbook.Save();
            }
            return new FileInfo(tempFile);
        }

        private static FileInfo CreatePresentationFile(string name)
        {
            var tempFile = GetNewFilePath(".pptx", name);

            using (var pptx = PresentationDocument.Create(tempFile, PresentationDocumentType.Presentation))
            {
                var presentationPart = pptx.AddPresentationPart();
                presentationPart.Presentation = new Presentation();

                PresentationHelper.CreatePresentationParts(presentationPart);
            }

            return new FileInfo(tempFile);
        }

        private static string GetNewFilePath(string extension, string name)
        {
            var path = Windows.GetWorkingTempFolder();
            var newFilePath = Path.Combine(path.FullName, Path.GetFileNameWithoutExtension(
                string.IsNullOrEmpty(name) ? Path.GetRandomFileName() : name));
            newFilePath += extension;
            return newFilePath;
        }

        public static double ConvertBytesToKb(double length, int numberOfDecimal)
        {
            return Math.Round(length / 1024, numberOfDecimal);
        }

        public static string AddSpacesToTextAtCamelCase(string name)
        {
            return Regex.Replace(name, "(\\B[A-Z])", " $1");
        }
    }
}
