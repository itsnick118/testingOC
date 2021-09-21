using System;
using System.Windows.Automation;
using IntegratedDriver;
using IntegratedDriver.ElementFinders;
using UITests.PageModel.Selectors;
using UITests.PageModel.Shared;
using Application = Microsoft.Office.Interop.Word.Application;

namespace UITests.PageModel.OfficeApps
{
    public class Word : OfficeApplication
    {
        public Word(TestEnvironment testEnvironment) : base(testEnvironment)
        {
        }

        public string ReadWordContent(string path)
        {
            var app = new Application();
            var doc = app.Documents.Open(path);
            var content = doc.Range(doc.Content.Start, doc.Content.End - 1);
            var allWords = content.Text;
            doc.Close();
            app.Quit();

            return allWords;
        }

        public void ReplaceTextWith(string content)
        {
            SetForegroundWindow();
            AppWindow.SetFocus();
            Wait();

            UserInput.SelectAll();
            UserInput.Type(content);
            Wait();
        }

        public string ReadActiveFileContent()
        {
            Wait();
            SetForegroundWindow();
            Wait();
            object wordAsObject;
            Application word;

            try
            {
                Wait();
                wordAsObject = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
                word = (Application)wordAsObject;
                var activeDocument = word.ActiveDocument;
                return activeDocument.Content.Text.Replace("\r", "");
            }
            catch (Exception ex)
            {
                throw new Exception($"Unable to read file : {ex}");
            }
        }

        public string GetActiveFilePath()
        {
            Wait();
            object wordAsObject;
            Application word;
            try
            {
                Wait();
                wordAsObject = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
                word = (Application)wordAsObject;
                var activeDocument = word.ActiveDocument;
                return activeDocument.FullName;
            }
            catch (Exception ex)
            {
                throw new Exception($"Unable to get file Path : {ex}");
            }
        }

        public void NewWorkBook()
        {
            UserInput.KeyPress("{ENTER}");
        }

        public void OpenNewWord()
        {
            var fileButton = NativeFinder.Find(AppWindow, Native.FileButton, ControlType.Button, 2);
            UserInput.LeftClick(fileButton);

            var word =
                NativeFinder.Find(AppWindow, Native.NewWordTitle, ControlType.ListItem, 2);
            UserInput.LeftClick(word);
        }
    }
}
