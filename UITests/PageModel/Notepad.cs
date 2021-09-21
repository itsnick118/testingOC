using System.IO;
using IntegratedDriver;
using IntegratedDriver.ElementFinders;
using System.Threading;
using System.Windows.Automation;

namespace UITests.PageModel
{
    public class Notepad
    {
        private readonly string _title;

        public Notepad(string documentName)
        {
            _title = Path.GetFileNameWithoutExtension(documentName);
        }

        public void ReplaceTextWith(string content)
        {
            var notepad = Windows.GetWindowWithName(_title, false);
            notepad.SetFocus();
            Wait();

            var notepadTextArea = NativeFinder.Find(notepad, ControlType.Document);
            notepadTextArea.SetFocus();
            Wait();

            UserInput.SelectAll();
            UserInput.Type(content);
            UserInput.Type("^s");
        }

        public void Close() => Windows.CloseWindowByName(_title);

        private static void Wait() => Thread.Sleep(100);
    }
}
