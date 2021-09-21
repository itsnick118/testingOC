using IntegratedDriver;
using IntegratedDriver.ElementFinders;
using System.Windows.Automation;
using UITests.PageModel.Selectors;

namespace UITests.PageModel
{
    public class OutlookContacts
    {
        private readonly AutomationElement _outlookWindow;

        public OutlookContacts(AutomationElement outlookWindow)
        {
            _outlookWindow = outlookWindow;
        }

        public void DeleteContact(string personName)
        {
            try
            {
                var contactList = NativeFinder.Find(_outlookWindow, Native.ContactList, ControlType.Group, 3);
                var contact = NativeFinder.FindByPartialMatch(contactList, personName, ControlType.Button, 3);
                while (contact != null)
                {
                    SelectContact(personName);
                    DeleteSelectedContact();
                    contact = NativeFinder.FindByPartialMatch(contactList, personName, ControlType.Button, 3);
                }
            }
            catch (ElementNotAvailableException)
            {
            }
        }

        public void DeleteSelectedContact()
        {
            var deleteContact = NativeFinder.Find(_outlookWindow, Native.DeleteButton, ControlType.Button);
            UserInput.LeftClick(deleteContact);
        }

        public string GetContactDetailsFromEditBox(string value)
        {
            var personDetails = NativeFinder.Find(_outlookWindow, value, ControlType.Text);
            return GetValueFromAutomationElement(personDetails);
        }

        public string GetPersonFullNameFromContactCard(string windowName)
        {
            var contactCardWindow = Windows.GetWindowWithName(windowName, false);
            var fullName = NativeFinder.Find(contactCardWindow, Native.FullName, ControlType.Edit);
            return GetValueFromAutomationElement(fullName);
        }

        public string GetPersonJobTitleFromContactCard(string windowName)
        {
            var contactCardWindow = Windows.GetWindowWithName(windowName, false);
            var jobTitle = NativeFinder.Find(contactCardWindow, Native.JobTitle, ControlType.Edit);
            return GetValueFromAutomationElement(jobTitle);
        }

        public int GetSavedContactCount()
        {
            var contactList = NativeFinder.Find(_outlookWindow, Native.ContactList, ControlType.Group);
            var totalContactList = NativeFinder.FindAll(contactList, ControlType.Button);
            return totalContactList == null ? 0 : totalContactList.Count;
        }

        public void Open()
        {
            var peopleButton = NativeFinder.Find(_outlookWindow, Native.PeopleTab, ControlType.Button);
            UserInput.LeftClick(peopleButton);
        }

        public void SaveContact(string windowName)
        {
            var contactCardWindow = Windows.GetWindowWithName(windowName, false);
            var saveAndClose = NativeFinder.Find(contactCardWindow, Native.SaveAndClose, ControlType.Button);
            UserInput.LeftClick(saveAndClose);
        }

        public void SelectContact(string personName)
        {
            var contact = NativeFinder.FindByPartialMatch(_outlookWindow, personName, ControlType.Button, 3);
            UserInput.LeftClick(contact);
        }

        private string GetValueFromAutomationElement(AutomationElement element)
        {
            object patternObj;
            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out patternObj))
            {
                var valuePattern = (ValuePattern)patternObj;
                return valuePattern.Current.Value;
            }
            else if (element.TryGetCurrentPattern(TextPattern.Pattern, out patternObj))
            {
                var textPattern = (TextPattern)patternObj;
                return textPattern.DocumentRange.GetText(-1).TrimEnd('\r'); // often there is an extra '\r' hanging off the end.
            }
            else
            {
                return element.Current.Name;
            }
        }
    }
}