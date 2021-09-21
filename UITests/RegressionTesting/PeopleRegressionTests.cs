using System;
using System.Linq;
using NUnit.Framework;
using UITests.PageModel;
using static UITests.Constants;
using static UITests.TestHelpers;

namespace UITests.RegressionTesting
{
    [TestFixture]
    public class PeopleRegressionTests : UITestBase
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
        [Description("Test case reference: TC_16351 Verify stylings and features for People list and View under people tab")]
        public void StylingAndFeatureOfPeopleListAndViewUnderPeopleTab()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var peopleListPage = _outlook.Oc.PeopleListPage;
            var peopleList = peopleListPage.ItemList;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenRandom();

            //clean the people list from matter except PIC.
            var personPIC = peopleListPage.RemoveAllPersonsExceptPIC();

            //validate to check PIC should not be removed.
            personPIC.Remove().Confirm();
            var messages = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(1, messages.Length);
            Assert.AreEqual(MessageonTryingToRemovePIC(personPIC.PersonName), messages[0]);
            _outlook.Oc.CloseAllToastMessages();
            personPIC = peopleList.GetPeopleListItemByIndex(0);
            Assert.IsNotNull(personPIC);

            //Verify the view person in list page
            peopleList.OpenAddDialog();
            var addPersonDialog = peopleListPage.AddPersonDialog;
            var personType = "Internal";
            addPersonDialog.Controls["Person Type"].Set(personType);
            addPersonDialog.Controls["Person"].SetValueOtherthan(personPIC.PersonName);
            addPersonDialog.Controls["Role/Involvement Type"].SetByIndex(1);
            addPersonDialog.Controls["Start Date"].Set(FormatDate(DateTime.Now.AddDays(0)));
            addPersonDialog.Controls["End Date"].Set(FormatDate(DateTime.Now.AddDays(60)));
            addPersonDialog.Controls["Comments"].Set("comments_test");

            var personName = addPersonDialog.Controls["Person"].GetValue();
            var roleInvolvementType = addPersonDialog.Controls["Role/Involvement Type"].GetValue();
            var startDate = addPersonDialog.Controls["Start Date"].GetValue();
            var endDate = addPersonDialog.Controls["End Date"].GetValue();
            var comment = addPersonDialog.Controls["Comments"].GetValue();
            var isActive = addPersonDialog.Controls["Active"].GetValue();
            addPersonDialog.Save();

            var addedPerson = peopleList.GetPeopleListItemFromText(personName);

            Assert.IsNotNull(addedPerson, "Newly added person is not listed after saving");
            Assert.AreEqual(true, peopleList.IsSortIconVisible);
            Assert.AreEqual(personName, addedPerson.PersonName);
            Assert.AreEqual(roleInvolvementType, addedPerson.Role);
            Assert.Contains(addedPerson.TernaryElement.Text.Split('-')[0].Trim(), new[] { "Internal", "External" });
            Assert.AreEqual(true, addedPerson.IsRemovePersonButtonVisible());
            Assert.AreEqual(true, addedPerson.IsEditButtonVisible());
            Assert.AreEqual(true, addedPerson.IsContactIconVisible());
            Assert.AreEqual(true, addedPerson.IsEmailIconVisible());
            Assert.AreEqual(BlackColorName, addedPerson.GetPersonNameColor().Name);

            //Verify the view person in view mode
            addedPerson = peopleList.GetPeopleListItemFromText(personName);
            addedPerson.Open();
            Assert.IsTrue(addPersonDialog.IsDisplayed());
            Assert.AreEqual(personType, addPersonDialog.Controls["Person Type"].GetReadOnlyValue());
            Assert.AreEqual(personName, addPersonDialog.Controls["Person"].GetReadOnlyValue());
            Assert.AreEqual(roleInvolvementType, addPersonDialog.Controls["Role/Involvement Type"].GetReadOnlyValue());
            Assert.AreEqual(startDate, addPersonDialog.Controls["Start Date"].GetReadOnlyValue());
            Assert.AreEqual(endDate, addPersonDialog.Controls["End Date"].GetReadOnlyValue());
            Assert.AreEqual(comment, addPersonDialog.Controls["Comments"].GetReadOnlyValue());
            Assert.AreEqual(isActive, addPersonDialog.Controls["Active"].GetReadOnlyValue());
            addPersonDialog.Cancel();
            Assert.IsFalse(addPersonDialog.IsDisplayed());
            addedPerson.Remove().Confirm();
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_16387 Verify the View People list page works as expected when View, Edit and Delete")]
        public void ViewPeopleListWorksExpectedOnViewEditAndDelete()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var peopleListPage = _outlook.Oc.PeopleListPage;
            var peopleList = peopleListPage.ItemList;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenRandom();

            //clean the people list from matter except PIC.
            var personPIC = peopleListPage.RemoveAllPersonsExceptPIC();

            //validate to check PIC should not be removed.
            personPIC.Remove().Confirm();
            var messages = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(1, messages.Length);
            Assert.AreEqual(MessageonTryingToRemovePIC(personPIC.PersonName), messages[0]);
            _outlook.Oc.CloseAllToastMessages();
            personPIC = peopleList.GetPeopleListItemByIndex(0);
            Assert.IsNotNull(personPIC);

            //validate view person works as expected.
            var personType = "Internal";
            peopleList.OpenAddDialog();
            var addPersonDialog = peopleListPage.AddPersonDialog;
            addPersonDialog.Controls["Person Type"].Set(personType);
            addPersonDialog.Controls["Person"].SetValueOtherthan(personPIC.PersonName);
            addPersonDialog.Controls["Role/Involvement Type"].SetByIndex(1);
            addPersonDialog.Controls["Start Date"].Set(FormatDate(DateTime.Now.AddDays(0)));
            addPersonDialog.Controls["End Date"].Set(FormatDate(DateTime.Now.AddDays(60)));
            addPersonDialog.Controls["Comments"].Set("comments_test");

            var personName = addPersonDialog.Controls["Person"].GetValue();
            var roleInvolvementType = addPersonDialog.Controls["Role/Involvement Type"].GetValue();
            var startDate = addPersonDialog.Controls["Start Date"].GetValue();
            var endDate = addPersonDialog.Controls["End Date"].GetValue();
            var comment = addPersonDialog.Controls["Comments"].GetValue();
            var isActive = addPersonDialog.Controls["Active"].GetValue();
            addPersonDialog.Save();

            var addedPerson = peopleList.GetPeopleListItemFromText(personName);
            addedPerson.Open();
            Assert.IsTrue(addPersonDialog.IsDisplayed());
            Assert.AreEqual(new[] { "Edit", "Cancel" }, addPersonDialog.GetDialogButtons());
            Assert.AreEqual(personType, addPersonDialog.Controls["Person Type"].GetReadOnlyValue());
            Assert.AreEqual(personName, addPersonDialog.Controls["Person"].GetReadOnlyValue());
            Assert.AreEqual(roleInvolvementType, addPersonDialog.Controls["Role/Involvement Type"].GetReadOnlyValue());
            Assert.AreEqual(startDate, addPersonDialog.Controls["Start Date"].GetReadOnlyValue());
            Assert.AreEqual(endDate, addPersonDialog.Controls["End Date"].GetReadOnlyValue());
            Assert.AreEqual(comment, addPersonDialog.Controls["Comments"].GetReadOnlyValue());
            Assert.AreEqual(isActive, addPersonDialog.Controls["Active"].GetReadOnlyValue());
            addPersonDialog.Cancel();
            Assert.IsFalse(addPersonDialog.IsDisplayed());

            //verify edit people work as expected.
            addedPerson.Edit();
            Assert.IsTrue(addPersonDialog.IsDisplayed());
            Assert.AreEqual(personType, addPersonDialog.Controls["Person Type"].GetValue());
            Assert.AreEqual(personName, addPersonDialog.Controls["Person"].GetValue());
            Assert.AreEqual(roleInvolvementType, addPersonDialog.Controls["Role/Involvement Type"].GetValue());
            Assert.AreEqual(startDate, addPersonDialog.Controls["Start Date"].GetValue());
            Assert.AreEqual(endDate, addPersonDialog.Controls["End Date"].GetValue());
            Assert.AreEqual(comment, addPersonDialog.Controls["Comments"].GetValue());
            Assert.AreEqual(editPopupButtons, addPersonDialog.GetDialogButtons());

            //validate reset works expected.
            addPersonDialog.Controls["Person"].SetValueOtherthan(personPIC.PersonName, true);
            addPersonDialog.Reset();
            Assert.AreEqual(personName, addPersonDialog.Controls["Person"].GetReadOnlyValue());
            addPersonDialog.Save();
            addedPerson = peopleList.GetPeopleListItemFromText(personName);
            Assert.AreEqual(personName, addedPerson.PersonName);
            Assert.IsNotNull(addedPerson);
            addedPerson.Remove().Confirm();

            //validate internal person(other than PIC) should get removed
            peopleList.OpenAddDialog();
            addPersonDialog = peopleListPage.AddPersonDialog;
            addPersonDialog.Controls["Person Type"].Set(personType);
            addPersonDialog.Controls["Person"].SetValueOtherthan(personPIC.PersonName);
            addPersonDialog.Controls["Role/Involvement Type"].SetByIndex(1);
            addPersonDialog.Controls["Start Date"].Set(FormatDate(DateTime.Now.AddDays(0)));
            addPersonDialog.Controls["End Date"].Set(FormatDate(DateTime.Now.AddDays(60)));
            addPersonDialog.Controls["Comments"].Set("comments_test");
            personName = addPersonDialog.Controls["Person"].GetValue();
            addPersonDialog.Save();

            addedPerson = peopleList.GetPeopleListItemFromText(personName);
            Assert.IsNotNull(addedPerson, "Newly added person is not listed after saving");
            addedPerson.Remove().Confirm();
            addedPerson = peopleList.GetPeopleListItemFromText(personName);
            Assert.IsNull(addedPerson, "Removed person is still visible in the list");

            //validate external person should get removed
            personType = "External - Other";
            peopleList.OpenAddDialog();
            addPersonDialog = peopleListPage.AddPersonDialog;
            addPersonDialog.Controls["Person Type"].Set(personType);
            addPersonDialog.Controls["Person"].SetValueOtherthan(personPIC.PersonName);
            addPersonDialog.Controls["Role/Involvement Type"].SetByIndex(1);
            addPersonDialog.Controls["Start Date"].Set(FormatDate(DateTime.Now.AddDays(0)));
            addPersonDialog.Controls["End Date"].Set(FormatDate(DateTime.Now.AddDays(60)));
            addPersonDialog.Controls["Comments"].Set("comments_test");
            personName = addPersonDialog.Controls["Person"].GetValue();
            addPersonDialog.Save();

            addedPerson = peopleList.GetPeopleListItemFromText(personName);
            Assert.IsNotNull(addedPerson, "Newly added person is not listed after saving");
            addedPerson.Remove().Confirm();
            addedPerson = peopleList.GetPeopleListItemFromText(personName);
            Assert.IsNull(addedPerson, "Removed person is still visible in the list");
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("TC_16374 : Verify to generate an email for person from person list")]
        public void GenerateEmailForPerson()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var peopleListPage = _outlook.Oc.PeopleListPage;
            var matterDetailPage = _outlook.Oc.MatterDetailsPage;
            var peopleList = peopleListPage.ItemList;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenRandom();
            var matterName = matterDetailPage.MatterName;

            //verify email icon displays primary email address
            foreach (var person in peopleList.GetAllPeopleListItems().Where(x => x.Role == "Primary Internal Contact"))
            {
                var primaryEmailAddress = person.GetEmailAddress();
                StringAssert.IsMatch(@"^[^\s@]+@[^\s@]+\.[^\s@]+$", primaryEmailAddress);
                person.OpenEmailWindow();
                _outlookEmailForm = new OutlookEmailForm(TestEnvironment);
                _outlookEmailForm.Attach(matterName);
                Assert.AreEqual(primaryEmailAddress, _outlookEmailForm.GetEmailRecipientTo());
                Assert.IsEmpty(_outlookEmailForm.GetEmailFormValue("Cc"));
                Assert.AreEqual(matterName, _outlookEmailForm.GetEmailFormValue("Subject"));
                _outlookEmailForm.SaveDocument();
                _outlookEmailForm.CloseDocument();
            }

            peopleList.OpenAddDialog();
            var addPersonDialog = _outlook.Oc.PeopleListPage.AddPersonDialog;

            addPersonDialog.Controls["Person Type"].Set("External - Other");
            var selectedPerson = addPersonDialog.Controls["Person"].Set("Timekeeper");
            addPersonDialog.Controls["Role/Involvement Type"].SetByIndex(2);
            addPersonDialog.Save();

            var createdPerson = peopleList.GetPeopleListItemFromText(selectedPerson);
            Assert.IsNotNull(createdPerson);
            Assert.AreEqual("No Primary Mail Address", createdPerson.GetEmailAddress());

            //clean up
            createdPerson.Remove().Confirm();
        }

        [Test]
        [Category(RegressionTestCategory)]
        [TestCase("Internal")]
        [TestCase("External - Associated Organization")]
        [TestCase("External - Other")]
        [Description("Test case reference: TC_16376 Verify Add/Edit People to the matter")]
        public void AddAndEditPersonToMatter(string personType)
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var peopleListPage = _outlook.Oc.PeopleListPage;
            var peopleList = peopleListPage.ItemList;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();

            //clean the people list from matter except PIC.
            var personPIC = peopleListPage.RemoveAllPersonsExceptPIC();

            //start date is in past validation
            //TODO : Open Person and RIT autocomplete and verify the records
            /*1. Open(Click search Icon) Person Autocomplete and validate all persons are Internal/External.
              2. Open(Click search Icon) RIT AutoComplete and validate all RIT.*/

            peopleList.OpenAddDialog();
            var addPersonDialog = peopleListPage.AddPersonDialog;
            addPersonDialog.Controls["Person Type"].Set(personType);
            var person = addPersonDialog.Controls["Person"].SetValueOtherthan(personPIC.PersonName);
            if (person == null)
            {
                Assert.Fail($"Person list is empty for {personType}");
            }
            addPersonDialog.Controls["Role/Involvement Type"].SetByIndex(1);
            addPersonDialog.Controls["Start Date"].Set(FormatDate(DateTime.Now.AddDays(-60)));
            addPersonDialog.Controls["End Date"].Set(FormatDate(DateTime.Now.AddDays(60)));
            addPersonDialog.Controls["Comments"].Set("comments_test");

            var personName = addPersonDialog.Controls["Person"].GetValue();
            var roleInvolvementType = addPersonDialog.Controls["Role/Involvement Type"].GetValue();
            addPersonDialog.Save();
            var addedPerson = peopleList.GetPeopleListItemFromText(personName);

            Assert.AreEqual(personName, addedPerson.PersonName);
            Assert.AreEqual(roleInvolvementType, addedPerson.Role);
            addedPerson.Remove().Confirm();

            // add/edit/reset/discard(discard cancel).
            peopleList.OpenAddDialog();
            addPersonDialog = peopleListPage.AddPersonDialog;
            Assert.AreEqual(addPersonDialog.GetDialogButtons(), new[] { "Save", "Cancel" });
            Assert.AreEqual(addPersonDialog.HeaderText, "Add Person");
            addPersonDialog.Controls["Person Type"].Set(personType);
            addPersonDialog.Controls["Person"].SetValueOtherthan(personPIC.PersonName);
            addPersonDialog.Controls["Role/Involvement Type"].SetByIndex(3);
            addPersonDialog.Controls["Start Date"].Set(FormatDate(DateTime.Now));
            addPersonDialog.Controls["End Date"].Set(FormatDate(DateTime.Now.AddDays(60)));
            addPersonDialog.Controls["Comments"].Set("comments_test");

            personName = addPersonDialog.Controls["Person"].GetValue();
            roleInvolvementType = addPersonDialog.Controls["Role/Involvement Type"].GetValue();

            //discard changes and do not discard.
            addPersonDialog.Cancel(false);
            Assert.AreEqual(CancelMessage, addPersonDialog.Text);
            addPersonDialog.DoNotDiscard();

            addPersonDialog.Save();
            addedPerson = peopleList.GetPeopleListItemFromText(personName);
            Assert.AreEqual(personName, addedPerson.PersonName);
            Assert.AreEqual(roleInvolvementType, addedPerson.Role);

            //edit person.
            addedPerson.Edit();
            Assert.AreEqual(addPersonDialog.GetDialogButtons(), editPopupButtons);
            Assert.AreEqual(addPersonDialog.HeaderText, "Edit Person");
            addPersonDialog.Controls["Person"].SetValueOtherthan(personPIC.PersonName, true);
            addPersonDialog.Controls["Role/Involvement Type"].SetByIndex(2, true);
            personName = addPersonDialog.Controls["Person"].GetValue();
            roleInvolvementType = addPersonDialog.Controls["Role/Involvement Type"].GetValue();

            addPersonDialog.Save();
            addedPerson = peopleList.GetPeopleListItemFromText(personName);
            Assert.AreEqual(personName, addedPerson.PersonName);
            Assert.AreEqual(roleInvolvementType, addedPerson.Role);

            //reset person details.
            addedPerson.Edit();
            var comment = addPersonDialog.Controls["Comments"].GetValue();
            personName = addPersonDialog.Controls["Person"].GetValue();
            roleInvolvementType = addPersonDialog.Controls["Role/Involvement Type"].GetValue();
            addPersonDialog.Controls["Person"].SetValueOtherthan(personPIC.PersonName, true);
            addPersonDialog.Controls["Role/Involvement Type"].SetByIndex(3, true);
            addPersonDialog.Controls["Comments"].Set("new comments_test");
            addPersonDialog.Reset();
            Assert.AreEqual(personName, addPersonDialog.Controls["Person"].GetValue());
            Assert.AreEqual(roleInvolvementType, addPersonDialog.Controls["Role/Involvement Type"].GetValue());
            Assert.AreEqual(comment, addPersonDialog.Controls["Comments"].GetValue());
            addPersonDialog.Save();
            addedPerson = peopleList.GetPeopleListItemFromText(personName);
            Assert.AreEqual(personName, addedPerson.PersonName);
            Assert.AreEqual(roleInvolvementType, addedPerson.Role);

            //cleanup
            addedPerson.Remove().Confirm();

            //discard changes and discard all.
            peopleList.OpenAddDialog();
            addPersonDialog = peopleListPage.AddPersonDialog;
            addPersonDialog.Controls["Person Type"].Set(personType);
            addPersonDialog.Controls["Person"].SetValueOtherthan(personPIC.PersonName);
            addPersonDialog.Controls["Role/Involvement Type"].SetByIndex(1);
            addPersonDialog.Controls["Start Date"].Set(FormatDate(DateTime.Now));
            addPersonDialog.Controls["End Date"].Set(FormatDate(DateTime.Now.AddDays(60)));
            addPersonDialog.Controls["Comments"].Set("comments_test");
            personName = addPersonDialog.Controls["Person"].GetValue();

            addPersonDialog.Cancel(false);
            Assert.AreEqual(CancelMessage, addPersonDialog.Text);
            addPersonDialog.DiscardChanges();
            addedPerson = peopleList.GetPeopleListItemFromText(personName);
            Assert.IsNull(addedPerson);

            //start date is in future validation
            peopleList.OpenAddDialog();
            addPersonDialog = peopleListPage.AddPersonDialog;
            addPersonDialog.Controls["Person Type"].Set(personType);
            addPersonDialog.Controls["Person"].SetValueOtherthan(personPIC.PersonName);
            addPersonDialog.Controls["Role/Involvement Type"].SetByIndex(1);
            addPersonDialog.Controls["Start Date"].Set(FormatDate(DateTime.Now.AddDays(60)));
            addPersonDialog.Controls["Comments"].Set("comments_test");
            personName = addPersonDialog.Controls["Person"].GetValue();

            addPersonDialog.Save();
            var messages = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(1, messages.Length);
            Assert.AreEqual(FutureStartDateMessage, messages[0]);
            _outlook.Oc.CloseAllToastMessages();
            addedPerson = peopleList.GetPeopleListItemFromText(personName);
            Assert.IsNull(addedPerson);

            //add same person with same start/end date.
            for (var i = 0; i < 2; i++)
            {
                peopleList.OpenAddDialog();
                addPersonDialog = peopleListPage.AddPersonDialog;
                addPersonDialog.Controls["Person Type"].Set(personType);
                addPersonDialog.Controls["Person"].SetByIndex(3);
                addPersonDialog.Controls["Role/Involvement Type"].SetByIndex(1);
                addPersonDialog.Controls["Start Date"].Set(FormatDate(DateTime.Now));
                addPersonDialog.Controls["End Date"].Set(FormatDate(DateTime.Now.AddDays(60)));
                addPersonDialog.Controls["Comments"].Set("comments_test");

                personName = addPersonDialog.Controls["Person"].GetValue();
                roleInvolvementType = addPersonDialog.Controls["Role/Involvement Type"].GetValue();

                //if person is from PIC then select another
                if (personName.Equals(personPIC.PersonName))
                {
                    addPersonDialog.Controls["Person"].SetByIndex(2);
                    personName = addPersonDialog.Controls["Person"].GetValue();
                }

                addPersonDialog.Save();
                messages = _outlook.Oc.GetAllToastMessages();
                _outlook.Oc.CloseAllToastMessages();
                addedPerson = peopleList.GetPeopleListItemFromText(personName);
                Assert.AreEqual(personName, addedPerson.PersonName);
                Assert.AreEqual(roleInvolvementType, addedPerson.Role);
                if (i == 1)
                {
                    Assert.AreEqual(1, messages.Length);
                    Assert.AreEqual(OverlappingTimePeriodMessage, messages[0]);
                    addedPerson.Remove().Confirm();
                }
            }
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("TC_16375 : Verify to Generate the Contact info card when hitting the Contact info icon and saving it locally to outlook")]
        public void GenerateContactInfoCardAndSaveLocallyInOutlook()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var peopleListPage = _outlook.Oc.PeopleListPage;
            var peopleList = peopleListPage.ItemList;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenRandom();

            //clean the people list from matter except PIC.
            var personPIC = peopleListPage.RemoveAllPersonsExceptPIC();

            //Open Contact Folder
            _outlook.Contacts.Open();

            //Delete existing contacts
            _outlook.Contacts.DeleteContact(personPIC.PersonName);

            var contactCount = _outlook.Contacts.GetSavedContactCount();
            personPIC.ViewContact();

            var windowName = $"{personPIC.PersonName} - Contact";
            var personName = _outlook.Contacts.GetPersonFullNameFromContactCard(windowName);
            var jobTitle = _outlook.Contacts.GetPersonJobTitleFromContactCard(windowName);

            Assert.AreEqual(personPIC.PersonName, personName);
            Assert.AreEqual(personPIC.Role, jobTitle.Split('/')[0]);
            Assert.AreEqual(personPIC.PersonType, jobTitle.Split('/')[1]);

            _outlook.Contacts.SaveContact(windowName);
            _outlook.Contacts.SelectContact(personPIC.PersonName);

            Assert.AreEqual(personName, _outlook.Contacts.GetContactDetailsFromEditBox(personName));
            Assert.AreEqual(jobTitle, _outlook.Contacts.GetContactDetailsFromEditBox(jobTitle));

            //Remove added contact
            _outlook.Contacts.DeleteSelectedContact();
            Assert.AreEqual(contactCount, _outlook.Contacts.GetSavedContactCount());

            //add external person
            peopleList.OpenAddDialog();
            var addPersonDialog = peopleListPage.AddPersonDialog;
            addPersonDialog.Controls["Person Type"].Set("External - Other");
            addPersonDialog.Controls["Person"].SetValueOtherthan(personPIC.PersonName);
            addPersonDialog.Controls["Role/Involvement Type"].SetByIndex(1);
            addPersonDialog.Controls["Start Date"].Set(FormatDate(DateTime.Now.AddDays(0)));
            addPersonDialog.Controls["End Date"].Set(FormatDate(DateTime.Now.AddDays(60)));
            addPersonDialog.Controls["Comments"].Set("comments_test");

            var addedPersonName = addPersonDialog.Controls["Person"].GetValue();

            addPersonDialog.Save();
            var addedPerson = peopleList.GetPeopleListItemFromText(addedPersonName);

            _outlook.Contacts.DeleteContact(addedPerson.PersonName);
            contactCount = _outlook.Contacts.GetSavedContactCount();

            addedPerson.ViewContact();

            windowName = $"{addedPerson.PersonName} - Contact";
            personName = _outlook.Contacts.GetPersonFullNameFromContactCard(windowName);
            jobTitle = _outlook.Contacts.GetPersonJobTitleFromContactCard(windowName);

            Assert.AreEqual(addedPerson.PersonName, personName);
            Assert.AreEqual(addedPerson.Role, jobTitle.Split('/')[0]);
            Assert.AreEqual(addedPerson.PersonType, jobTitle.Split('/')[1]);

            _outlook.Contacts.SaveContact(windowName);
            _outlook.Contacts.SelectContact(addedPerson.PersonName);

            Assert.AreEqual(personName, _outlook.Contacts.GetContactDetailsFromEditBox(personName));
            Assert.AreEqual(jobTitle, _outlook.Contacts.GetContactDetailsFromEditBox(jobTitle));

            //Remove added contact
            _outlook.Contacts.DeleteSelectedContact();
            Assert.AreEqual(contactCount, _outlook.Contacts.GetSavedContactCount());

            //Remove added person
            addedPerson.Remove().Confirm();
        }

        [TearDown]
        public void TearDown()
        {
            SaveScreenShotsAndLogs(_outlook);
            _outlook?.Destroy();
        }
    }
}
