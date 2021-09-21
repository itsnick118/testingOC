using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using UITests.PageModel;
using UITests.PageModel.Shared;
using static UITests.Constants;
using static UITests.TestHelpers;

namespace UITests.RegressionTesting
{
    [TestFixture]
    public class NarrativesRegressionTests : UITestBase
    {
        private Outlook _outlook;

        [SetUp]
        public void SetUp()
        {
            _outlook = new Outlook(TestEnvironment);
            _outlook.Launch();
            _outlook.Oc.BasicSettingsPage.LogInAsStandardUser();
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_16516 Verify to View/Search Narratives List")]
        public void ViewSearchNarrativesList()
        {
            const string ToolTipEllipsis = "ellipsis";

            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var narrativesListPage = _outlook.Oc.NarrativesListPage;
            var narrativesList = narrativesListPage.ItemList;

            mattersListPage.Open();
            mattersList.OpenFirst();
            matterDetails.Tabs.Open("Narratives");

            // create test narrative for hover tests
            var bigDescription = GetRandomTextWithSpaces(255);
            var bigNarrative = GetRandomTextWithSpaces(1000);
            var type = GetRandomFrom(new[] { "Note", "Status" });

            narrativesList.OpenAddDialog();
            narrativesListPage.AddNarrativeDialog.Controls["Narrative Type"].Set(type);
            narrativesListPage.AddNarrativeDialog.Controls["Description"].Set(bigDescription);
            narrativesListPage.AddNarrativeDialog.Controls["Narrative Date"].Set(FormatDateTime(DateTime.Now));
            narrativesListPage.AddNarrativeDialog.Controls["Narrative"].Set(bigNarrative);

            narrativesListPage.AddNarrativeDialog.Save();

            var selectedNarrative = narrativesList.GetNarrativeListItemFromText(bigDescription);

            // hover over narrative description
            var descToolTipBcg = selectedNarrative.GetToolTipBackground(selectedNarrative.SecondaryElement);
            var descToolTipFont = selectedNarrative.GetToolTipFontColor(selectedNarrative.SecondaryElement);
            var descToolTipShape = selectedNarrative.GetToolTipShape(selectedNarrative.SecondaryElement);
            Assert.Multiple(() =>
            {
                Assert.AreEqual(descToolTipBcg.Name, BlackColorName, "Narrative Description tooltip background is not expected color");
                Assert.AreEqual(descToolTipFont.Name, WhiteColorName, "Narrative Description tooltip font is not expected color");
                Assert.AreEqual(descToolTipShape, ToolTipEllipsis, "Narrative Description tooltip container is not expected shape");
            });

            // hover over narrative body
            var bodyToolTipBcg = selectedNarrative.GetToolTipBackground(selectedNarrative.TernaryElement);
            var bodyToolTipFont = selectedNarrative.GetToolTipFontColor(selectedNarrative.TernaryElement);
            var bodyToolTipShape = selectedNarrative.GetToolTipShape(selectedNarrative.TernaryElement);
            Assert.Multiple(() =>
            {
                Assert.AreEqual(bodyToolTipBcg.Name, BlackColorName, "Narrative Body tooltip background is not expected color");
                Assert.AreEqual(bodyToolTipFont.Name, WhiteColorName, "Narrative Body tooltip font is not expected color");
                Assert.AreEqual(bodyToolTipShape, ToolTipEllipsis, "Narrative Body tooltip container is not expected shape");
            });

            // verify columns
            StringAssert.Contains(selectedNarrative.GetParentClass(selectedNarrative.PrimaryElement).ToLower(), "left");
            StringAssert.Contains(selectedNarrative.GetParentClass(selectedNarrative.SecondaryElement).ToLower(), "left");
            StringAssert.Contains(selectedNarrative.GetParentClass(selectedNarrative.TernaryElement).ToLower(), "left");
            StringAssert.Contains(selectedNarrative.GetParentClass(selectedNarrative.ActionItems).ToLower(), "right");
            StringAssert.Contains(selectedNarrative.GetParentClass(selectedNarrative.Meta2Element).ToLower(), "right");
            StringAssert.Contains(selectedNarrative.GetParentClass(selectedNarrative.Meta3Element).ToLower(), "right");

            // create test narratives for scroll test
            var narrativesCount = narrativesList.GetCount();

            for (var i = narrativesCount; i < 10; i++)
            {
                var smallDescription = GetRandomText(GetRandomNumber(50));
                var smallNarrative = GetRandomText(GetRandomNumber(50));

                narrativesList.OpenAddDialog();
                narrativesListPage.AddNarrativeDialog.Controls["Narrative Type"].Set(type);
                narrativesListPage.AddNarrativeDialog.Controls["Description"].Set(smallDescription);
                narrativesListPage.AddNarrativeDialog.Controls["Narrative Date"].Set(FormatDateTime(DateTime.Now));
                narrativesListPage.AddNarrativeDialog.Controls["Narrative"].Set(smallNarrative);

                narrativesListPage.AddNarrativeDialog.Save();
            }

            // verify list is scrollable
            Assert.IsFalse(narrativesList.ScrollDownIfNotAtBottom(), "List is not scrollable to bottom");

            // select random narrative
            selectedNarrative = narrativesList.GetNarrativeListItemByIndex(GetRandomNumber(narrativesCount - 1));

            // search narrative by - type, description, narrative, entered by
            narrativesListPage.QuickSearch.SearchBy(selectedNarrative.Type);
            Assert.That(narrativesList.GetAllNarrativeListItems(), Has.All.Property(nameof(NarrativeListItem.Type)).Contains(selectedNarrative.Type).IgnoreCase,
                $"Filtered list has items not containing Type - {selectedNarrative.Type}");

            narrativesListPage.QuickSearch.SearchBy(selectedNarrative.Description);
            Assert.That(narrativesList.GetAllNarrativeListItems(), Has.All.Property(nameof(NarrativeListItem.Description)).Contains(selectedNarrative.Description).IgnoreCase,
                $"Filtered list has items not containing Type - {selectedNarrative.Description}");

            narrativesListPage.QuickSearch.SearchBy(selectedNarrative.Narrative);
            Assert.That(narrativesList.GetAllNarrativeListItems(), Has.All.Property(nameof(NarrativeListItem.Narrative)).Contains(selectedNarrative.Narrative).IgnoreCase,
                $"Filtered list has items not containing Type - {selectedNarrative.Narrative}");

            narrativesListPage.QuickSearch.SearchBy(selectedNarrative.EnteredBy);
            Assert.That(narrativesList.GetAllNarrativeListItems(), Has.All.Property(nameof(NarrativeListItem.EnteredBy)).Contains(selectedNarrative.EnteredBy).IgnoreCase,
                $"Filtered list has items not containing Type - {selectedNarrative.EnteredBy}");

            narrativesListPage.QuickSearch.Close();

            // cleanup
            for (var i = narrativesList.GetCount() - 1; i >= 0; i--)
            {
                var narrative = narrativesList.GetNarrativeListItemByIndex(i);
                narrative.Delete().Confirm();
            }

            Assert.Zero(narrativesList.GetCount(), "Unable to delete narratives");
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("Test case reference: TC_16519 Verify Edit, Reset, Delete Narrative")]
        public void VerifyEditResetDeleteNarrative()
        {
            const string TypeNote = "Note";
            const string TypeStatus = "Status";

            var mattersListPage = _outlook.Oc.MattersListPage;
            var mattersList = mattersListPage.ItemList;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var narrativesListPage = _outlook.Oc.NarrativesListPage;
            var narrativesList = narrativesListPage.ItemList;

            mattersListPage.Open();
            mattersList.OpenFirst();
            matterDetails.Tabs.Open("Narratives");

            var description = GetRandomText(100);
            var narrative = GetRandomText(100);
            var date = DateTime.Now;

            var narrativesCount = narrativesList.GetCount();

            if (narrativesCount < 1)
            {
                narrativesList.OpenAddDialog();
                narrativesListPage.AddNarrativeDialog.Controls["Narrative Type"].Set(TypeNote);
                narrativesListPage.AddNarrativeDialog.Controls["Description"].Set(description);
                narrativesListPage.AddNarrativeDialog.Controls["Narrative Date"].Set(FormatDateTime(date));
                narrativesListPage.AddNarrativeDialog.Controls["Narrative"].Set(narrative);

                narrativesListPage.AddNarrativeDialog.Save();
            }

            var selectedNarrative = narrativesList.GetNarrativeListItemByIndex(0);
            Assert.IsNotNull(selectedNarrative);

            // Edit Narrative
            description = GetRandomText(50);
            narrative = GetRandomText(50);
            date = date.AddYears(1);

            selectedNarrative.Edit();

            // required sections
            Assert.Multiple(() =>
            {
                Assert.IsTrue(narrativesListPage.EditNarrativeDialog.IsDisplayed(), "Edit Narrative Dialog does not appear");
                Assert.IsTrue(narrativesListPage.EditNarrativeDialog.Controls["Narrative Type"].IsRequired(), "NarrativeType Field should be marked as required");
                Assert.IsTrue(narrativesListPage.EditNarrativeDialog.Controls["Description"].IsRequired(), "Description Field should be marked as required");
                Assert.AreEqual(narrativesListPage.EditNarrativeDialog.GetDialogButtons(), editPopupButtons, "Dialog Buttons are not displayed");
            });

            // update narrative
            narrativesListPage.EditNarrativeDialog.Controls["Narrative Type"].Set(TypeStatus);
            narrativesListPage.EditNarrativeDialog.Controls["Description"].Set(description);
            narrativesListPage.EditNarrativeDialog.Controls["Narrative Date"].Set(FormatDateTime(date));
            narrativesListPage.EditNarrativeDialog.Controls["Narrative"].Set(narrative);

            narrativesListPage.EditNarrativeDialog.Save();

            // narrative updated
            selectedNarrative = narrativesList.GetNarrativeListItemFromText(description);
            Assert.NotNull(selectedNarrative, "Narrative Description is not updated");
            Assert.AreEqual(selectedNarrative.Type, TypeStatus, "Narrative Type is not updated");
            Assert.AreEqual(FormatDateTime(Convert.ToDateTime(selectedNarrative.NarrativeDate)), FormatDateTime(date), "Narrative Date is not updated");
            Assert.AreEqual(selectedNarrative.Narrative, narrative, "Narrative Body is not updated");

            // edit narrative from edit button in narrative
            selectedNarrative.Open();
            narrativesListPage.EditNarrativeDialog.Edit();

            // clear description
            narrativesListPage.EditNarrativeDialog.Controls["Description"].Set(string.Empty);
            narrativesListPage.EditNarrativeDialog.Save(false);
            var actualRequiredText = narrativesListPage.EditNarrativeDialog.Controls["Description"].GetRequiredWarning();
            Assert.AreEqual(actualRequiredText, FieldIsRequiredWarning, "Field is required message is not correct");

            // changes in narrative fields
            narrativesListPage.EditNarrativeDialog.Controls["Narrative Type"].Set(TypeNote);
            narrativesListPage.EditNarrativeDialog.Controls["Description"].Set(GetRandomText(10));
            narrativesListPage.EditNarrativeDialog.Controls["Narrative Date"].Set(FormatDateTime(date.AddYears(1)));
            narrativesListPage.EditNarrativeDialog.Controls["Narrative"].Set(GetRandomText(10));

            // verify cancel
            narrativesListPage.EditNarrativeDialog.Cancel(false);
            var actualDialogText = narrativesListPage.EditNarrativeDialog.Text;
            Assert.AreEqual(actualDialogText, CancelMessage, "Cancel warning message is not correct.");

            narrativesListPage.EditNarrativeDialog.DiscardChanges();
            Assert.IsFalse(narrativesListPage.EditNarrativeDialog.IsDisplayed(), "Narrative Edit dialog is not closed");

            // verify reset narrative
            selectedNarrative.Edit();
            narrativesListPage.EditNarrativeDialog.Controls["Narrative Type"].Set(TypeNote);
            narrativesListPage.EditNarrativeDialog.Controls["Description"].Set(GetRandomText(10));
            narrativesListPage.EditNarrativeDialog.Controls["Narrative Date"].Set(FormatDateTime(date.AddYears(1)));
            narrativesListPage.EditNarrativeDialog.Controls["Narrative"].Set(GetRandomText(10));

            narrativesListPage.EditNarrativeDialog.Reset();
            var resetType = narrativesListPage.EditNarrativeDialog.Controls["Narrative Type"].GetValue();
            var resetDesc = narrativesListPage.EditNarrativeDialog.Controls["Description"].GetValue();
            var resetDate = narrativesListPage.EditNarrativeDialog.Controls["Narrative Date"].GetValue();
            var resetBody = narrativesListPage.EditNarrativeDialog.Controls["Narrative"].GetValue();

            Assert.AreEqual(resetType, TypeStatus);
            Assert.AreEqual(resetDesc, description);
            Assert.AreEqual(FormatDateTime(resetDate), FormatDateTime(date));
            Assert.AreEqual(resetBody, narrative);

            // edit after reset
            description = GetRandomText(10);
            narrative = GetRandomText(10);
            date = date.AddYears(2);
            narrativesListPage.EditNarrativeDialog.Controls["Narrative Type"].Set(TypeNote);
            narrativesListPage.EditNarrativeDialog.Controls["Description"].Set(description);
            narrativesListPage.EditNarrativeDialog.Controls["Narrative Date"].Set(FormatDateTime(date));
            narrativesListPage.EditNarrativeDialog.Controls["Narrative"].Set(narrative);

            narrativesListPage.EditNarrativeDialog.Save();

            // narrative updated
            selectedNarrative = narrativesList.GetNarrativeListItemFromText(description);

            Assert.NotNull(selectedNarrative, "Narrative Description is not updated");
            Assert.AreEqual(selectedNarrative.Type, TypeNote, "Narrative Type is not updated");
            Assert.AreEqual(FormatDateTime(Convert.ToDateTime(selectedNarrative.NarrativeDate)), FormatDateTime(date), "Narrative Date is not updated");
            Assert.AreEqual(selectedNarrative.Narrative, narrative, "Narrative Body is not updated");

            // delete narrative
            selectedNarrative.Delete();
            actualDialogText = narrativesListPage.EditNarrativeDialog.Text;
            Assert.AreEqual(actualDialogText, DeleteMessage, "Delete warning message is not correct.");
            narrativesListPage.EditNarrativeDialog.Cancel();
            Assert.IsNotNull(selectedNarrative);

            selectedNarrative.Delete().Confirm();
            selectedNarrative = narrativesList.GetNarrativeListItemFromText(description, false);
            Assert.IsNull(selectedNarrative);
        }

        [Test]
        [Description("Test case reference: TC 16517 Verify Narrative list")]
        public void NarrativesListCount()
        {
            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var narrativesListPage = _outlook.Oc.NarrativesListPage;
            var narrativesList = narrativesListPage.ItemList;

            var description = GetRandomText(100);

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetailsPage.Tabs.Open("Narratives");
            var narrativesCount = narrativesList.GetCount();

            if (narrativesCount < 1)
            {
                narrativesList.OpenAddDialog();
                narrativesListPage.AddNarrativeDialog.Controls["Narrative Type"].Set("Note");
                narrativesListPage.AddNarrativeDialog.Controls["Description"].Set(AutomatedComment);

                narrativesListPage.AddNarrativeDialog.Save();
            }

            //Verify Narrative Before Count display
            var beforeCount = narrativesList.GetFooterCount();

            // Add narrative
            narrativesList.OpenAddDialog();
            narrativesListPage.AddNarrativeDialog.Controls["Narrative Type"].Set("Note");
            narrativesListPage.AddNarrativeDialog.Controls["Description"].Set(description);
            narrativesListPage.AddNarrativeDialog.Save();

            var createdNote = narrativesList.GetNarrativeListItemFromText(description);
            Assert.IsNotNull(createdNote);
            Assert.IsEmpty(createdNote.Narrative);

            //Verify Narrative After Count display
            var afterCount = narrativesList.GetFooterCount();
            Assert.AreEqual(beforeCount + 1, afterCount, "Footer list count is not incremented upon adding a narrative");

            // cleanup
            createdNote.Delete().Confirm();
            createdNote = narrativesList.GetNarrativeListItemFromText(description, false);
            Assert.IsNull(createdNote);
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("TC:16522 - To verify single or multiple emails be dragged and dropped to Narrative section and new Narrative is created accordingly")]
        public void DragAndDropNarrative()
        {
            var subject = _outlook.AddTestEmailsToFolder(1, useDifferentTemplates: true).First().Key;
            _outlook.OpenTestEmailFolder();
            _outlook.SelectAllItems();

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetailsPage = _outlook.Oc.MatterDetailsPage;
            var narrativesListPage = _outlook.Oc.NarrativesListPage;
            var narrativesList = narrativesListPage.ItemList;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenRandom();
            matterDetailsPage.Tabs.Open("Narratives");

            //  Drag And Drop single email
            _outlook.DragAndDropEmailToOc(_outlook.GetNthItem(0), matterDetailsPage.DropPoint.GetElement());
            narrativesListPage.QuickSearch.SearchBy(subject);
            Assert.NotNull(narrativesList.GetNarrativeListItemFromText(subject), "List does not contain a narrative based on your search");

            // Drag And Drop same email again for duplicate message.
            _outlook.DragAndDropEmailToOc(_outlook.GetNthItem(0), matterDetailsPage.DropPoint.GetElement());
            var messages = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(1, messages.Length);
            StringAssert.Contains(NarrativeDuplicateTestMessage, messages[0]);

            _outlook.Oc.CloseAllToastMessages();

            // Validate Narrative fields has values
            var createdNarrative = narrativesList.GetNarrativeListItemFromText(subject);
            createdNarrative.Open();
            narrativesListPage.EditNarrativeDialog.Edit();

            var type = narrativesListPage.EditNarrativeDialog.Controls["Narrative Type"].GetValue();
            var descriptions = narrativesListPage.EditNarrativeDialog.Controls["Description"].GetValue();
            var date = narrativesListPage.EditNarrativeDialog.Controls["Narrative Date"].GetValue();
            var narratives = narrativesListPage.EditNarrativeDialog.Controls["Narrative"].GetValue();

            Assert.IsNotNull(type, "Type is empty after drag and drop file");
            Assert.IsNotNull(descriptions, "Description is empty after drag and drop file");
            Assert.IsNotNull(date, "Date field is empty after drag and drop file");
            Assert.IsNotNull(narratives, "Narrative is empty after drag and drop file");

            narrativesListPage.EditNarrativeDialog.Save();
            narrativesListPage.QuickSearch.Close();

            // Verify OC does not allow emails without subject
            _outlook.AddTestEmailsToFolder(1, subject: string.Empty, useDifferentTemplates: true);
            _outlook.OpenTestEmailFolder();
            _outlook.SelectAllItems();

            _outlook.DragAndDropEmailToOc(_outlook.GetNthItem(0), matterDetailsPage.DropPoint.GetElement());
            messages = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(1, messages.Length);
            StringAssert.Contains(NarrativeErrorMessage, messages[0]);
            _outlook.Oc.CloseAllToastMessages();

            // Multiple email Drag And Drop
            var emailsToUpload = _outlook.AddTestEmailsToFolder(2, useDifferentTemplates: true);
            var subjectList = new List<string>(emailsToUpload.Keys);
            _outlook.OpenTestEmailFolder();
            _outlook.SelectAllItems();

            // Upload emails
            var firstSubject = subjectList[0];
            var secondSubject = subjectList[1];
            _outlook.DragAndDropEmailToOc(_outlook.GetNthItem(0), matterDetailsPage.DropPoint.GetElement());

            narrativesListPage.QuickSearch.SearchBy(firstSubject);
            Assert.NotNull(narrativesList.GetNarrativeListItemFromText(subjectList[0]), "List does not contain a narrative based on your search");
            narrativesListPage.QuickSearch.Close();
            narrativesListPage.QuickSearch.SearchBy(secondSubject);
            Assert.NotNull(narrativesList.GetNarrativeListItemFromText(subjectList[1]), "List does not contain a narrative based on your search");

            // Verify fields with any one of the narrative after multiple Drag And Drop
            createdNarrative = narrativesList.GetNarrativeListItemFromText(secondSubject);
            createdNarrative.Open();
            narrativesListPage.EditNarrativeDialog.Edit();

            type = narrativesListPage.EditNarrativeDialog.Controls["Narrative Type"].GetValue();
            descriptions = narrativesListPage.EditNarrativeDialog.Controls["Description"].GetValue();
            date = narrativesListPage.EditNarrativeDialog.Controls["Narrative Date"].GetValue();
            narratives = narrativesListPage.EditNarrativeDialog.Controls["Narrative"].GetValue();

            Assert.Multiple(() =>
            {
                Assert.IsNotNull(type, "Type is empty after  Drag And Drop ");
                Assert.IsNotNull(descriptions, "Description is empty after  Drag And Drop ");
                Assert.IsNotNull(date, "Date field is empty after  Drag And Drop ");
                Assert.IsNotNull(narratives, "Narrative is empty after  Drag And Drop ");
            });
            narrativesListPage.EditNarrativeDialog.Save();
            narrativesListPage.QuickSearch.Close();

            // Clean up
            narrativesList.GetNarrativeListItemFromText(subject).Delete().Confirm();
            narrativesList.GetNarrativeListItemFromText(firstSubject).Delete().Confirm();
            narrativesList.GetNarrativeListItemFromText(secondSubject).Delete().Confirm();
        }

        [Test]
        [Category(RegressionTestCategory)]
        [Description("TC:16521 - To verify single or multiple emails be quick filed to Narrative section and new Narrative is created accordingly")]
        public void QuickFileNarrative()
        {
            var subject = _outlook.AddTestEmailsToFolder(1, useDifferentTemplates: true).First().Key;
            _outlook.OpenTestEmailFolder();
            _outlook.SelectAllItems();

            var mattersListPage = _outlook.Oc.MattersListPage;
            var matterDetails = _outlook.Oc.MatterDetailsPage;
            var narrativesListPage = _outlook.Oc.NarrativesListPage;
            var narrativesList = narrativesListPage.ItemList;

            mattersListPage.Open();
            mattersListPage.ItemList.OpenFirst();
            matterDetails.Tabs.Open("Narratives");

            // Quick file single email
            matterDetails.QuickFile();
            narrativesListPage.QuickSearch.SearchBy(subject);
            Assert.NotNull(narrativesList.GetNarrativeListItemFromText(subject), "List does not contain a narrative based on your search");

            // Quick file same email again for duplicate message.
            matterDetails.QuickFile();
            var messages = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(1, messages.Length);
            StringAssert.Contains(NarrativeDuplicateTestMessage, messages[0]);
            _outlook.Oc.CloseAllToastMessages();

            // Validate Narrative fields has values
            var createdNarrative = narrativesList.GetNarrativeListItemFromText(subject);
            createdNarrative.Open();
            narrativesListPage.EditNarrativeDialog.Edit();

            var type = narrativesListPage.EditNarrativeDialog.Controls["Narrative Type"].GetValue();
            var descriptions = narrativesListPage.EditNarrativeDialog.Controls["Description"].GetValue();
            var date = narrativesListPage.EditNarrativeDialog.Controls["Narrative Date"].GetValue();
            var narratives = narrativesListPage.EditNarrativeDialog.Controls["Narrative"].GetValue();

            Assert.IsNotNull(type, "Type is empty after quick file");
            Assert.IsNotNull(descriptions, "Description is empty after quick file");
            Assert.IsNotNull(date, "Date field is empty after quick file");
            Assert.IsNotNull(narratives, "Narrative is empty after quick file");

            narrativesListPage.EditNarrativeDialog.Save();
            narrativesListPage.QuickSearch.Close();

            // Verify OC does not allow emails without subject
            _outlook.AddTestEmailsToFolder(1, subject: string.Empty, useDifferentTemplates: true);
            _outlook.OpenTestEmailFolder();
            _outlook.SelectAllItems();

            matterDetails.QuickFile();
            messages = _outlook.Oc.GetAllToastMessages();
            Assert.AreEqual(1, messages.Length);
            StringAssert.Contains(NarrativeErrorMessage, messages[0]);
            _outlook.Oc.CloseAllToastMessages();

            // Multiple email quick file
            var emailsToUploads = _outlook.AddTestEmailsToFolder(2, useDifferentTemplates: true);
            var subjectList = new List<string>(emailsToUploads.Keys);
            _outlook.OpenTestEmailFolder();
            _outlook.SelectAllItems();

            // Upload emails
            var firstSubject = subjectList[0];
            var secondSubject = subjectList[1];
            matterDetails.QuickFile();

            narrativesListPage.QuickSearch.SearchBy(firstSubject);
            Assert.NotNull(narrativesList.GetNarrativeListItemFromText(subjectList[0]), "List does not contain a narrative based on your search");
            narrativesListPage.QuickSearch.Close();
            narrativesListPage.QuickSearch.SearchBy(secondSubject);
            Assert.NotNull(narrativesList.GetNarrativeListItemFromText(subjectList[1]), "List does not contain a narrative based on your search");

            // Verify fields with any one of the narrative after multiple quick file
            createdNarrative = narrativesList.GetNarrativeListItemFromText(secondSubject);
            createdNarrative.Open();
            narrativesListPage.EditNarrativeDialog.Edit();

            type = narrativesListPage.EditNarrativeDialog.Controls["Narrative Type"].GetValue();
            descriptions = narrativesListPage.EditNarrativeDialog.Controls["Description"].GetValue();
            date = narrativesListPage.EditNarrativeDialog.Controls["Narrative Date"].GetValue();
            narratives = narrativesListPage.EditNarrativeDialog.Controls["Narrative"].GetValue();

            Assert.Multiple(() =>
            {
                Assert.IsNotNull(type, "Type is empty after quick file");
                Assert.IsNotNull(descriptions, "Description is empty after quick file");
                Assert.IsNotNull(date, "Date field is empty after quick file");
                Assert.IsNotNull(narratives, "Narrative is empty after quick file");
            });
            narrativesListPage.EditNarrativeDialog.Save();
            narrativesListPage.QuickSearch.Close();

            // Clean up
            narrativesList.GetNarrativeListItemFromText(subject).Delete().Confirm();
            narrativesList.GetNarrativeListItemFromText(firstSubject).Delete().Confirm();
            narrativesList.GetNarrativeListItemFromText(secondSubject).Delete().Confirm();
        }

        [TearDown]
        public void TearDown()
        {
            SaveScreenShotsAndLogs(_outlook);
            _outlook?.Destroy();
        }
    }
}
