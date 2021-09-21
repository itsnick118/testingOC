using IntegratedDriver;
using NUnit.Framework;
using NUnit.Framework.Interfaces;
using System;
using System.Threading;
using UITests.PageModel;
using UITests.PerformanceTesting.Report;

namespace UITests.PerformanceTesting.Tests
{
    [TestFixture]
    public class OutlookMemoryTests
    {
        private TestEnvironment _testEnvironment;
        private PerformanceTestHelpers _perfTestHelpers;
        private static PerformanceLog _log;

        private Outlook _outlook;

        private const string BaselineComparisonDirectory = "baseline_EYDrop2";

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            _testEnvironment = new TestEnvironment(EnvironmentType.PerformanceTestEnvironment);
            _log = new PerformanceLog(_testEnvironment, true);
        }

        [SetUp]
        public void SetUp()
        {
            _testEnvironment.CleanUp();
            _testEnvironment.DeleteProfile();
            _testEnvironment.GenerateConfigFile();
            _testEnvironment.StartMockPassport();

            _outlook = new Outlook(_testEnvironment);
            _outlook.Launch();

            _log.StartNewRun(_outlook.Process);

            _perfTestHelpers = new PerformanceTestHelpers(_log,
                Constants.WarmupTimeMilliseconds,
                Constants.CooldownTimeMilliseconds,
                Constants.WaitStepMilliseconds);
        }

        [Test]
        public void ReloadMatterList()
        {
            _log.PerformanceCheckpoint("Log in");
            _outlook.Oc.BasicSettingsPage.LogIn();

            _log.PerformanceCheckpoint("Open matter list");
            _outlook.Oc.Header.OpenMattersAppTab();

            for (var i = 0; i < Constants.FavoriteMatters; i++)
            {
                _outlook.Oc.MattersListPage.SetNthMatterAsFavorite(i);
            }

            _perfTestHelpers.WarmUp();

            _log.PerformanceCheckpoint("Start reload iterations");
            for (var i = 0; i < Constants.ReloadIterations; i++)
            {
                _outlook.Oc.MattersListPage.OpenAllMattersList();
                _outlook.Oc.MattersListPage.OpenFavoritesList();
                _log.PerformanceCheckpoint(i);
            }

            _perfTestHelpers.CoolDown();

            // clean up
            _outlook.Oc.MattersListPage.ClearAllFavorites();
        }

        [Test]
        public void ReloadMatterSummaryTabs()
        {
            _log.PerformanceCheckpoint("Log in");
            _outlook.Oc.BasicSettingsPage.LogIn();
            _log.PerformanceCheckpoint("Open first matter");

            _outlook.Oc.Header.OpenMattersAppTab();
            _outlook.Oc.MattersListPage.ItemList.OpenFirst();

            _perfTestHelpers.WarmUp();

            _log.PerformanceCheckpoint("Start reload iterations");
            for (var i = 0; i < Constants.ReloadIterations; i++)
            {
                _outlook.Oc.MatterDetailsPage.Tabs.Open("Emails");
                _outlook.Oc.MatterDetailsPage.Tabs.Open("Documents");
                _outlook.Oc.MatterDetailsPage.Tabs.Open("Narratives");
                _outlook.Oc.MatterDetailsPage.Tabs.Open("People");
                _outlook.Oc.MatterDetailsPage.Tabs.Open("Tasks/Events");
                _log.PerformanceCheckpoint(i);
            }

            _perfTestHelpers.CoolDown();
        }

        [Test]
        public void ReloadInvoiceList()
        {
            _log.PerformanceCheckpoint("Log in");
            _outlook.Oc.BasicSettingsPage.LogIn();
            _outlook.Oc.Header.OpenMattersAppTab();

            _perfTestHelpers.WarmUp();

            _log.PerformanceCheckpoint("Start reload iterations");
            for (var i = 0; i < Constants.ReloadIterations; i++)
            {
                _outlook.Oc.Header.OpenMattersAppTab();
                _outlook.Oc.Header.OpenSpendAppTab();
                _log.PerformanceCheckpoint(i);
            }

            _perfTestHelpers.CoolDown();
        }

        [Test]
        public void ScrollMatterList()
        {
            _log.PerformanceCheckpoint("Log in");
            _outlook.Oc.BasicSettingsPage.LogIn();

            _log.PerformanceCheckpoint("Open matter list");
            _outlook.Oc.Header.OpenMattersAppTab();

            _perfTestHelpers.WarmUp();

            _log.PerformanceCheckpoint("Start scrolling");
            var scrolls = 0;

            while (true)
            {
                var atBottom = !_outlook.Oc.MattersListPage.ItemList.ScrollDownIfNotAtBottom();
                if (atBottom) break;
                _log.PerformanceCheckpoint(scrolls++);
            }

            _perfTestHelpers.CoolDown();
        }

        [Test]
        public void ScrollInvoiceList()
        {
            _log.PerformanceCheckpoint("Log in");
            _outlook.Oc.BasicSettingsPage.LogIn();

            _log.PerformanceCheckpoint("Open invoice list");
            _outlook.Oc.Header.OpenSpendAppTab();

            _perfTestHelpers.WarmUp();

            _log.PerformanceCheckpoint("Start scrolling");
            var scrolls = 0;

            while (true)
            {
                var atBottom = !_outlook.Oc.InvoicesListPage.ItemList.ScrollDownIfNotAtBottom();
                if (atBottom) break;
                _log.PerformanceCheckpoint(scrolls++);
            }

            _perfTestHelpers.CoolDown();
        }

        [Test]
        public void UploadNEmailsAtOnce()
        {
            _outlook.AddTestEmailsToFolder(Constants.MassEmailCount, FileSize.Small);

            _log.PerformanceCheckpoint("Log in");
            _outlook.Oc.BasicSettingsPage.LogIn();

            _log.PerformanceCheckpoint("Open matter summary");
            _outlook.Oc.MattersListPage.ItemList.OpenFirst();
            _outlook.Oc.MatterDetailsPage.Tabs.Open("Emails");

            _perfTestHelpers.WarmUp();

            _outlook.OpenTestEmailFolder();

            _log.PerformanceCheckpoint("Start uploading emails");
            DragAndDrop.AllFromElementToElement(_outlook.GetCurrentMailList(), _outlook.Oc.MatterDetailsPage.DropPoint.GetElement());

            _outlook.Oc.WaitForQueueComplete();

            _perfTestHelpers.CoolDown();
        }

        [Test]
        public void UploadNEmailsOneAtATime()
        {
            _outlook.AddTestEmailsToFolder(Constants.MassEmailCount, FileSize.Small);

            _log.PerformanceCheckpoint("Log in");
            _outlook.Oc.BasicSettingsPage.LogIn();

            _log.PerformanceCheckpoint("Open matter summary");
            _outlook.Oc.MattersListPage.ItemList.OpenFirst();
            _outlook.Oc.MatterDetailsPage.Tabs.Open("Emails");

            _perfTestHelpers.WarmUp();

            _outlook.OpenTestEmailFolder();

            _log.PerformanceCheckpoint("Start uploading emails");

            for (var i = 0; i < Constants.MassEmailCount; i++)
            {
                var source = _outlook.GetNthEmailInTestFolder(i);
                DragAndDrop.FromElementToElement(source, _outlook.Oc.MatterDetailsPage.DropPoint.GetElement());
                _log.PerformanceCheckpoint(i);
            }

            _perfTestHelpers.CoolDown();
        }

        [Test]
        public void OpenExpandAndCloseNInspectorWindows()
        {
            _outlook.AddTestEmailsToFolder(Constants.InspectorWindowCount, FileSize.Small);

            _log.PerformanceCheckpoint("Log in");
            _outlook.Oc.BasicSettingsPage.LogIn();

            _perfTestHelpers.WarmUp();

            _outlook.OpenTestEmailFolder();

            _log.PerformanceCheckpoint("Start opening inspector windows");
            for (var i = 0; i < Constants.InspectorWindowCount; i++)
            {
                var mailInspectorWindow = _outlook.OpenNthTestEmail(i);
                var isExpanded = _outlook.ToggleTaskPane(mailInspectorWindow);
                Console.WriteLine(isExpanded);

                if (isExpanded)
                {
                    Thread.Sleep(5000);
                    _outlook.Oc.WaitForLoadComplete();
                    Thread.Sleep(2000);
                }

                _outlook.CloseInspector(mailInspectorWindow);
                _log.PerformanceCheckpoint(i);
            }

            _perfTestHelpers.CoolDown();
        }

        [Test]
        public void OpenAndExpandNInspectorWindowsWithoutClosing()
        {
            _outlook.AddTestEmailsToFolder(10, FileSize.Small);

            _log.PerformanceCheckpoint("Log in");
            _outlook.Oc.BasicSettingsPage.LogIn();

            _perfTestHelpers.WarmUp();

            _outlook.OpenTestEmailFolder();

            _log.PerformanceCheckpoint("Start opening inspector windows");
            for (var i = 0; i < Constants.InspectorWindowCount; i++)
            {
                var mailInspectorWindow = _outlook.OpenNthTestEmail(i);
                var isExpanded = _outlook.ToggleTaskPane(mailInspectorWindow);
                Console.WriteLine(isExpanded);

                if (isExpanded)
                {
                    _outlook.Oc.WaitForLoadComplete();
                    Thread.Sleep(2000);
                }

                _log.PerformanceCheckpoint(i);
            }

            _outlook.CloseAllInspectors();

            _perfTestHelpers.CoolDown();
        }

        [TearDown]
        public void TearDown()
        {
            var methodName = TestContext.CurrentContext.Test.MethodName;

            _log.FinalizeLogs(methodName,
                TestContext.CurrentContext.Result.Outcome.Status.Equals(TestStatus.Passed));
            _outlook.Destroy();
            _testEnvironment.CleanUp();

            Console.WriteLine($@"Test complete at: {DateTime.Now}");
        }

        [OneTimeTearDown, STAThread]
        public static void OneTimeTearDown()
        {
            new PerformanceReportGenerator(BaselineComparisonDirectory, _log).Generate();
        }
    }
}