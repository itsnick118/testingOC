using NUnit.Framework;
using System;
using System.Collections.Generic;
using UITests.Models;
using UITests.PageModel;
using UITests.PerformanceTesting.Report;

namespace UITests.PerformanceTesting.Tests
{
    [TestFixture]
    public class OutlookStartUpTests
    {
        private TestEnvironment _testEnvironment;

        // ReSharper disable once NotAccessedField.Local
        private PerformanceTestHelpers _perfTestHelpers;

        private static PerformanceLog _log;
        private List<TimeToLoadModel> _list;
        private static bool _loginFlag = true;
        private string _methodName;

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
            _perfTestHelpers = new PerformanceTestHelpers(_log,
                Constants.WarmupTimeMilliseconds,
                Constants.CooldownTimeMilliseconds,
                Constants.WaitStepMilliseconds);
        }

        [Category("LoadTimeTest")]
        [Test]
        public void StartUpFirstRunTest()
        {
            _list = new List<TimeToLoadModel>();
            _methodName = TestContext.CurrentContext.Test.MethodName;
            for (int i = 0; i < 3; i++)
            {
                StartUpRun(1);
            }
        }

        [Category("LoadTimeTest")]
        [Test]
        public void StartUpSecondRunTest()
        {
            _list = new List<TimeToLoadModel>();
            _methodName = TestContext.CurrentContext.Test.MethodName;
            for (int i = 0; i < Constants.ReloadIterations; i++)
            {
                StartUpRun(2);
            }
        }

        [Category("LoadTimeTest")]
        [Test]
        public void StartUpThirdRunTest()
        {
            _list = new List<TimeToLoadModel>();
            _methodName = TestContext.CurrentContext.Test.MethodName;
            for (int i = 0; i < Constants.ReloadIterations; i++)
            {
                StartUpRun(3);
            }
        }

        /// <summary>
        /// gets load time based on login to spinners disappearance
        /// </summary>
        /// <param name="runs">
        /// will count for only run you want example 4 will only log for 4th time page is loaded
        /// </param>
        public void StartUpRun(int runs)
        {
            string guid = Guid.NewGuid().ToString();
            _testEnvironment.CleanUp();
            _testEnvironment.DeleteProfile();
            _testEnvironment.GenerateConfigFile();
            _testEnvironment.StartMockPassport();
            for (var i = 1; i <= runs; i++)
            {
                IList<TimeToLoadModel> tempList = new List<TimeToLoadModel>();
                //start outlook
                _outlook = new Outlook(_testEnvironment);
                _outlook.Launch();

                var startTime = DateTime.Now;

                tempList.Add(new TimeToLoadModel { LoadEvent = "startTime", TimeToLoad = startTime.Ticks, Id = guid });

                _log.StartNewRun(_outlook.Process);
                //start outlook
                if (_loginFlag)
                {
                    _outlook.Oc.BasicSettingsPage.LogIn();
                }
                _loginFlag = false;
                //start loading
                var startLoading = DateTime.Now.Subtract(startTime);
                tempList.Add(new TimeToLoadModel { LoadEvent = "startLoading", TimeToLoad = startLoading.Ticks, Id = guid });

                _outlook.Oc.BasicSettingsPage.CheckForLoading();

                //end loading
                var endLoading = DateTime.Now.Subtract(startTime);
                tempList.Add(new TimeToLoadModel { LoadEvent = "endLoading", TimeToLoad = endLoading.Ticks, Id = guid });
                tempList.Add(new TimeToLoadModel { LoadEvent = "LoadingTimeDiff", TimeToLoad = endLoading.Ticks - startLoading.Ticks, Id = guid });

                //kill
                _outlook.Destroy();
                var endTime = DateTime.Now.Subtract(startTime);
                tempList.Add(new TimeToLoadModel { LoadEvent = "endTime", TimeToLoad = endTime.Ticks, Id = guid });
                if (i == runs)
                {
                    _loginFlag = true;
                    _list.AddRange(tempList);
                }
            }
            _testEnvironment.CleanUp();
            Console.WriteLine($@"Test complete at: {DateTime.Now}");
        }

        [TearDown]
        public void TearDown()
        {
            _log.FinalizeLogs(_methodName, true);
            _log.WriteLogListToCsvWithName(_list, _methodName);
            Console.WriteLine($@"Test complete at: {DateTime.Now}");
        }

        [OneTimeTearDown, STAThread]
        public static void OneTimeTearDown()
        {
            new PerformanceReportGenerator(BaselineComparisonDirectory, _log).Generate();
        }
    }
}