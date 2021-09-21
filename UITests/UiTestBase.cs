using NUnit.Framework;
using NUnit.Framework.Interfaces;
using System;
using UITests.PageModel;

namespace UITests
{
    public abstract class UITestBase
    {
        protected TestEnvironment TestEnvironment { get; private set; }
        protected EnvironmentConfiguration Configuration { get; set; } = EnvironmentConfiguration.GA;

        [OneTimeSetUp]
        protected void OneTimeSetUpBase()
        {
            TestEnvironment = new TestEnvironment(EnvironmentType.UITestEnvironment, Configuration);
        }

        [SetUp]
        protected void SetUpBase()
        {
            TestEnvironment.CleanUp();
            TestEnvironment.DeleteProfile();
        }

        [TearDown]
        protected void TearDownBase()
        {
            TestEnvironment.CleanUp();
            Console.WriteLine($@"Test complete at: {DateTime.Now}");
        }

        protected void SaveScreenShotsAndLogs(AddInHost addInHost)
        {
            if (addInHost != null && TestContext.CurrentContext.Result.Outcome.Status == TestStatus.Failed)
            {
                TestHelpers.SaveScreenShotsAndLogs(addInHost, TestContext.CurrentContext.Test.Name);
            }
        }
    }
}
