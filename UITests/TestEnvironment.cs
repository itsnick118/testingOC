using System;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using IntegratedDriver;
using OpenQA.Selenium;

namespace UITests
{
    public class TestEnvironment
    {
        private readonly string _baseDirectory;
        private readonly EnvironmentType _environmentType;

        public readonly EnvironmentConfiguration? Configuration;

        public string BaseUrl { get; protected set; }
        public string StandardUser { get; set; }
        public string StandardUserPassword { get; set; }
        public string StandarUserName { get; set; }
        public string AttorneyUser { get; set; }
        public string AttorneyUserPassword { get; set; }
        public string ElevatedUser { get; set; }
        public string ElevatedUserPassword { get; set; }
        public string ElevatedUserDisplayName { get; set; }
        public string ElevatedUserPrimaryPABU { get; set; }
        public string ElevatedUserSecondaryPABU { get; set; }
        public string UserDataPath { get; set; }
        public string TestLogDirectory { get; set; }
        public string TestOutputDirectory { get; set; }
        public string OcDatabasePath { get; set; }
        public bool UseMockPassport { get; set; }

        public TestEnvironment(EnvironmentType environmentType, EnvironmentConfiguration? configuration = null)
        {
            _environmentType = environmentType;
            Configuration = configuration;

            ParseConfigForEnvironment();
            _baseDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        }

        public void CleanUp()
        {
            KillRunningProcesses();
            Windows.ClearWorkingTempFolder();
        }

        public void GenerateConfigFile()
        {
            var profileDirectory = new DirectoryInfo(
                Path.Combine(Environment.ExpandEnvironmentVariables(UserDataPath), "Profile"));

            if (!profileDirectory.Exists)
            {
                profileDirectory.Create();
            }

            var content = Resources.CompanionConfig;

            File.WriteAllText(Path.Combine(profileDirectory.ToString(), "companion.config"), content);
        }

        public void StartMockPassport()
        {
            if (!UseMockPassport) return;

            var mockPassportProcess = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = Path.Combine(_baseDirectory, "..\\..\\..\\MockPassport\\bin\\Debug", "MockPassport.exe"),
                    UseShellExecute = true,
                    RedirectStandardOutput = false,
                    RedirectStandardError = false,
                    WorkingDirectory = Path.Combine(_baseDirectory, "..\\..\\..\\MockPassport\\bin\\Debug"),
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Minimized
                }
            };

            mockPassportProcess.Start();
        }

        public void DeleteProfile()
        {
            const string rawPath = @"%USERPROFILE%\AppData\Local\Wolters Kluwer ELM Solutions\Office Companion\Profile";
            var profileDirectory = new DirectoryInfo(Environment.ExpandEnvironmentVariables(rawPath));

            if (profileDirectory.Exists)
                profileDirectory.Delete(true);
        }

        public void SaveToTestOutputDirectory(Screenshot screenShot, string testName)
        {
            var targetDirectory = Path.Combine(TestOutputDirectory, TestRunId(), testName);
            if (!Directory.Exists(targetDirectory))
            {
                Directory.CreateDirectory(targetDirectory);
            }

            var fullPath = Path.Combine(targetDirectory, "screenshot.png");

            screenShot.SaveAsFile(fullPath, ScreenshotImageFormat.Png);
        }

        public void CopyToTestOutputDirectory(string filename, string testName)
        {
            var targetDirectory = Path.Combine(TestOutputDirectory, TestRunId(), testName);
            if (!Directory.Exists(targetDirectory))
            {
                Directory.CreateDirectory(targetDirectory);
            }

            File.Copy(filename, Path.Combine(targetDirectory, new FileInfo(filename).Name));
        }

        private static void KillRunningProcesses()
        {
            var runningProcesses = true;

            while (runningProcesses)
            {
                runningProcesses = false;
                var processNames = Constants.ProcessName.Values.ToList();
                processNames.Add("MockPassport");
                processNames.Add("chromedriver");
                processNames.Add("PassportOffice.BrowserSubprocess.exe");

                foreach (var processName in processNames)
                {
                    foreach (var process in Process.GetProcessesByName(processName))
                    {
                        runningProcesses = true;

                        try
                        {
                            process.Kill();
                            process.WaitForExit();
                        }
                        catch (Win32Exception)
                        {
                            // ignore
                        }
                    }
                }
            }
        }

        private void ParseConfigForEnvironment()
        {
            var sharedSection = ConfigurationManager.GetSection("Shared") as NameValueCollection;

            if (!(ConfigurationManager.GetSection(_environmentType.ToString() + Configuration) is NameValueCollection section) || sharedSection == null)
            {
                return;
            }

            UserDataPath = Environment.ExpandEnvironmentVariables(sharedSection["localOfficeCompanionDataPath"]);

            var dbPath = string.Format(sharedSection["localOfficeCompanionDatabasePath"], section["elevatedUser"],
                new Uri(section["baseUrl"]).Host);
            OcDatabasePath = Path.Combine(UserDataPath, dbPath);

            BaseUrl = section["baseUrl"];
            if (!BaseUrl.EndsWith("/")) BaseUrl += '/';

            StandardUser = section["standardUser"];
            StandardUserPassword = section["standardPass"];
            StandarUserName = section["standardName"];
            AttorneyUser = section["attorneyUser"];
            AttorneyUserPassword = section["attorneyPass"];
            ElevatedUser = section["elevatedUser"];
            ElevatedUserPassword = section["elevatedPass"];
            ElevatedUserDisplayName = section["elevatedUserDisplayName"];
            ElevatedUserPrimaryPABU = section["elevatedUserPrimaryPABU"];
            ElevatedUserSecondaryPABU = section["elevatedUserSecondaryPABU"];
            TestLogDirectory = Path.Combine(UserDataPath, section["testLogDirectory"], TestRunId());
            TestOutputDirectory = Path.Combine(UserDataPath, section["testOutputDirectory"]);

            UseMockPassport = Convert.ToBoolean(section["useMockPassport"]);
        }

        public static string TestRunId()
        {
            var date = DateTime.UtcNow;

            var result = date.ToLocalTime();

            return result.ToString(CultureInfo.InvariantCulture)
                .Replace("/", "")
                .Replace(":", "_")
                .Replace(".", "_")
                .Replace(" ", "_");
        }
    }
}
