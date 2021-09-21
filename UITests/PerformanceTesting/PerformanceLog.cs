using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Reflection;
using System.Text;
using System.Threading;
using UITests.PerformanceTesting.Report;

namespace UITests.PerformanceTesting
{
    public class PerformanceLog
    {
        public IList<PerformanceLogEntry> CurrentLogEntries { get; private set; }
        public Dictionary<string, IList<PerformanceLogEntry>> AllLogEntries { get; }
        public ExecutingMachineInfo ExecutingMachineInfo { get; set; }

        private readonly TestEnvironment _environment;
        private readonly bool _saveToDisk;
        private Process _appProcess;
        private Dictionary<string, PerformanceCounter> _memoryCounters;
        private Dictionary<string, PerformanceCounter> _cpuCounters;

        public PerformanceLog(TestEnvironment environment, bool saveToDisk)
        {
            _environment = environment;
            _saveToDisk = saveToDisk;

            ExecutingMachineInfo = new ExecutingMachineInfo();

            AllLogEntries = new Dictionary<string, IList<PerformanceLogEntry>>();
        }

        public void StartNewRun(Process appProcess)
        {
            _appProcess = appProcess;
            CurrentLogEntries = new List<PerformanceLogEntry>();
            _memoryCounters = new Dictionary<string, PerformanceCounter>();
            _cpuCounters = new Dictionary<string, PerformanceCounter>();
            PerformanceCheckpoint("Start");
        }

        public void CheckPoint(int? iterations = null)
        {
            var currentMilliseconds = GetCurrentMilliseconds();
            var memory = GetCurrentMemoryKb();
            var cpu = GetCurrentCpuPercent();

            Console.WriteLine($@"{iterations},{currentMilliseconds},{memory},{cpu}");

            CurrentLogEntries.Add(new PerformanceLogEntry
            {
                Iteration = iterations,
                Time = currentMilliseconds,
                Memory = memory,
                Cpu = cpu
            });
        }

        public void WriteLogListToCsvWithName<T>(IList<T> timeToLoadList, string testMethodName)
        {
            var destination = new DirectoryInfo(_environment.TestLogDirectory);

            if (!destination.Exists)
            {
                destination.Create();
            }

            File.WriteAllText(
                Path.Combine(
                destination.ToString(), $"{testMethodName}TimeToLoad.csv"),
               ConvertListToCsvFormat(timeToLoadList));
        }

        public string ConvertListToCsvFormat<T>(IList<T> items)
        {
            StringBuilder sb = new StringBuilder();
            //Get all the properties by using reflection
            PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            List<string> ls = new List<string>();
            foreach (PropertyInfo prop in props)
            {
                //Setting column names as Property names
                ls.Add(prop.Name);
            }
            sb.AppendLine(string.Join(",", ls));
            foreach (T item in items)
            {
                var rowValues = new object[props.Length];
                for (int i = 0; i < props.Length; i++)
                {
                    rowValues[i] = props[i].GetValue(item, null);
                }
                sb.AppendLine(string.Join(",", rowValues));
            }
            return sb.ToString();
        }

        public void PerformanceCheckpoint(int? iterations = null)
        {
            var currentMilliseconds = GetCurrentMilliseconds();
            var memory = GetCurrentMemoryKb();
            var cpu = GetCurrentCpuPercent();

            Console.WriteLine($@"{iterations},{currentMilliseconds},{memory},{cpu}");

            CurrentLogEntries.Add(new PerformanceLogEntry
            {
                Iteration = iterations,
                Time = currentMilliseconds,
                Memory = memory,
                Cpu = cpu
            });
        }

        public void PerformanceCheckpoint(string logEvent, int? iterations = null)
        {
            Comment(logEvent);
            PerformanceCheckpoint(iterations);
        }

        public void Comment(string comment)
        {
            Console.WriteLine($@"// {comment}");

            CurrentLogEntries.Add(new PerformanceLogEntry
            {
                Comment = comment
            });
        }

        public void WaitWithPerformanceLogging(int waitMilliseconds, int step = 500)
        {
            for (var i = 0; i <= waitMilliseconds; i += step)
            {
                Thread.Sleep(step);
                CheckPoint();
            }
        }

        public void FinalizeLogs(string testMethodName, bool keepRun)
        {
            CloseCounters();

            if (_saveToDisk && keepRun)
            {
                Console.WriteLine($@"Closing file for {testMethodName}");
                var destination = new DirectoryInfo(_environment.TestLogDirectory);

                if (!destination.Exists)
                    destination.Create();

                var logFile = Path.Combine(destination.ToString(), $"{testMethodName}.log");

                File.WriteAllText(logFile, GetCurrentLogAsString());

                Console.WriteLine($@"Logs written to: {destination.FullName}");

                var localConfigLogFile = Path.Combine(destination.ToString(), "config.info");
                if (!File.Exists(localConfigLogFile))
                    File.WriteAllText(localConfigLogFile, ExecutingMachineInfo.ToString());
            }

            if (keepRun)
            {
                AllLogEntries[testMethodName] = CurrentLogEntries;
            }
        }

        private long GetCurrentMemoryKb()
        {
            PerformanceCounter baseProcessCounter;
            var process = _appProcess;

            if (!_memoryCounters.TryGetValue(process.ProcessName, out baseProcessCounter))
            {
                baseProcessCounter = new PerformanceCounter
                {
                    CategoryName = "Process",
                    CounterName = "Working Set - Private",
                    InstanceName = process.ProcessName
                };

                _memoryCounters[process.ProcessName] = baseProcessCounter;
            }

            var totalMemory = Convert.ToInt64(baseProcessCounter.NextValue() / 1024);

            var processNames = GetAllOcProcessNames(process.Id);

            foreach (var name in processNames)
            {
                PerformanceCounter chromeCounter;
                var counterIndex = name.Replace(".exe", string.Empty);

                if (!_memoryCounters.TryGetValue(counterIndex, out chromeCounter))
                {
                    chromeCounter = new PerformanceCounter
                    {
                        CategoryName = "Process",
                        CounterName = "Working Set - Private",
                        InstanceName = counterIndex
                    };
                    _memoryCounters[counterIndex] = chromeCounter;
                }

                totalMemory += Convert.ToInt64(chromeCounter.NextValue() / 1024);
            }

            return totalMemory;
        }

        private float GetCurrentCpuPercent()
        {
            var process = _appProcess;
            PerformanceCounter baseProcessCounter;

            if (!_cpuCounters.TryGetValue(process.ProcessName, out baseProcessCounter))
            {
                baseProcessCounter = new PerformanceCounter
                {
                    CategoryName = "Process",
                    CounterName = "% Processor Time",
                    InstanceName = process.ProcessName
                };

                _cpuCounters[process.ProcessName] = baseProcessCounter;
                GetNextNonZeroValue(baseProcessCounter);
            }

            var totalCpu = baseProcessCounter.NextValue();

            var processNames = GetAllOcProcessNames(process.Id);

            foreach (var name in processNames)
            {
                PerformanceCounter chromeCounter;
                var counterIndex = name.Replace(".exe", string.Empty);

                if (!_cpuCounters.TryGetValue(counterIndex, out chromeCounter))
                {
                    chromeCounter = new PerformanceCounter
                    {
                        CategoryName = "Process",
                        CounterName = "% Processor Time",
                        InstanceName = name.Replace(".exe", string.Empty)
                    };
                    GetNextNonZeroValue(chromeCounter);
                }

                // remove counter if not needed anymore !

                totalCpu += chromeCounter.NextValue();
            }

            return totalCpu;
        }

        private static long GetCurrentMilliseconds()
        {
            var timeSpan = DateTime.UtcNow;
            return timeSpan.Ticks / 10000;
        }

        private static IEnumerable<string> GetAllOcProcessNames(int parentProcessId)
        {
            var query = $"Select * from Win32_Process where ParentProcessId = {parentProcessId}";
            var managementObjectCollection = new ManagementObjectSearcher(query).Get();

            var resultList = new List<string>();
            foreach (var managementObject in managementObjectCollection)
            {
                var processName = managementObject.GetPropertyValue("Name").ToString();

                if (processName.Contains("PassportOffice.BrowserSubprocess"))
                {
                    resultList.Add(processName);
                }
            }

            return resultList;
        }

        private static void GetNextNonZeroValue(PerformanceCounter counter)
        {
            const int retryLimit = 30;

            for (var i = retryLimit; i > 0; i--)
            {
                var result = counter.NextValue();
                if (result > 0.001) return;
            }
        }

        private void CloseCounters()
        {
            if (_memoryCounters != null && _memoryCounters.Values.Count > 0)
            {
                foreach (var counter in _memoryCounters.Values)
                {
                    counter.Close();
                    counter.Dispose();
                }
            }

            if (_cpuCounters != null && _cpuCounters.Values.Count > 0)
            {
                foreach (var counter in _cpuCounters.Values)
                {
                    counter.Close();
                    counter.Dispose();
                }
            }
        }

        private string GetCurrentLogAsString()
        {
            var sb = new StringBuilder();

            foreach (var logEntry in CurrentLogEntries)
            {
                sb.AppendLine(string.IsNullOrWhiteSpace(logEntry.Comment)
                    ? $@"{logEntry.Iteration},{logEntry.Time},{logEntry.Memory},{logEntry.Cpu}"
                    : $@"//{logEntry.Comment}");
            }

            return sb.ToString();
        }
    }
}