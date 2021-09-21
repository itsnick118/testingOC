using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using UITests.DataAccess;
using UITests.PerformanceTesting.Report;

namespace UITests.PerformanceTesting
{
    public class TestResultManager
    {
        private readonly IList<PerformanceLogEntry> _performanceLogs;
        public readonly PerformanceTestRun TestRun;

        public TestResultManager(string testId,
            IList<PerformanceLogEntry> performanceLogs)
        {
            _performanceLogs = performanceLogs;
            TestRun = ProcessTestRun(testId);
        }

        public TestResultManager(FileInfo fileInfo)
        {
            _performanceLogs = new List<PerformanceLogEntry>();
            TestRun = ProcessTestRunFromFile(fileInfo);
        }

        public void StoreTestRun(int testSuiteRunId)
        {
            try
            {
                TestResultDatabase.InsertTestRun(TestRun, testSuiteRunId);
            }
            catch (Exception exception)
            {
                Console.WriteLine(@"Could not insert run into database: " + exception.Message);
            }
        }

        private PerformanceTestRun ProcessTestRun(string testId)
        {
            var testRun = new PerformanceTestRun
            {
                Id = testId,
                TestTitle = Regex.Replace(testId, "(\\B[A-Z])", " $1").Replace('_', ' '),
                Entries = _performanceLogs
            };

            var nonZeroCpu = _performanceLogs.Select(p => p.Cpu).Where(c => Math.Abs(c) > 0.0001).ToArray();
            Array.Sort(nonZeroCpu);

            var oneSigmaSample = Convert.ToInt32(Math.Floor(nonZeroCpu.Length * 0.6827));

            var oneSigmaCpu = nonZeroCpu
                .Skip((nonZeroCpu.Length - oneSigmaSample) / 2).Take(oneSigmaSample).ToArray();

            testRun.CpuOneSigma = new CpuPerformanceData
            {
                Max = oneSigmaCpu.Max(),
                Mean = oneSigmaCpu.Average(),
            };

            var twoSigmaSample = Convert.ToInt32(Math.Floor(nonZeroCpu.Length * 0.9545));

            var twoSigmaCpu = nonZeroCpu
                .Skip((nonZeroCpu.Length - twoSigmaSample) / 2).Take(twoSigmaSample).ToArray();

            testRun.CpuTwoSigma = new CpuPerformanceData
            {
                Max = twoSigmaCpu.Max(),
                Mean = twoSigmaCpu.Average(),
            };

            testRun.Cpu = new CpuPerformanceData
            {
                Max = nonZeroCpu.Max(),
                Mean = nonZeroCpu.Average(),
            };

            var nonZeroMemory = _performanceLogs.Select(p => p.Memory).Where(m => m > 0).ToArray();

            testRun.MemoryData = new MemoryData
            {
                Max = nonZeroMemory.Max(),
                Net = nonZeroMemory.Max() - nonZeroMemory.Min(),
                PercentChange = 100 * (nonZeroMemory.Max() - nonZeroMemory.Min()) / Convert.ToSingle(nonZeroMemory.Min())
            };

            var startTime = testRun.Entries.First(e => e.Time != 0).Time;
            var endTime = testRun.Entries.Last(e => e.Time != 0).Time;
            testRun.TotalTestTime = TimeSpan.FromMilliseconds(endTime - startTime);

            if (Constants.NLookup.ContainsKey(testRun.Id))
            {
                testRun.N = Constants.NLookup[testRun.Id];
            }

            return testRun;
        }

        private PerformanceTestRun ProcessTestRunFromFile(FileInfo fileInfo)
        {
            var file = File.ReadAllLines(fileInfo.FullName);
            var id = 0;

            foreach (var line in file)
            {
                var logEntry = new PerformanceLogEntry
                {
                    Id = id++
                };

                if (line.StartsWith("//"))
                {
                    logEntry.Comment = line.TrimStart(' ', '/');
                    _performanceLogs.Add(new PerformanceLogEntry
                    {
                        Id = id++,
                    });
                }
                else
                {
                    var split = line.Split(',');

                    if (!string.IsNullOrWhiteSpace(split[0]))
                    {
                        logEntry.Iteration = int.Parse(split[0]);
                    }
                    logEntry.Time = Convert.ToInt64(split[1]);
                    logEntry.Memory = Convert.ToInt32(split[2]);
                    logEntry.Cpu = float.Parse(split[3]);
                }

                _performanceLogs.Add(logEntry);
            }

            var testId = Path.GetFileNameWithoutExtension(fileInfo.Name);

            return ProcessTestRun(testId);
        }
    }
}
