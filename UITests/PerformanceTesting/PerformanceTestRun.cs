using System;
using System.Collections.Generic;
using OpenQA.Selenium;
using UITests.PerformanceTesting.Report;

namespace UITests.PerformanceTesting
{
    public class PerformanceTestRun
    {
        public string TestTitle { get; set; }

        public string Id { get; set; }

        public int N { get; set; }

        public TimeSpan TotalTestTime { get; set; }

        public CpuPerformanceData CpuOneSigma { get; set; }

        public CpuPerformanceData CpuTwoSigma { get; set; }

        public CpuPerformanceData Cpu { get; set; }

        public MemoryData MemoryData { get; set; }

        public IEnumerable<LogEntry> JavaScriptErrors { get; set; }

        public IEnumerable<LogEntry> JavaScriptWarnings { get; set; }

        public IList<PerformanceLogEntry> Entries { get; set; }
    }
}
