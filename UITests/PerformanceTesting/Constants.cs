using System.Collections.Generic;

namespace UITests.PerformanceTesting
{
    public class Constants
    {
        public const int FavoriteMatters = 15;
        public const int ReloadIterations = 50;
        public const int MassEmailCount = 25;
        public const int InspectorWindowCount = 10;
        public const int WarmupTimeMilliseconds = 10000;
        public const int CooldownTimeMilliseconds = 10000;
        public const int WaitStepMilliseconds = 1000;

        public static IDictionary<string, int> NLookup = new Dictionary<string, int>
        {
            { "ReloadInvoiceList", ReloadIterations },
            { "ReloadMatterList", ReloadIterations },
            { "ReloadMatterSummaryTabs", ReloadIterations },
            { "OpenAndExpandNInspectorWindowsWithoutClosing", InspectorWindowCount },
            { "OpenExpandAndCloseNInspectorWindows", InspectorWindowCount },
            { "UploadNEmailsAtOnce", MassEmailCount },
            { "UploadNEmailsOneAtATime", MassEmailCount },
        };

        public class Labels
        {
            public const string RunningTimeMetric = "Running Time";
            public const string CpuOneSigmaMetric = "Mean/Max CPU, 1σ";
            public const string CpuTwoSigmaMetric = "Mean/Max CPU, 2σ";
            public const string CpuAllMetric = "Mean/Max CPU";
            public const string MaxMemoryMetric = "Max Memory";
            public const string NetMemoryMetric = "Net Memory";
            public const string MemoryPercentDeltaMetric = "Memory Percent Change";

            public const string Current = "Current";
            public const string Baseline = "Baseline";

            public const string MemoryYAxis = "Memory Consumption (MB)";
            public const string CpuYAxis = "% CPU Utilization";
            public const string XAxis = "Seconds";
        }
    }
}