using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Resources;
using System.Threading;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Wpf;
using LineSeries = OxyPlot.Series.LineSeries;
using Series = OxyPlot.Series.Series;

namespace UITests.PerformanceTesting.Report
{
    [SuppressMessage("ReSharper", "PossibleLossOfFraction")]
    public class TestInformation
    {
        private readonly PerformanceTestRun _testRun;
        private readonly PerformanceTestRun _baselineRun;
        private readonly IList<DataPoint> _currentCpuData;
        private readonly IList<DataPoint> _currentMemoryData;
        private readonly IList<DataPoint> _baseCpuData;
        private readonly IList<DataPoint> _baseMemoryData;

        public string Title { get; }
        public string[] TestSteps { get; }
        public string MemoryPlotFile { get; }
        public string CpuPlotFile { get; }
        public string[][] Info { get; }

        private const string PctFormat = "0.##";
        private const string MemoryFormat = "n0";

        public TestInformation(PerformanceTestRun testRun, PerformanceTestRun baselineRun=null)
        {
            _testRun = testRun;
            _baselineRun = baselineRun;

            _currentCpuData = GetDataPoints(_testRun, TestType.Cpu);
            _currentMemoryData = GetDataPoints(_testRun, TestType.Memory);

            Title = GenerateTitle();
            TestSteps = GenerateTestSteps();

            if (baselineRun != null)
            {
                _baseCpuData = GetDataPoints(_baselineRun, TestType.Cpu);
                _baseMemoryData = GetDataPoints(_baselineRun, TestType.Memory);
                Info = GenerateInfoWithBaseline();
            }
            else
            {
                Info = GenerateInfo();
            }

            MemoryPlotFile = GetPlotImageFile(TestType.Memory);
            CpuPlotFile = GetPlotImageFile(TestType.Cpu);
        }

        private string GenerateTitle()
        {
            var textInfo = new CultureInfo("en-US", false).TextInfo;

            return textInfo.ToTitleCase(
                _testRun.TestTitle.Replace(" N ",
                    $" {(_testRun.N > 0 ? _testRun.N.ToString() : " N ")} ").Replace('_', ' '));
        }

        private string[] GenerateTestSteps()
        {
            var resourceManager = new ResourceManager(typeof(Resources));
            var steps = (string)resourceManager.GetObject(_testRun.Id);

            if (steps == null) throw new Exception("Identifier did not map to object in Resources.resx.");

            if (_testRun.N > 0)
                steps = steps.Replace("{n}", _testRun.N.ToString());

            return steps.Split(new[] {"\r\n"}, StringSplitOptions.RemoveEmptyEntries);
        }

        private string[][] GenerateInfo()
        {
            var cores = Convert.ToInt32(ExecutingMachineInfo.AsOrderedDictionary()[ExecutingMachineInfo.LogicalProcessors]);
            var resultList = new List<string[]>
            {
                new []
                {
                    Constants.Labels.RunningTimeMetric,
                    $"{_testRun.TotalTestTime:mm\\:ss}"
                },
                new[]
                {
                    Constants.Labels.CpuOneSigmaMetric,
                    $"{(_testRun.CpuOneSigma.Mean / cores).ToString(PctFormat)}% / " +
                    $"{(_testRun.CpuOneSigma.Max / cores).ToString(PctFormat)}%"
                },
                new[]
                {
                    Constants.Labels.CpuTwoSigmaMetric,
                    $"{(_testRun.CpuTwoSigma.Mean / cores).ToString(PctFormat)}% / " +
                    $"{(_testRun.CpuTwoSigma.Max / cores).ToString(PctFormat)}%"
                },
                new[]
                {
                    Constants.Labels.CpuAllMetric,
                    $"{(_testRun.Cpu.Mean / cores).ToString(PctFormat)}% / " +
                    $"{(_testRun.Cpu.Max / cores).ToString(PctFormat)}%"
                },
                new[]
                {
                    Constants.Labels.MaxMemoryMetric,
                    $"{_testRun.MemoryData.Max.ToString(MemoryFormat)}MB"
                },
                new[]
                {
                    Constants.Labels.NetMemoryMetric,
                    $"{_testRun.MemoryData.Net.ToString(MemoryFormat)}MB"
                },
                new[]
                {
                    Constants.Labels.MemoryPercentDeltaMetric,
                    $"{_testRun.MemoryData.PercentChange.ToString(PctFormat)}%"
                }
            };

            return resultList.ToArray();
        }

        private string[][] GenerateInfoWithBaseline()
        {
            var cores = Convert.ToInt32(ExecutingMachineInfo.AsOrderedDictionary()[ExecutingMachineInfo.LogicalProcessors]);
            var shorterTime = _testRun.TotalTestTime < _baselineRun.TotalTestTime;
            var resultList = new List<string[]>
            {
                new []
                {
                    Constants.Labels.RunningTimeMetric,
                    $"{_testRun.TotalTestTime:mm\\:ss}",
                    $"{_baselineRun.TotalTestTime:mm\\:ss}",
                    $"{(shorterTime ? "-" : string.Empty)}{_testRun.TotalTestTime - _baselineRun.TotalTestTime:mm\\:ss}"
                },
                new[]
                {
                    Constants.Labels.CpuOneSigmaMetric,
                    $"{(_testRun.CpuOneSigma.Mean / cores).ToString(PctFormat)}% / " +
                    $"{(_testRun.CpuOneSigma.Max / cores).ToString(PctFormat)}% ",
                    $"{(_baselineRun.CpuOneSigma.Mean / cores).ToString(PctFormat)}% / " +
                    $"{(_baselineRun.CpuOneSigma.Max / cores).ToString(PctFormat)}%",
                    $"{((_testRun.CpuOneSigma.Mean - _baselineRun.CpuOneSigma.Mean) / cores).ToString(PctFormat)}% / " +
                    $"{((_testRun.CpuOneSigma.Max - _baselineRun.CpuOneSigma.Max) / cores).ToString(PctFormat)}%"
                },
                new[]
                {
                    Constants.Labels.CpuTwoSigmaMetric,
                    $"{(_testRun.CpuTwoSigma.Mean / cores).ToString(PctFormat)}% / " +
                    $"{(_testRun.CpuTwoSigma.Max / cores).ToString(PctFormat)}%",
                    $"{(_baselineRun.CpuTwoSigma.Mean / cores).ToString(PctFormat)}% / " +
                    $"{(_baselineRun.CpuTwoSigma.Max / cores).ToString(PctFormat)}%",
                    $"{((_testRun.CpuTwoSigma.Mean - _baselineRun.CpuTwoSigma.Mean) / cores).ToString(PctFormat)}% / " +
                    $"{((_testRun.CpuTwoSigma.Max - _baselineRun.CpuTwoSigma.Max) / cores).ToString(PctFormat)}%"
                },
                new[]
                {
                    Constants.Labels.CpuAllMetric,
                    $"{(_testRun.Cpu.Mean / cores).ToString(PctFormat)}% /" +
                    $"{(_testRun.Cpu.Max / cores).ToString(PctFormat)}%",
                    $"{(_baselineRun.Cpu.Mean / cores).ToString(PctFormat)}% /" +
                    $"{(_baselineRun.Cpu.Max / cores).ToString(PctFormat)}%",
                    $"{((_testRun.Cpu.Mean - _baselineRun.Cpu.Mean) / cores).ToString(PctFormat)}% / " +
                    $"{((_testRun.Cpu.Max - _baselineRun.Cpu.Max) / cores).ToString(PctFormat)}%"
                },
                new[]
                {
                    Constants.Labels.MaxMemoryMetric,
                    $"{(_testRun.MemoryData.Max / 1000).ToString(MemoryFormat)}MB",
                    $"{(_baselineRun.MemoryData.Max / 1000).ToString(MemoryFormat)}MB",
                    $"{(_testRun.MemoryData.Max - _baselineRun.MemoryData.Max) / 1000}MB"
                },
                new[]
                {
                    Constants.Labels.NetMemoryMetric,
                    $"{(_testRun.MemoryData.Net / 1000).ToString(MemoryFormat)}MB",
                    $"{(_baselineRun.MemoryData.Net / 1000).ToString(MemoryFormat)}MB",
                    $"{(_testRun.MemoryData.Net - _baselineRun.MemoryData.Net) / 1000}MB"
                },
                new[]
                {
                    Constants.Labels.MemoryPercentDeltaMetric,
                    $"{_testRun.MemoryData.PercentChange.ToString(PctFormat)}%",
                    $"{_baselineRun.MemoryData.PercentChange.ToString(PctFormat)}%",
                    (_testRun.MemoryData.PercentChange - _baselineRun.MemoryData.PercentChange)
                     .ToString(PctFormat) + "%"
                }
            };

            return resultList.ToArray();
        }

        private IList<DataPoint> GetDataPoints(PerformanceTestRun testRun, TestType testType)
        {
            var timeZero = testRun.Entries.First(e => e.Time != 0).Time;
            var cores = Convert.ToInt32(ExecutingMachineInfo.AsOrderedDictionary()[ExecutingMachineInfo.LogicalProcessors]);

            var dataPoints = new List<DataPoint>();
            const int xScale = 1000;
            var yScale = testType == TestType.Cpu ? cores : 1024;

            foreach (var logEntry in testRun.Entries)
            {
                var value = testType == TestType.Cpu ? logEntry.Cpu : logEntry.Memory;

                if (string.IsNullOrWhiteSpace(logEntry.Comment) && logEntry.Time != 0)
                {
                    dataPoints.Add(
                        new DataPoint((logEntry.Time - timeZero) / xScale, value / yScale));
                }
            }

            return dataPoints;
        }

        private Series GetSeries(IList<DataPoint> dataPoints, string title)
        {
            var series = new LineSeries
            {
                ItemsSource = dataPoints,
                Title = title
            };

            return series;
        }

        private string GetPlotImageFile(TestType testType)
        {
            string resultFileName;
            var model = new PlotModel {DefaultFontSize = 8};

            if (testType == TestType.Memory)
            {
                resultFileName = $"{_testRun.Id}_memory.png";
                var dataSeries = GetSeries(_currentMemoryData, Constants.Labels.Current);

                if (_testRun != null)
                {
                    var baseSeries = GetSeries(_baseMemoryData, Constants.Labels.Baseline);
                    model.Series.Add(baseSeries);
                }

                model.Axes.Add(new OxyPlot.Axes.LinearAxis
                {
                    Position = AxisPosition.Left,
                    Title = Constants.Labels.MemoryYAxis,
                    Minimum = 0,
                    Maximum = 1024,
                    MajorStep = 128,
                    MinorStep = 32,
                    MajorGridlineStyle = LineStyle.Dot
                });

                model.Axes.Add(new OxyPlot.Axes.LinearAxis
                {
                    Position = AxisPosition.Bottom,
                    Title = Constants.Labels.XAxis
                });

                model.Series.Add(dataSeries);
            }
            else
            {
                resultFileName = $"{_testRun.Id}_cpu.png";
                var dataSeries = GetSeries(_currentCpuData, Constants.Labels.Current);

                if (_baseCpuData != null)
                {
                    var baseSeries = GetSeries(_baseCpuData, Constants.Labels.Baseline);
                    model.Series.Add(baseSeries);
                }

                model.Series.Add(dataSeries);

                model.Axes.Add(new OxyPlot.Axes.LinearAxis
                {
                    Minimum = 0,
                    Maximum = 100,
                    Position = AxisPosition.Left,
                    Title = Constants.Labels.CpuYAxis,
                    MajorStep = 10,
                    MinorStep = 5,
                    MajorGridlineStyle = LineStyle.Dot
                });

                model.Axes.Add(new OxyPlot.Axes.LinearAxis
                {
                    Position = AxisPosition.Bottom,
                    Title = Constants.Labels.XAxis
                });
            }

            var thread = new Thread(() => PngExporter.Export(model, resultFileName, 485, 190, OxyColors.White));

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            return resultFileName;
        }
    }
}
