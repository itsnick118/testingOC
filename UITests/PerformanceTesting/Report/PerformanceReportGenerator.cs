using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Web.UI;
using HtmlAgilityPack;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using TheArtOfDev.HtmlRenderer.PdfSharp;
using UITests.DataAccess;
using PageSize = PdfSharp.PageSize;

namespace UITests.PerformanceTesting.Report
{
    public class PerformanceReportGenerator
    {
        private readonly bool _hasBaseline;
        private DateTime _runDate;
        private readonly ISet<string> _tests;

        private readonly Dictionary<string, PerformanceTestRun> _testRuns;
        private readonly Dictionary<string, PerformanceTestRun> _baselineRuns;
        private readonly OrderedDictionary _parsedConfig;
        private readonly OrderedDictionary _parsedBaselineConfig;
        private readonly TestEnvironment _testEnvironment;

        public PerformanceReportGenerator(string baselineDirectory, PerformanceLog log)
        {
            _testRuns = new Dictionary<string, PerformanceTestRun>();
            _baselineRuns = new Dictionary<string, PerformanceTestRun>();
            _testEnvironment = new TestEnvironment(EnvironmentType.PerformanceTestEnvironment);

            if (log == null)
            {
                var logDirectory = new DirectoryInfo(_testEnvironment.TestLogDirectory).Parent?.EnumerateDirectories()
                    .Where(d => !d.Name.StartsWith("baseline"))
                    .OrderByDescending(d => d.CreationTime)
                    .First();

                if (logDirectory == null)
                {
                    Console.WriteLine(@"Log directory missing and in-memory log not provided. Exiting.");
                    return;
                }
                _runDate = logDirectory.CreationTime;

                _tests = new SortedSet<string>();

                var logFiles = logDirectory.GetFiles().ToList();
                foreach (var logFile in logFiles)
                {
                    if (logFile.Name.EndsWith("info"))
                    {
                        _parsedConfig = ParseConfig(logFile);
                    }
                    else
                    {
                        var testResult = new TestResultManager(logFile).TestRun;
                        _tests.Add(testResult.Id);
                        _testRuns[testResult.Id] = testResult;
                    }
                }
            }
            else
            {
                _runDate = DateTime.Now;
                _parsedConfig = ExecutingMachineInfo.AsOrderedDictionary();
                _tests = new HashSet<string>(log.AllLogEntries.Keys);

                var testSuiteRunId = TestResultDatabase.GetTestSuiteRunId();

                foreach (var test in _tests)
                {
                    var manager = new TestResultManager(test, log.AllLogEntries[test]);
                    manager.StoreTestRun(testSuiteRunId);
                    _testRuns[test] = manager.TestRun;
                }
            }

            _hasBaseline = baselineDirectory != null;
            if (baselineDirectory == null) return;

            var baselineFiles =
                new DirectoryInfo(
                    Path.Combine(_testEnvironment.TestLogDirectory, "..", baselineDirectory));

            if (!baselineFiles.Exists)
            {
                _hasBaseline = false;
                return;
            }

            foreach (var logFile in baselineFiles.GetFiles())
            {
                if (logFile.Name.EndsWith("info"))
                {
                    _parsedBaselineConfig = ParseConfig(logFile);
                }
                else
                {
                    var testResult = new TestResultManager(logFile).TestRun;
                    _baselineRuns[testResult.Id] = testResult;
                }
            }
        }

        public void Generate()
        {
            var pdfOutput = new PdfDocument();

            var headerDocument = GenerateCoverPage();

            var headerPdf = ImportablePdfFrom(headerDocument);

            foreach (PdfPage page in headerPdf.Pages)
            {
                pdfOutput.Pages.Add(page);
            }

            var testDocuments = GenerateTestPages();

            foreach (var testDocument in testDocuments)
            {
                var testPdf = ImportablePdfFrom(testDocument);

                foreach (PdfPage page in testPdf.Pages)
                {
                    pdfOutput.Pages.Add(page);
                }
            }

            SaveOutput(pdfOutput);
        }

        private HtmlDocument GenerateCoverPage()
        {
            var headerDocument = new HtmlDocument();
            headerDocument.LoadHtml(Resources.ReportTemplate);
            var headerBody = headerDocument.DocumentNode.SelectSingleNode("//body");

            headerBody.ChildNodes.Add(GetNode(HtmlTextWriterTag.H1, Resources.MemoryAndCpuReportTitle));
            headerBody.ChildNodes.Add(GetNode(HtmlTextWriterTag.H3, _runDate.ToShortDateString()));

            if (_hasBaseline)
            {
                var configTable = GetNode(HtmlTextWriterTag.Table);
                var configTableHeader = GetNode(HtmlTextWriterTag.Tr);

                configTableHeader.ChildNodes.Add(GetNode(HtmlTextWriterTag.Th));
                configTableHeader.ChildNodes.Add(GetNode(HtmlTextWriterTag.Th, "Current"));
                configTableHeader.ChildNodes.Add(GetNode(HtmlTextWriterTag.Th, "Baseline"));

                configTable.ChildNodes.Add(configTableHeader);

                foreach (string key in _parsedConfig.Keys)
                {
                    var configRow = GetNode(HtmlTextWriterTag.Tr, null, "info-list");
                    configRow.ChildNodes.Add(GetNode(HtmlTextWriterTag.Td, key));
                    var currentValue = _parsedConfig[key] != null ? _parsedConfig[key].ToString() : "N/A";
                    var baseValue = _parsedBaselineConfig[key] != null ? _parsedBaselineConfig[key].ToString() : "N/A";
                    configRow.ChildNodes.Add(GetNode(HtmlTextWriterTag.Td, currentValue));
                    configRow.ChildNodes.Add(GetNode(HtmlTextWriterTag.Td, baseValue));

                    configTable.ChildNodes.Add(configRow);
                }

                headerBody.ChildNodes.Add(configTable);
            }
            else
            {
                foreach (string key in _parsedConfig.Keys)
                {
                    headerBody.ChildNodes.Add(
                        GetNode(HtmlTextWriterTag.Div, $"{key}: {_parsedConfig[key]}", "info-list"));
                }
            }

            return headerDocument;
        }

        private IList<HtmlDocument> GenerateTestPages()
        {
            var testPages = new List<HtmlDocument>();

            foreach (var test in _tests)
            {
                var testDocument = new HtmlDocument();
                testDocument.LoadHtml(Resources.ReportTemplate);
                var testBody = testDocument.DocumentNode.SelectSingleNode("//body");

                var sectionContainer = GetNode(HtmlTextWriterTag.Table, null, "test-container");
                var sectionRow = GetNode(HtmlTextWriterTag.Tr);
                var section = GetNode(HtmlTextWriterTag.Td);

                var testInformation = new TestInformation(_testRuns[test], _hasBaseline ? _baselineRuns[test] : null);

                section.ChildNodes.Add(GetNode(HtmlTextWriterTag.H2, testInformation.Title));

                var stepsList = GetNode(HtmlTextWriterTag.Div);
                stepsList.ChildNodes.Add(GetNode(HtmlTextWriterTag.Div, "Step 0", "first-row"));
                var stepNumber = 1;

                foreach (var step in testInformation.TestSteps)
                {
                    var stepText = $"\t{stepNumber++}. " + step;
                    stepsList.ChildNodes.Add(GetNode(HtmlTextWriterTag.Div, stepText, "test-step"));
                }
                section.ChildNodes.Add(stepsList);

                var infoTable = GetNode(HtmlTextWriterTag.Table, null, "info-container");
                for (var i = 0; i < 2; i++) // Hack to make it show the first row
                {
                    var infoTableHeaderRow = GetNode(HtmlTextWriterTag.Tr);
                    var headerFields = new[] { "Metric", "Current", "Baseline", "Change" };
                    foreach (var field in headerFields)
                    {
                        var headerCell = GetNode(HtmlTextWriterTag.Th, field, "header");
                        infoTableHeaderRow.ChildNodes.Add(headerCell);
                    }
                    infoTable.ChildNodes.Add(infoTableHeaderRow);
                }

                foreach (var row in testInformation.Info)
                {
                    var infoTableRow = GetNode(HtmlTextWriterTag.Tr);

                    if (row[0] == Constants.Labels.CpuTwoSigmaMetric || row[0] == Constants.Labels.NetMemoryMetric)
                    {
                        infoTableRow.AddClass("highlight");
                    }

                    var cellNumber = 0;
                    foreach (var cell in row)
                    {
                        HtmlNode infoTableCell;
                        if (cellNumber == row.Length - 1)
                        {
                            var cssClass = cell.Contains("-")
                                ? "good"
                                : "bad";
                            infoTableCell = GetNode(HtmlTextWriterTag.Td, cell, cssClass);
                        }
                        else
                        {
                            infoTableCell = GetNode(HtmlTextWriterTag.Td, cell);
                        }
                        infoTableRow.ChildNodes.Add(infoTableCell);
                        cellNumber++;
                    }
                    infoTable.ChildNodes.Add(infoTableRow);
                }
                section.ChildNodes.Add(infoTable);

                var memoryPlotDiv = GetNode(HtmlTextWriterTag.Div, null, "plot-image");
                var memoryPlotImg = GetImageNode(testInformation.MemoryPlotFile);
                memoryPlotDiv.ChildNodes.Add(memoryPlotImg);

                var cpuPlotDiv = GetNode(HtmlTextWriterTag.Div, null, "plot-image");
                var cpuPlotImg = GetImageNode(testInformation.CpuPlotFile);
                cpuPlotDiv.ChildNodes.Add(cpuPlotImg);

                section.ChildNodes.Add(memoryPlotDiv);
                section.ChildNodes.Add(cpuPlotDiv);

                sectionRow.ChildNodes.Add(section);
                sectionContainer.ChildNodes.Add(sectionRow);
                testBody.ChildNodes.Add(sectionContainer);

                testPages.Add(testDocument);
            }

            return testPages;
        }

        private OrderedDictionary ParseConfig(FileSystemInfo configLogFile)
        {
            var parsedConfig = new OrderedDictionary();

            var file = File.ReadAllLines(configLogFile.FullName);
            foreach (var line in file)
            {
                var entry = line.Split(',');
                parsedConfig[entry[0]] = entry[1];
            }

            return parsedConfig;
        }

        private PdfDocument ImportablePdfFrom(HtmlDocument document)
        {
            var pdf = PdfGenerator.GeneratePdf(document.DocumentNode.OuterHtml, PageSize.Letter);
            using (var stream = new MemoryStream())
            {
                pdf.Save(stream, false);
                stream.Position = 0;
                var result = PdfReader.Open(stream, PdfDocumentOpenMode.Import);
                return result;
            }
        }

        private string GetOutputFileName()
        {
            var datePart = _runDate.ToShortDateString().Replace("/", "_");
            var timePart = _runDate.ToShortTimeString().Replace(":", "-").Replace(" ", "");
            return $"{datePart}@{timePart}.pdf";
        }

        private void SaveOutput(PdfDocument document)
        {
            var destination = new DirectoryInfo(_testEnvironment.TestOutputDirectory);

            if (!destination.Exists)
                destination.Create();

            var outputFile = Path.Combine(destination.FullName, GetOutputFileName());
            document.Save(outputFile);
        }

        private HtmlNode GetNode(HtmlTextWriterTag htmlTag, string innerString = null, string cssClass = null)
        {
            var tag = htmlTag.ToString();
            return HtmlNode.CreateNode($"<{tag} class='{cssClass}'>{innerString}</{tag}>");
        }

        private HtmlNode GetImageNode(string src)
        {
            var tag = HtmlTextWriterTag.Img.ToString();
            return HtmlNode.CreateNode($"<{tag} src='{src}' />");
        }
    }
}
