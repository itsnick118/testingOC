namespace UITests.PerformanceTesting.Report
{
    public class PerformanceLogEntry
    {
        public int Id { get; set; }

        public int? Iteration { get; set; }

        public long Time { get; set; }

        public string Comment { get; set; }

        public long Memory { get; set; }

        public float Cpu { get; set; }
    }
}
