namespace UITests.PerformanceTesting
{
    public class PerformanceTestHelpers
    {
        private readonly PerformanceLog _performanceLog;
        private readonly int _warmupMilliseconds;
        private readonly int _cooldownMilliseconds;
        private readonly int _loggingStepMilliseconds;

        public PerformanceTestHelpers(PerformanceLog performanceLog, int warmupMilliseconds, int cooldownMilliseconds, int loggingStepMilliseconds)
        {
            _performanceLog = performanceLog;
            _warmupMilliseconds = warmupMilliseconds;
            _cooldownMilliseconds = cooldownMilliseconds;
            _loggingStepMilliseconds = loggingStepMilliseconds;
        }

        public void WarmUp()
        {
            _performanceLog.Comment("Starting warmup");
            _performanceLog.WaitWithPerformanceLogging(_warmupMilliseconds, _loggingStepMilliseconds);
        }

        public void CoolDown()
        {
            _performanceLog.Comment("Starting cooldown");
            _performanceLog.WaitWithPerformanceLogging(_cooldownMilliseconds, _loggingStepMilliseconds);
        }
    }
}