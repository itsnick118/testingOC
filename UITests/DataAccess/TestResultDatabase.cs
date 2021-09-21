using System;
using System.Collections.Specialized;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Resources;
using UITests.PerformanceTesting;
using Exception = System.Exception;

namespace UITests.DataAccess
{
    public class TestResultDatabase
    {
        private static readonly OrderedDictionary Environment = ExecutingMachineInfo.AsOrderedDictionary();

        private static readonly string PerformanceDbConnectionString =
            ConfigurationManager.ConnectionStrings["db-eng-oc-qe-perf"].ConnectionString;

        public static void InsertTestRun(PerformanceTestRun testRun, int testSuiteRunId)
        {
            var testDefinitionId = FindOrCreateTestDefinitionAndGetId(testRun);
            InsertTestRun(testRun, testDefinitionId, testSuiteRunId);
        }

        public static int GetTestSuiteRunId()
        {
            var machineId = FindOrCreateMachineAndGetId();
            var testSuiteRunId = -1;

            const string testSuiteRunInsert =
                "Insert into TEST_SUITE_RUN (branch_name, date_time, machine_id) " +
                "output INSERTED.test_suite_run_id " +
                "values (@branch_name, GetUtcDate(), @machine_id)";

            using (var connection = new SqlConnection(PerformanceDbConnectionString))
            {
                connection.Open();

                using (var cmd = new SqlCommand(testSuiteRunInsert, connection))
                {
                    cmd.Parameters.AddWithValue("@branch_name", Environment[ExecutingMachineInfo.BranchName]);
                    cmd.Parameters.AddWithValue("@machine_id", machineId);

                    var result = cmd.ExecuteScalar();

                    if (result != null)
                    {
                        testSuiteRunId = (int)result;
                    }

                    if (connection.State == ConnectionState.Open)
                        connection.Close();
                }
            }

            return testSuiteRunId;
        }

        public static void InsertTestRun(PerformanceTestRun testRun, int testDefinitionId, int testSuiteRunId)
        {
            const string testRunInsert =
                "Insert into TEST_RUN ([time], [cpu_max_one_sigma], [cpu_max_two_sigma], [cpu_max], [cpu_mean_one_sigma], " +
                "[cpu_mean_two_sigma], [cpu_mean], [memory_max], [memory_net], [test_definition_id], [test_suite_run_id]) " +
                "values (@time, @cpu_max_one_sigma, @cpu_max_two_sigma, @cpu_max, @cpu_mean_one_sigma, " +
                "@cpu_mean_two_sigma, @cpu_mean, @memory_max, @memory_net, @test_definition_id, @test_suite_run_id)";

            using (var connection = new SqlConnection(PerformanceDbConnectionString))
            {
                connection.Open();

                using (var cmd = new SqlCommand(testRunInsert, connection))
                {
                    cmd.Parameters.AddWithValue("@time", testRun.TotalTestTime.TotalMilliseconds);
                    cmd.Parameters.AddWithValue("@cpu_max_one_sigma", testRun.CpuOneSigma.Max);
                    cmd.Parameters.AddWithValue("@cpu_max_two_sigma", testRun.CpuTwoSigma.Max);
                    cmd.Parameters.AddWithValue("@cpu_max", testRun.Cpu.Max);
                    cmd.Parameters.AddWithValue("@cpu_mean_one_sigma", testRun.CpuOneSigma.Mean);
                    cmd.Parameters.AddWithValue("@cpu_mean_two_sigma", testRun.CpuTwoSigma.Mean);
                    cmd.Parameters.AddWithValue("@cpu_mean", testRun.Cpu.Mean);
                    cmd.Parameters.AddWithValue("@memory_max", testRun.MemoryData.Max);
                    cmd.Parameters.AddWithValue("@memory_net", testRun.MemoryData.Net);
                    cmd.Parameters.AddWithValue("@test_definition_id", testDefinitionId);
                    cmd.Parameters.AddWithValue("@test_suite_run_id", testSuiteRunId);

                    cmd.ExecuteNonQuery();

                    if (connection.State == ConnectionState.Open)
                        connection.Close();
                }
            }
        }

        public static void CreateTestRunEntries(PerformanceTestRun testRun, int testRunId)
        {
            const string testRunInsert =
                "Insert into TEST_RUN_ENTRY ([comment], [iteration], [memory], [cpu], [time], [test_run_id]) " +
                "values (@comment, @iteration, @memory, @cpu, @time, @test_run_id)";

            using (var connection = new SqlConnection(PerformanceDbConnectionString))
            {
                connection.Open();

                using (var cmd = new SqlCommand(testRunInsert, connection))
                {
                    cmd.Parameters.Add("@comment", SqlDbType.NVarChar, 128);
                    cmd.Parameters.Add("@iteration", SqlDbType.Int);
                    cmd.Parameters.Add("@memory", SqlDbType.BigInt);
                    cmd.Parameters.Add("@cpu", SqlDbType.Float);
                    cmd.Parameters.Add("@time", SqlDbType.BigInt);
                    cmd.Parameters.Add("@test_run_id", SqlDbType.Int);

                    foreach (var logEntry in testRun.Entries)
                    {
                        cmd.Parameters["@comment"].Value = logEntry.Comment ?? (object)DBNull.Value;
                        cmd.Parameters["@iteration"].Value = logEntry.Iteration ?? (object)DBNull.Value;
                        cmd.Parameters["@memory"].Value = logEntry.Memory;
                        cmd.Parameters["@cpu"].Value = logEntry.Cpu;
                        cmd.Parameters["@time"].Value = logEntry.Time;
                        cmd.Parameters["@test_run_id"].Value = testRunId;

                        cmd.ExecuteNonQuery();
                    }
                }

                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }
        }

        private static int FindOrCreateMachineAndGetId()
        {
            var machineName = System.Environment.MachineName;

            var machineId = -1;

            using (var connection = new SqlConnection(PerformanceDbConnectionString))
            {
                connection.Open();

                const string machineLookup = "select machine_id from dbo.MACHINE where machine_name = @machine_name " +
                                             "and version_number = " +
                                             "  (select max(version_number) from dbo.MACHINE where machine_name=@machine_name) " +
                                             "and cores=@cores and operating_system=@operating_system and office_version=@office_version";

                using (var cmd = new SqlCommand(machineLookup, connection))
                {
                    cmd.Parameters.AddWithValue("@machine_name", machineName);
                    cmd.Parameters.AddWithValue("@cores", Convert.ToInt32(Environment[ExecutingMachineInfo.LogicalProcessors]));
                    cmd.Parameters.AddWithValue("@operating_system", Environment[ExecutingMachineInfo.OperatingSystem]);
                    cmd.Parameters.AddWithValue("@office_version", Environment[ExecutingMachineInfo.OfficeVersion]);

                    var result = cmd.ExecuteScalar();

                    if (result != null)
                    {
                        machineId = (int)result;
                    }

                    if (connection.State == ConnectionState.Open)
                        connection.Close();
                }
            }

            if (machineId < 0)
            {
                const string machineInsert =
                    "Insert into MACHINE (machine_name, version_number, cpu_speed, cores, memory, " +
                    "operating_system, office_version) " +
                    "output INSERTED.machine_id " +
                    "values (@machine_name, (select (ISNULL(MAX(version_number), 0) + 1) from MACHINE where machine_name = " +
                    "@machine_name), @cpu_speed, @cores, @memory, @operating_system, @office_version)";

                using (var connection = new SqlConnection(PerformanceDbConnectionString))
                {
                    connection.Open();

                    using (var cmd = new SqlCommand(machineInsert, connection))
                    {
                        cmd.Parameters.AddWithValue("@machine_name", machineName);
                        cmd.Parameters.AddWithValue("@cpu_speed", Convert.ToSingle(Environment[ExecutingMachineInfo.CpuSpeed]));
                        cmd.Parameters.AddWithValue("@cores", Convert.ToInt32(Environment[ExecutingMachineInfo.LogicalProcessors]));
                        cmd.Parameters.AddWithValue("@memory", Convert.ToInt32(Environment[ExecutingMachineInfo.TotalMemory]));
                        cmd.Parameters.AddWithValue("@operating_system", Environment[ExecutingMachineInfo.OperatingSystem]);
                        cmd.Parameters.AddWithValue("@office_version", Environment[ExecutingMachineInfo.OfficeVersion]);
                        var result = cmd.ExecuteScalar();

                        if (result != null)
                        {
                            machineId = (int)result;
                        }

                        if (connection.State == ConnectionState.Open)
                            connection.Close();
                    }
                }
            }

            if (machineId < 0) throw new Exception("Unable to retrieve test definition ID from database.");
            return machineId;
        }

        private static int FindOrCreateTestDefinitionAndGetId(PerformanceTestRun testRun)
        {
            const string testDefinitionSelect =
                "Select test_definition_id from TEST_DEFINITION where file_identifier = @identifier";

            var testDefinitionId = -1;

            using (var connection = new SqlConnection(PerformanceDbConnectionString))
            {
                connection.Open();

                using (var cmd = new SqlCommand(testDefinitionSelect, connection))
                {
                    cmd.Parameters.AddWithValue("@identifier", testRun.Id);
                    var result = cmd.ExecuteScalar();

                    if (result != null)
                    {
                        testDefinitionId = (int)result;
                    }

                    if (connection.State == ConnectionState.Open)
                        connection.Close();
                }
            }

            if (testDefinitionId < 0)
            {
                const string testDefinitionInsert =
                    "Insert into TEST_DEFINITION ([file_identifier], [title], [steps]) " +
                    "output INSERTED.test_definition_id " +
                    "values (@identifier, @title, @steps)";

                var resourceManager = new ResourceManager(typeof(Resources));
                var steps = (string)resourceManager.GetObject(testRun.Id);

                using (var connection = new SqlConnection(PerformanceDbConnectionString))
                {
                    connection.Open();

                    using (var cmd = new SqlCommand(testDefinitionInsert, connection))
                    {
                        cmd.Parameters.AddWithValue("@identifier", testRun.Id);
                        cmd.Parameters.AddWithValue("@title", testRun.TestTitle);
                        cmd.Parameters.AddWithValue("@steps", string.Join(System.Environment.NewLine, steps));
                        var result = cmd.ExecuteScalar();

                        if (result != null)
                        {
                            testDefinitionId = (int)result;
                        }

                        if (connection.State == ConnectionState.Open)
                            connection.Close();
                    }
                }
            }

            if (testDefinitionId < 0) throw new Exception("Unable to retrieve test definition ID from database.");
            return testDefinitionId;
        }
    }
}