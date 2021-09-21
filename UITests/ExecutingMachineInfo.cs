using System;
using System.Collections;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
using LibGit2Sharp;
using Microsoft.Win32;
using Microsoft.VisualBasic.Devices;

namespace UITests
{
    public class ExecutingMachineInfo
    {
        private static readonly RegistryKey OfficeKey =
            Registry.ClassesRoot.OpenSubKey(Constants.OfficeRegistryKey);
        private static readonly RegistryKey CpuSpeedKey =
            Registry.LocalMachine.OpenSubKey(Constants.FirstCoreInfoRegistryKey);
        private static readonly RegistryKey CpuCoresKey =
            Registry.LocalMachine.OpenSubKey(Constants.CpuInfoRegistryKey);

        public override string ToString()
        {
            var sb = new StringBuilder();

            foreach (DictionaryEntry entry in AsOrderedDictionary())
            {
                sb.Append(entry.Key)
                  .Append(',')
                  .Append(entry.Value)
                  .AppendLine();
            }

            return sb.ToString();
        }

        public static OrderedDictionary AsOrderedDictionary()
        {
            var computerInfo = new ComputerInfo();

            var dictionary = new OrderedDictionary
            {
                {BranchName, $"{GetCurrentCodeReference()}"},
                {CpuSpeed, $"{((int?) CpuSpeedKey?.GetValue("~MHz") ?? 0) / 1000.0}"},
                {LogicalProcessors, CpuCoresKey?.GetSubKeyNames().Length.ToString()},
                {TotalMemory, $"{(int) (computerInfo.TotalPhysicalMemory / Math.Pow(1024, 3))}"},
                {OperatingSystem, computerInfo.OSFullName},
                {OsVersion, computerInfo.OSVersion},
                {OfficeVersion, OfficeKey?.GetValue(null).ToString().Split('.')[2]},
                {ReloadIterations, Constants.ReloadIterations},
                {NumInspectorWindows, Constants.InspectorWindowCount},
                {NumEmails, Constants.MassEmailCount},
                {WarmupTime, $"{Constants.WarmupTimeMilliseconds / 1000}"},
                {CooldownTime, $"{Constants.CooldownTimeMilliseconds / 1000}"}
            };

            return dictionary;
        }

        public const string BranchName = "Branch Name";
        public const string CpuSpeed = "CPU Speed (GHz)";
        public const string LogicalProcessors = "Logical Processors";
        public const string TotalMemory = "Total Physical Memory (GB)";
        public const string OperatingSystem = "Operating System";
        public const string OsVersion = "Operating System Version";
        public const string OfficeVersion = "Office Version";
        public const string ReloadIterations = "Reload Iterations";
        public const string NumInspectorWindows = "Number of Inspector Windows";
        public const string NumEmails = "Number of Emails";
        public const string WarmupTime = "Warmup Time (seconds)";
        public const string CooldownTime = "Cooldown Time (seconds)";

        private static string GetCurrentCodeReference()
        {
            var currentPath = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            var found = false;
            var result = "Not available";

            while (!found)
            {
                if (currentPath == null) return result;

                if (currentPath.GetDirectories(".git").Length > 0)
                {
                    found = true;
                }
                else
                {
                    currentPath = currentPath.Parent;
                }
            }

            using (var repo = new Repository(currentPath.FullName))
            {
                var commit = repo.Head.Tip;
                var tagName = repo.Tags.FirstOrDefault(t => t?.Target?.Sha == commit?.Sha);

                result = tagName == null ? repo.Head.FriendlyName : tagName.CanonicalName.Split('/').Last();
            }

            return result;
        }
    }
}