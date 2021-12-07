using System.Diagnostics;
using System.Threading;

namespace Veille.Tools
{
    public static class PerformanceAnalysis
    {
        public static PerformanceCounter CpuCounter { get; set; } = new PerformanceCounter("Processor", "% Processor Time", "_Total");
        public static PerformanceCounter RamCounter { get; set; } = new PerformanceCounter("Memory", "Available MBytes");
      
        public static string GetCurrentCpuUsage()
        {
            //var temp = CpuCounter.NextValue();
            //Thread.Sleep(100);
            return CpuCounter.NextValue() + "%";
            //return "";
        }

        public static string GetAvailableRAM()
        {
            return RamCounter.NextValue() + "MB";
        }


    }
}
