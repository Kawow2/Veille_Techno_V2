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
            var cs1 = CpuCounter.NextSample();
            System.Threading.Thread.Sleep(100);
            var cs2 = CpuCounter.NextSample();
            var finalCpuCounter = CounterSample.Calculate(cs1, cs2);
            return finalCpuCounter + " %";
        }

        public static string GetAvailableRAM()
        {
            return RamCounter.NextValue() + "MB";
        }


    }
}
