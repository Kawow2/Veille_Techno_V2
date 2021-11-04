using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Veille.Tools
{
    public class Timer
    {
        private static Stopwatch clock { get; set; } = new Stopwatch();

        public static void Start()
        {
            clock.Reset();
            clock.Start();
        }

        public static void Stop()
        {
            clock.Stop();
        }

        public static string GetTime()
        {
            return clock.ElapsedMilliseconds.ToString() + " ms";
        }
    }
}
