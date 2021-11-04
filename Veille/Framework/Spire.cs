using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Veille.Tools;

namespace Veille.Framework
{
    public class Spire
    {
        public Workbook wb { get; set; } 
        public Spire()
        {
            wb = new Workbook(); ;
        }
        public ResultAnalysis OpenFile(string fileName)
        {
            var analysis = new ResultAnalysis();
            var ext = fileName.Split('.')[1];
            if (ext == "csv")
            {
                Timer.Start();
                wb.LoadFromFile(fileName,",");
                Timer.Stop();
                analysis.CPUUsage = PerformanceAnalysis.GetCurrentCpuUsage();
                analysis.TimeInMs = Timer.GetTime();
            }
            else
            {
                Timer.Start();
                wb.LoadFromFile(fileName);
                Timer.Stop();
                analysis.CPUUsage = PerformanceAnalysis.GetCurrentCpuUsage();
                analysis.TimeInMs = Timer.GetTime();
            }
            analysis.TimeInMs = Timer.GetTime();
            return analysis;
        }
        public ResultAnalysis WriteFile(string filename)
        {
            var analysis = new ResultAnalysis();
            Timer.Start();
            this.wb.SaveToFile(filename);
            Timer.Stop();
            analysis.CPUUsage = PerformanceAnalysis.GetCurrentCpuUsage();
            analysis.TimeInMs = Timer.GetTime();
            return analysis;
        }
    }
}
