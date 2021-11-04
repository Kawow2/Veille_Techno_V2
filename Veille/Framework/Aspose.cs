using Aspose.Cells;
using System;
using System.IO;
using System.Linq;
using System.Windows;
using Veille.Tools;

namespace Veille.Framework
{
    public class Aspose
    {
        public Workbook Workbook { get; set; } = new Workbook();
        public ResultAnalysis OpenFile(string filename, LoadOptions lo = null)
        {
            var analysis = new ResultAnalysis();
            try
            {
                if (this.Workbook != null)
                {
                    this.Workbook.Dispose();
                
                }
                Timer.Start();
                var fstream = new FileStream(filename, FileMode.Open);
                this.Workbook = new Workbook(fstream, lo);
                Timer.Stop();
                analysis.CPUUsage = PerformanceAnalysis.GetCurrentCpuUsage();
                analysis.TimeInMs = Timer.GetTime();
            } catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            return analysis;
        }
        public ResultAnalysis WriteFile(string fileName,SaveFormat ext)
        {
            var analysis = new ResultAnalysis();
            Timer.Start();
            this.Workbook.Save(fileName,ext);
            Timer.Stop();
            analysis.CPUUsage = PerformanceAnalysis.GetCurrentCpuUsage(); 
            analysis.TimeInMs = Timer.GetTime();
            return analysis;
        }
    }
}
