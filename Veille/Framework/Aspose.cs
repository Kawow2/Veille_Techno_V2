using Aspose.Cells;
using Aspose.Cells.Charts;
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

        public ResultAnalysis CreateChart()
        {
            var analysis = new ResultAnalysis();
            OpenFile("..\\..\\FileExample\\dataforchart.xlsx");
            Timer.Start();
            var chartIndex = this.Workbook.Worksheets[0].Charts.Add(ChartType.Column, 23, 15, 39, 24);
            var chart = Workbook.Worksheets[0].Charts[chartIndex]; 
            chart.SetChartDataRange("A1:B100", true);
            chart.Legend.Position = LegendPositionType.Bottom;
            Timer.Stop();
            analysis.TimeInMs = Timer.GetTime();
            analysis.CPUUsage = PerformanceAnalysis.GetCurrentCpuUsage();

            WriteFile("..\\..\\FileExample\\chart_aspose.xlsx", SaveFormat.Auto);
            return analysis;
        }


    }
}
