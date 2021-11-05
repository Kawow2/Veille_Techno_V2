using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Veille.Tools;

namespace Veille.Framework
{
    public class Gembox
    {

        public ExcelFile ef { get; set; } 
        public Gembox()
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            SpreadsheetInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;
            this.ef = new ExcelFile();
        }
        public ResultAnalysis OpenFile(string filename)
        {
            var analysis = new ResultAnalysis();

            Timer.Start();
            this.ef = ExcelFile.Load(filename);
            Timer.Stop();
            analysis.TimeInMs = Timer.GetTime();
            analysis.CPUUsage = PerformanceAnalysis.GetCurrentCpuUsage();
           
            return analysis;
        }

        public ResultAnalysis WriteFile(string filename)
        {
            var analysis = new ResultAnalysis();

            Timer.Start();
            this.ef.Save(filename);
            Timer.Stop();
            analysis.TimeInMs = Timer.GetTime();
            analysis.CPUUsage = PerformanceAnalysis.GetCurrentCpuUsage();
            return analysis;
        }

        public ResultAnalysis CreateChart()
        {
            var analysis = new ResultAnalysis();
            OpenFile("..\\..\\FileExample\\dataforchart.xlsx");
            Timer.Start();
            var ws = ef.Worksheets.FirstOrDefault();
            var chart = ws.Charts.Add<ColumnChart>(ChartGrouping.Standard, "A1", "B100");
            chart.SelectData(ws.Cells.GetSubrangeAbsolute(0, 0, 100, 1));
            Timer.Stop();
            analysis.CPUUsage = PerformanceAnalysis.GetCurrentCpuUsage();
            analysis.TimeInMs = Timer.GetTime();
            WriteFile("..\\..\\FileExample\\chart_gembox.xlsx");
            return analysis;
        }


    }
}
