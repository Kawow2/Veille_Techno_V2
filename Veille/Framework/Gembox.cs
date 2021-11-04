using GemBox.Spreadsheet;
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


    }
}
