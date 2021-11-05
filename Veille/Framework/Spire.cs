using Spire.Xls;
using System;
using System.Collections.Generic;
using System.IO;
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
        public ResultAnalysis OpenFile(string fileName, ExcelVersion excelVersion = ExcelVersion.Version2016, bool analyseTime = true)
        {
            var analysis = new ResultAnalysis();
            var ext = fileName.Split('.')[1];
            if (ext == "csv")
            {
                if (analyseTime)
                    Timer.Start();
                wb.LoadFromFile(fileName,",");
                if (analyseTime)
                {
                    Timer.Stop();
                    analysis.TimeInMs = Timer.GetTime();

                }
                analysis.CPUUsage = PerformanceAnalysis.GetCurrentCpuUsage();
            }
            else
            {
                if (analyseTime)
                    Timer.Start();
                wb.LoadFromFile(fileName,excelVersion);
                if (analyseTime)
                {
                    Timer.Stop();
                    analysis.TimeInMs = Timer.GetTime();

                }
                analysis.CPUUsage = PerformanceAnalysis.GetCurrentCpuUsage();
                analysis.TimeInMs = Timer.GetTime();
            }
            analysis.TimeInMs = Timer.GetTime();
            return analysis;
        }

        public ResultAnalysis WriteFile(string filename, bool analyseTime = true)
        {
            var analysis = new ResultAnalysis();
            if (analyseTime)
                Timer.Start();
            this.wb.SaveToFile(filename);
            if (analyseTime)
            {
                Timer.Stop();
                analysis.TimeInMs = Timer.GetTime();
            }
            analysis.CPUUsage = PerformanceAnalysis.GetCurrentCpuUsage();
            return analysis;
        }

        public ResultAnalysis CreateChart()
        {
            var analysis = new ResultAnalysis();
            Timer.Start();
            OpenFile("..\\..\\FileExample\\dataforchart.xlsx",analyseTime:false);
            var chart = wb.Worksheets[0].Charts.Add(ExcelChartType.ColumnClustered);
            chart.DataRange = wb.Worksheets[0].Range["A1:B100"];
            WriteFile("..\\..\\FileExample\\chart_spire.xlsx",analyseTime: false);
            Timer.Stop();
            analysis.CPUUsage = PerformanceAnalysis.GetCurrentCpuUsage();
            analysis.TimeInMs = Timer.GetTime();

            return analysis;
        }

        public ResultAnalysis  CreatePivotTable()
        {
            var analysis = new ResultAnalysis();
            Timer.Start();
            OpenFile("..\\..\\FileExample\\template.xls",ExcelVersion.Version97to2003, analyseTime: false);
            var cells = wb.Worksheets["data"];
            //cells[0, 0].Value = "Departments";
            //cells[0, 1].Value = "Names";
            //cells[0, 2].Value = "Years of Service";
            //cells[0, 3].Value = "Salaries";
            var random = new Random();
            var departments = new string[] { "Legal", "Marketing", "Finance", "Planning", "Purchasing" };
            var names = new string[] { "John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat" };
            var years = new string[] { "1-10", "11-20", "21-30", "over 30" };
            for (int i = 1; i < 101; ++i)
            {
                cells[i + 1, 1].Value = departments[random.Next(departments.Length)];
                cells[i + 1, 2].Value = names[random.Next(names.Length)] + ' ' + (i + 1).ToString();
                cells[i + 1, 3].Value = years[random.Next(years.Length)];
                cells[i + 1, 4].Value = (random.Next(10, 101) * 100).ToString();
            }
            var pt = wb.Worksheets[0].PivotTables["TCD"];
            //wb.SaveToFile("..\\..\\FileExample\\spire_pivottable.xlsx");
            WriteFile("..\\..\\FileExample\\spire_pivottable.xlsx", analyseTime: false);
            Timer.Stop();
            analysis.TimeInMs = Timer.GetTime();
            analysis.CPUUsage = PerformanceAnalysis.GetCurrentCpuUsage();
            return analysis;
        }

        //private Workbook CreateDa
    }
}
