using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Charts;
using GemBox.Spreadsheet.PivotTables;
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
        }
        public ResultAnalysis OpenFile(string filename, bool analyseTime = true)
        {
            ef = ef ?? new ExcelFile();

            var analysis = new ResultAnalysis();

            if (analyseTime)
                Timer.Start();
            this.ef = ExcelFile.Load(filename);

            if(analyseTime)
            {
                Timer.Stop();
                analysis.TimeInMs = Timer.GetTime();
            }
            analysis.CPUUsage = PerformanceAnalysis.GetCurrentCpuUsage();
           
            return analysis;
        }

        public ResultAnalysis WriteFile(string filename,bool analyseTime = true)
        {
            ef = ef ?? new ExcelFile();

            var analysis = new ResultAnalysis();
            if (analyseTime)
                Timer.Start();
            ef.Save(filename);
            if(analyseTime)
            {
                Timer.Stop();   
                analysis.TimeInMs = Timer.GetTime();
            }
            analysis.CPUUsage = PerformanceAnalysis.GetCurrentCpuUsage();
            return analysis;
        }

        public ResultAnalysis CreateChart()
        {
            ef = ef ?? new ExcelFile();

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

        public ResultAnalysis CreatePivotTable()
        {
            ef = ef ?? new ExcelFile();

            var analysis = new ResultAnalysis();
            Timer.Start();
            var workbook = CreateDataForPivotTable();
            workbook.Save("..\\..\\FileExample\\testgembox.xlsx");

            var source = workbook.Worksheets.ToList();
            var cache = workbook.PivotCaches.AddWorksheetSource("SourceSheet!A1:D100");

            // Create new sheet for pivot table.
            var worksheet2 = workbook.Worksheets.Add("PivotSheet");

            // Create pivot table "Company Profile" using the specified pivot cache and add it to the worksheet at the cell location 'A1'.
            var table = worksheet2.PivotTables.Add(cache, "Company Profile", "A1");

            // Aggregate 'Names' values into count value and show it as a percentage of row.
            var field = table.DataFields.Add("Names");
            field.Function = PivotFieldCalculationType.Count;
            field.ShowDataAs = PivotFieldDisplayFormat.PercentageOfRow;
            field.Name = "% of Empl.";

            // Aggregate 'Salaries' values into average value.
            field = table.DataFields.Add("Salaries");
            field.Function = PivotFieldCalculationType.Average;
            field.Name = "Avg. Salary";
            field.NumberFormat = "[$$-409]#,##0.00";

            // Group rows into 'Departments'.
            table.RowFields.Add("Departments");

            // Group columns first into 'Years of Service' and then into 'Values' (count 'Names' and average 'Salaries').
            table.ColumnFields.Add("Years of Service");
            table.ColumnFields.Add(table.DataPivotField);

            // Specify the string to be displayed in row and column header.
            table.RowHeaderCaption = "Departments";
            table.ColumnHeaderCaption = "Years of Service";

            // Do not show grand totals for rows.
            table.RowGrandTotals = false;

            // Set pivot table style.
            table.BuiltInStyle = BuiltInPivotStyleName.PivotStyleMedium7;
            Timer.Stop();
            workbook.Save("..\\..\\FileExample\\gembox_pivottable.xlsx");
            analysis.CPUUsage = PerformanceAnalysis.GetCurrentCpuUsage();
            analysis.TimeInMs = Timer.GetTime();
            return analysis;

        }

        public ExcelFile CreateDataForPivotTable()
        {
            var wb = new ExcelFile();
            var w1 = wb.Worksheets.Add("SourceSheet");

            var cells = w1.Cells;
            cells[0, 0].Value = "Departments";
            cells[0, 1].Value = "Names";
            cells[0, 2].Value = "Years of Service";
            cells[0, 3].Value = "Salaries";
            var random = new Random();
            var departments = new string[] { "Legal", "Marketing", "Finance", "Planning", "Purchasing" };
            var names = new string[] { "John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat" };
            var years = new string[] { "1-10", "11-20", "21-30", "over 30" };
            for (int i = 0; i < 100; ++i)
            {
                cells[i + 1, 0].Value = departments[random.Next(departments.Length)];
                cells[i + 1, 1].Value = names[random.Next(names.Length)] + ' ' + (i + 1).ToString();
                cells[i + 1, 2].Value = years[random.Next(years.Length)];
                cells[i + 1, 3].SetValue(random.Next(10, 101) * 100);
            }
            wb.Save("..\\..\\FileExample\\testaspose.xlsx");
            return wb;
        }
    }
}
