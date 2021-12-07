using Aspose.Cells;
using Aspose;
using Aspose.Cells.Charts;
using System;
using System.IO;
using System.Linq;
using System.Windows;
using Veille.Tools;
using Aspose.Cells.Pivot;

namespace Veille.Framework
{
    public class Aspose
    {
        public Workbook Workbook { get; set; }
        public ResultAnalysis OpenFile(string filename, LoadOptions lo = null, bool analyseTime = true)
        {
            var analysis = new ResultAnalysis();
            try
            {
                this.Workbook?.Dispose();
                if (analyseTime)
                    Timer.Start();
                using (var fstream = new FileStream(filename, FileMode.Open))
                {
                    this.Workbook = new Workbook(fstream, lo);
                }

                //mettre des using 

                if (analyseTime)
                {
                    Timer.Stop();
                    analysis.TimeInMs = Timer.GetTime();   
                }
                analysis.CPUUsage = PerformanceAnalysis.GetCurrentCpuUsage();
            } catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            return analysis;
        }
        public ResultAnalysis WriteFile(string fileName,SaveFormat ext, bool analyseTime = true)
        {
            var analysis = new ResultAnalysis();
            if (analyseTime)
                Timer.Start();
            this.Workbook.Save(fileName,ext);
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
            Workbook = Workbook ?? new Workbook();
            var analysis = new ResultAnalysis();
            OpenFile("..\\..\\FileExample\\dataforchart.xlsx",analyseTime:false);
            Timer.Start();
            var chartIndex = this.Workbook.Worksheets[0].Charts.Add(ChartType.Column, 23, 15, 39, 24);
            var chart = Workbook.Worksheets[0].Charts[chartIndex]; 
            chart.SetChartDataRange("A1:B100", true);
            chart.Legend.Position = LegendPositionType.Bottom;
            Timer.Stop();
            analysis.TimeInMs = Timer.GetTime();
            analysis.CPUUsage = PerformanceAnalysis.GetCurrentCpuUsage();

            WriteFile("..\\..\\FileExample\\chart_aspose.xlsx", SaveFormat.Auto,false);
            return analysis;
        }

        public ResultAnalysis CreatePivotTable()
        {
            var analysis = new ResultAnalysis();
            Timer.Start();
            Workbook = CreateDataForPivotTable2();
            var sheet = Workbook.Worksheets[Workbook.Worksheets.Add()];
            sheet.Name = "PivotTable";
            //sheet.Cells[]
            var pivotTables = sheet.PivotTables;
            var index = pivotTables.Add("Data!A1:D101", "B3", "PivotTable1");
            var pivotTable = pivotTables[index];
            pivotTable.RowGrand = true;
            pivotTable.ColumnGrand = true;
            pivotTable.IsAutoFormat = true;
            //department
            //names
            //years
            //salaries 
            pivotTable.AutoFormatType = PivotTableAutoFormatType.Report6;
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            pivotTable.AddFieldToArea(PivotFieldType.Column, 2);
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);
            pivotTable.AddFieldToArea(PivotFieldType.Data, 3);

            pivotTable.DataFields[0].Function = ConsolidationFunction.Count;
            pivotTable.DataFields[0].DataDisplayFormat = PivotFieldDataDisplayFormat.PercentageOfRow;

            pivotTable.DataFields[1].Function = ConsolidationFunction.Average;
            pivotTable.DataFields[1].NumberFormat = "[$$-409]#,##0.00"; ;
            pivotTable.DataFields[1].DisplayName = "Avg. Salary";

            pivotTable.ColumnFields.Add(pivotTable.DataField);
            pivotTable.ColumnFields[0].Function = ConsolidationFunction.Sum;
            //pivotTable.DataFields[2].Function = ConsolidationFunction.Pe;
            //pivotTable.AddFieldToArea(PivotFieldType.Data, 5);

            pivotTable.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium7;
            Timer.Stop();
            analysis.TimeInMs = Timer.GetTime();
            analysis.CPUUsage = PerformanceAnalysis.GetCurrentCpuUsage();
            WriteFile("..\\..\\FileExample\\pivottable_aspose.xlsx",SaveFormat.Auto,false);
            

            return analysis;
        }

    
        private Workbook CreateDataForPivotTable2()
        {
            var wb = new Workbook();
            var ws = wb.Worksheets.Add("Data");
            var cells = ws.Cells;
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
                cells[i + 1, 3].PutValue(random.Next(10, 101) * 100);
            }
            return wb;
        }
    }
}
