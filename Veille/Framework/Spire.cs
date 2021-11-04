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
        public string OpenFile(string fileName)
        {
            var ext = fileName.Split('.')[1];
            if (ext == "csv")
            {
                Timer.Start();
                wb.LoadFromFile(fileName,",");
                Timer.Stop();
            }
            else
            {
                Timer.Start();
                wb.LoadFromFile(fileName);
                Timer.Stop();
            }            
            return Timer.GetTime();
        }
        public string WriteFile(string filename)
        {
            Timer.Start();
            this.wb.SaveToFile(filename);
            Timer.Stop();
            return Timer.GetTime();
        }
    }
}
