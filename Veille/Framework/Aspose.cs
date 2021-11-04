using Aspose.Cells;
using System;
using System.IO;
using System.Linq;
using System.Windows;
using Veille.Tools;

namespace BIGEARD_VEILLE_TECHNOLOGIQUE.Framework
{
    public class Aspose
    {
        public Workbook Workbook { get; set; } = new Workbook();
        public string OpenFile(string filename, LoadOptions lo = null)
        {            
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
            } catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            return Timer.GetTime();
        }
        public string WriteFile(string fileName,SaveFormat ext)
        {
            Timer.Start();
            this.Workbook.Save(fileName,ext);
            Timer.Stop();
            return Timer.GetTime();
        }
    }
}
