using GemBox.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Veille.Tools;

namespace BIGEARD_VEILLE_TECHNOLOGIQUE.Framework
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
        public string OpenFile(string filename)
        {
            Timer.Start();
            this.ef = ExcelFile.Load(filename);
            Timer.Stop();
            return Timer.GetTime();
        }

        public string WriteFile(string filename)
        {
            Timer.Start();
            this.ef.Save(filename);
            Timer.Stop();
            return Timer.GetTime();
        }


    }
}
