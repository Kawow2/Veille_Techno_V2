using System.Text;

namespace Veille.Tools
{
    public class ResultAnalysis
    {
        public string TimeInMs { get; set;}
        public string CPUUsage { get; set;}
        public string CPUPeak { get; set;}
        public string RAMUsage { get; set;}

        
        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append("Time in ms : " + TimeInMs + "\n"); 
            return sb.ToString();
        }
    }
}
