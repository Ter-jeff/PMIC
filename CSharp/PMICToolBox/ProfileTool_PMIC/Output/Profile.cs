using System.Collections.Generic;

namespace ProfileTool_PMIC.Output
{
    public class Profile
    {
        #region Properity
        public string FilePath { get; set; }
        public string Item { get; set; }//FLow or instance
        public string Link
        {
            get
            {
                if (string.IsNullOrEmpty(HyperLink)) return "";
                return "=HYPERLINK(\"" + HyperLink + "\",\"Link\")";
            }
        }

        public int Site { get; set; }
        public string Pin { get; set; }
        public double SampleRate { get; set; }
        public double SampleSize { get; set; }
        public string Date { get; set; }
        public double MaxBeforeFilter { get; set; }
        public double MinBeforeFilter { get; set; }
        public double CountBeforeFilter { get; set; }
        public double MaxAfterFilter { get; set; }
        public double MinAfterFilter { get; set; }
        public double CountAfterFilter { get; set; }
        public string ChartType { get; set; }
        public List<double> Value;
        public int MaxIndex;
        public string HyperLink;
        #endregion

        #region Constructor
        public Profile()
        {
            Item = "";
            Pin = "";
            Date = "";
            Value = new List<double>();
        }
        #endregion
    }
}