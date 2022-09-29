using System.Drawing;

namespace ShmooLog.Base
{
    public class ColorSetting
    {
        public ColorSetting()
        {
            PPColor = Color.LimeGreen;
            PFColor = Color.Yellow;
            FPColor = Color.Orange;
            FFColor = Color.Red;
        }

        public Color PPColor { get; set; }
        public Color PFColor { get; set; }
        public Color FFColor { get; set; }
        public Color FPColor { get; set; }
    }
}