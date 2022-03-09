using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace AutomationCommon.Controls
{
    public class MyGroupBox : GroupBox
    {
        [Description("設定或取得外框顏色")]
        public Color BorderColor { get; set; } = SystemColors.Window;

        protected override void OnPaint(PaintEventArgs e)
        {
            //取得text字型大小
            Size fontSize = TextRenderer.MeasureText(Text,
                Font);
            //畫框線
            Rectangle rec = new Rectangle(e.ClipRectangle.Y,
                Font.Height / 2,
                e.ClipRectangle.Width - 1,
                e.ClipRectangle.Height - 1 -
                Font.Height / 2);

            e.Graphics.DrawRectangle(new Pen(BorderColor), rec);

            //填滿text的背景
            e.Graphics.FillRectangle(new SolidBrush(BackColor),
                new Rectangle(6, 0, fontSize.Width, fontSize.Height));

            //text
            e.Graphics.DrawString(Text, Font,
                new Pen(ForeColor).Brush, 6, 0);
        }
    }
}