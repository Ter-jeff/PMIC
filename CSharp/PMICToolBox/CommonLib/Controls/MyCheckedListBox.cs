using System.Windows.Forms;

namespace CommonLib.Controls
{
    public partial class MyCheckededListBox : CheckedListBox
    {
        public MyCheckededListBox()
        {
            InitializeComponent();
        }

        public override int ItemHeight { get { return Font.Height + 4; } set { } }
    }
}
