using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CommonLib.Controls
{
    public partial class ComparisonOTPRow : UserControl
    {
        public event EventHandler DelButtonClick;

        public bool IsFirstRow {
            set {
                chkVC.Visible = !value;
            }
            get {
                return !chkVC.Visible;
            }
        }

        public string FileName
        {
            set {
                txtFile.Text = value;
                txtFile.Select(txtFile.TextLength, 0);
            }
            get {
                return txtFile.Text;
            }
        }

        public bool HCChecked { 
            set {
                chkHC.Checked = value;
            }
            get {
                return chkHC.Checked;
            }
        }

        public bool VCChecked
        {
            set
            {
                chkVC.Checked = value;
            }
            get
            {
                return chkVC.Checked;
            }
        }

        public ComparisonOTPRow()
        {
            InitializeComponent();
        }

        private void btndel_Click(object sender, EventArgs e)
        {
            DelButtonClick?.Invoke(this, e);
        }
    }
}
