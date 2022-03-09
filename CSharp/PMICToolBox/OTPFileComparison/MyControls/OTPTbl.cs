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
    public partial class ComparisonOTPTbl : UserControl
    {
        public List<OTPFileInfo> OTPFileInfoList
        { 
            get{
                List<OTPFileInfo> oTPFileInfos = new List<OTPFileInfo>();
                for (int i = 0; i < this.flowLayoutPanelTbl.Controls.Count; i++)
                {
                    ComparisonOTPRow comparisonOTPRow1 = (ComparisonOTPRow)this.flowLayoutPanelTbl.Controls[i];
                    oTPFileInfos.Add(
                        new OTPFileInfo() { 
                            FileName = comparisonOTPRow1.FileName, 
                            HorizontalComparison = comparisonOTPRow1.HCChecked, 
                            VerticalComparison = comparisonOTPRow1.VCChecked });
                }

                return oTPFileInfos;
            }
        }

        public ComparisonOTPTbl()
        {
            InitializeComponent();
        }

        public void AddNewFile(string file)
        {
            if (DuplicateCheck(file) == false)
            {
                MessageBox.Show("this file has been existed in list", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ComparisonOTPRow comparisonOTPRow1 = new ComparisonOTPRow();
            comparisonOTPRow1.DelButtonClick += ComparisonOTPRow_DelButtonClick;
            comparisonOTPRow1.FileName = file;
            comparisonOTPRow1.Height = 24;
            if (this.flowLayoutPanelTbl.Controls.Count == 0)
            {
                comparisonOTPRow1.IsFirstRow = true;
            }
            this.flowLayoutPanelTbl.Controls.Add(comparisonOTPRow1);
        }

        private bool DuplicateCheck(string file)
        {
            for (int i = 0; i < this.flowLayoutPanelTbl.Controls.Count; i++)
            {
                ComparisonOTPRow comparisonOTPRow1 = (ComparisonOTPRow)this.flowLayoutPanelTbl.Controls[i];
                if (comparisonOTPRow1.FileName == file)
                {
                    return false;
                }
            }

            return true;
        }

        private void ComparisonOTPRow_DelButtonClick(object sender, EventArgs e)
        {
            DelItem((ComparisonOTPRow)sender);
        }

        public void DelItem(ComparisonOTPRow comparisonOTPRow)
        {
            this.flowLayoutPanelTbl.Controls.Remove(comparisonOTPRow);
            if (this.flowLayoutPanelTbl.Controls.Count > 0)
            {
                ((ComparisonOTPRow)this.flowLayoutPanelTbl.Controls[0]).IsFirstRow = true;
            }
        }

        private void ComparisonOTPTbl_Load(object sender, EventArgs e)
        {
        }
    }

    public class OTPFileInfo
    {
        public string FileName;
        public bool HorizontalComparison;
        public bool VerticalComparison;
    }
}
