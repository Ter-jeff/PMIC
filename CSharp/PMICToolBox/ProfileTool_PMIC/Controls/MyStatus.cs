using System.Drawing;
using System.Windows.Forms;

namespace AutomationCommon.Controls
{
    public sealed class MyStatus : StatusStrip
    {
        private StatusStrip _statusStrip1;

        public MyStatus()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            _statusStrip1 = new StatusStrip();
            ToolStripStatusLabel = new ToolStripStatusLabel();
            ToolStripStatusLabel1 = new ToolStripStatusLabel();
            ProcessTimeToolStripStatusLabel = new ToolStripStatusLabel();
            ToolStripProgressBar = new ToolStripProgressBar();
            _statusStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // statusStrip1
            // 
            this.Items.AddRange(new ToolStripItem[] {
            ToolStripStatusLabel,
            ToolStripStatusLabel1,
            ProcessTimeToolStripStatusLabel,
            ToolStripProgressBar});
            _statusStrip1.Location = new Point(0, 645);
            _statusStrip1.Name = "_statusStrip1";
            _statusStrip1.Size = new Size(740, 22);
            _statusStrip1.TabIndex = 8;
            _statusStrip1.Text = "statusStrip1";
            // 
            // ToolStripStatusLabel
            // 
            ToolStripStatusLabel.Name = "ToolStripStatusLabel";
            ToolStripStatusLabel.Size = new Size(39, 17);
            ToolStripStatusLabel.Text = "Status";
            // 
            // toolStripStatusLabel1
            // 
            ToolStripStatusLabel1.Name = "ToolStripStatusLabel1";
            ToolStripStatusLabel1.Size = new Size(507, 17);
            ToolStripStatusLabel1.Spring = true;
            // 
            // ProcessTime_toolStripStatusLabel
            // 
            ProcessTimeToolStripStatusLabel.Name = "ProcessTimeToolStripStatusLabel";
            ProcessTimeToolStripStatusLabel.Size = new Size(77, 17);
            ProcessTimeToolStripStatusLabel.Text = "Process Time";
            ProcessTimeToolStripStatusLabel.TextAlign = ContentAlignment.MiddleRight;
            // 
            // toolStripProgressBar
            // 
            ToolStripProgressBar.Alignment = ToolStripItemAlignment.Right;
            ToolStripProgressBar.Maximum = 10;
            ToolStripProgressBar.Name = "ToolStripProgressBar";
            ToolStripProgressBar.Size = new Size(150, 16);
            // 
            // PatListCheckForm
            // 
            ClientSize = new Size(740, 667);
            //Controls.Add(_statusStrip1);
            Font = new Font("Microsoft Sans Serif", 9F);
            _statusStrip1.ResumeLayout(false);
            _statusStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        public ToolStripStatusLabel ToolStripStatusLabel;
        public ToolStripProgressBar ToolStripProgressBar;
        public ToolStripStatusLabel ProcessTimeToolStripStatusLabel;
        public ToolStripStatusLabel ToolStripStatusLabel1;
    }
}