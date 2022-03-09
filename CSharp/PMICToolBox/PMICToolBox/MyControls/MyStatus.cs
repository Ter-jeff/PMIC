using System.Windows.Forms;

namespace PmicAutomation.MyControls
{
    public sealed class MyStatus : StatusStrip
    {
        public MyStatus()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            StatusStrip1 = new StatusStrip();
            LabelStatus = new ToolStripStatusLabel();
            LabelEmpty = new ToolStripStatusLabel();
            LabelProcessTime = new ToolStripStatusLabel();
            ProgressBar = new ToolStripProgressBar();
            StatusStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // statusStrip1
            // 
            Items.AddRange(new ToolStripItem[] {
            LabelStatus,
            LabelEmpty,
            LabelProcessTime,
            ProgressBar});
            StatusStrip1.Location = new System.Drawing.Point(0, 128);
            StatusStrip1.Name = "StatusStrip1";
            StatusStrip1.Size = new System.Drawing.Size(575, 22);
            StatusStrip1.TabIndex = 0;
            StatusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            LabelStatus.Name = "LabelStatus";
            LabelStatus.Size = new System.Drawing.Size(39, 17);
            LabelStatus.Text = "Status";
            // 
            // toolStripStatusLabel2
            // 
            LabelEmpty.Name = "LabelEmpty";
            LabelEmpty.Size = new System.Drawing.Size(388, 17);
            LabelEmpty.Spring = true;
            // 
            // toolStripStatusLabel3
            // 
            LabelProcessTime.Name = "LabelProcessTime";
            LabelProcessTime.Size = new System.Drawing.Size(388, 17);
            LabelProcessTime.Text = "Process Time";
            // 
            // toolStripProgressBar1
            // 
            ProgressBar.Name = "ProgressBar";
            ProgressBar.Size = new System.Drawing.Size(100, 16);
            ProgressBar.Maximum = 10;
            ProgressBar.Minimum = 0;
            // 
            // MyStatus
            // 
            Name = "MyStatus";
            Size = new System.Drawing.Size(575, 150);
            StatusStrip1.ResumeLayout(false);
            StatusStrip1.PerformLayout();
            ResumeLayout(false);

        }

        public StatusStrip StatusStrip1;
        public ToolStripStatusLabel LabelStatus;
        public ToolStripStatusLabel LabelEmpty;
        public ToolStripStatusLabel LabelProcessTime;
        public ToolStripProgressBar ProgressBar;
    }
}