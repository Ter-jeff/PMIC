using System.Drawing;
using System.Windows.Forms;

namespace CommonLib.Controls
{
    partial class MyFormMini
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.myStatus = new System.Windows.Forms.StatusStrip();
            this.ToolStripStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.ToolStripStatusLabelControlLocation = new System.Windows.Forms.ToolStripStatusLabel();
            this.ProcessTimeToolStripStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.ToolStripProgressBar = new System.Windows.Forms.ToolStripProgressBar();
            this.myStatus.SuspendLayout();
            this.SuspendLayout();
            // 
            // myStatus
            // 
            this.myStatus.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.myStatus.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ToolStripStatusLabel,
            this.ToolStripStatusLabelControlLocation,
            this.ProcessTimeToolStripStatusLabel,
            this.ToolStripProgressBar});
            this.myStatus.Location = new System.Drawing.Point(0, 482);
            this.myStatus.Name = "myStatus";
            this.myStatus.Padding = new System.Windows.Forms.Padding(1, 0, 16, 0);
            this.myStatus.Size = new System.Drawing.Size(716, 26);
            this.myStatus.TabIndex = 13;
            // 
            // ToolStripStatusLabel
            // 
            this.ToolStripStatusLabel.Name = "ToolStripStatusLabel";
            this.ToolStripStatusLabel.Size = new System.Drawing.Size(49, 20);
            this.ToolStripStatusLabel.Text = "Status";
            // 
            // ToolStripStatusLabelControlLocation
            // 
            this.ToolStripStatusLabelControlLocation.Name = "ToolStripStatusLabelControlLocation";
            this.ToolStripStatusLabelControlLocation.Size = new System.Drawing.Size(453, 20);
            this.ToolStripStatusLabelControlLocation.Spring = true;
            // 
            // ProcessTimeToolStripStatusLabel
            // 
            this.ProcessTimeToolStripStatusLabel.Name = "ProcessTimeToolStripStatusLabel";
            this.ProcessTimeToolStripStatusLabel.Size = new System.Drawing.Size(95, 20);
            this.ProcessTimeToolStripStatusLabel.Text = "Process Time";
            this.ProcessTimeToolStripStatusLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // ToolStripProgressBar
            // 
            this.ToolStripProgressBar.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.ToolStripProgressBar.Name = "ToolStripProgressBar";
            this.ToolStripProgressBar.Size = new System.Drawing.Size(100, 18);
            // 
            // MyFormMini
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(716, 508);
            this.Controls.Add(this.myStatus);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "MyFormMini";
            this.myStatus.ResumeLayout(false);
            this.myStatus.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public StatusStrip myStatus;
        public ToolStripStatusLabel ToolStripStatusLabel;
        public ToolStripStatusLabel ToolStripStatusLabelControlLocation;
        public ToolStripStatusLabel ProcessTimeToolStripStatusLabel;
        public ToolStripProgressBar ToolStripProgressBar;
    }
}