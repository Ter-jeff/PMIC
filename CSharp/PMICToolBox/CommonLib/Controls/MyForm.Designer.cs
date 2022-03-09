using System.Drawing;
using System.Windows.Forms;

namespace CommonLib.Controls
{
    partial class MyForm
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
            this.ToolStripProgressBar = new System.Windows.Forms.ToolStripProgressBar();
            this.panel2 = new System.Windows.Forms.Panel();
            this.richTextBox = new System.Windows.Forms.RichTextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.ProcessTimeToolStripStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.myStatus.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // myStatus
            // 
            this.myStatus.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ToolStripStatusLabel,
            this.ToolStripStatusLabelControlLocation,
            this.ProcessTimeToolStripStatusLabel,
            this.ToolStripProgressBar});
            this.myStatus.Location = new System.Drawing.Point(0, 486);
            this.myStatus.Name = "myStatus";
            this.myStatus.Padding = new System.Windows.Forms.Padding(1, 0, 16, 0);
            this.myStatus.Size = new System.Drawing.Size(716, 22);
            this.myStatus.TabIndex = 13;
            // 
            // toolStripStatusLabel1
            // 
            this.ToolStripStatusLabel.Name = "toolStripStatusLabel1";
            this.ToolStripStatusLabel.Size = new System.Drawing.Size(118, 17);
            this.ToolStripStatusLabel.Text = "Status";
            // 
            // toolStripStatusLabel2
            // 
            this.ToolStripStatusLabelControlLocation.Name = "toolStripStatusLabel2";
            this.ToolStripStatusLabelControlLocation.Size = new System.Drawing.Size(118, 17);
            this.ToolStripStatusLabelControlLocation.Spring = true;
            // 
            // toolStripStatusLabel3
            // 
            this.ProcessTimeToolStripStatusLabel.Name = "toolStripStatusLabel3";
            this.ProcessTimeToolStripStatusLabel.Size = new System.Drawing.Size(118, 17);
            this.ProcessTimeToolStripStatusLabel.Text = "Process Time";
            this.ProcessTimeToolStripStatusLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // toolStripProgressBar1
            // 
            this.ToolStripProgressBar.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.ToolStripProgressBar.Name = "toolStripProgressBar1";
            this.ToolStripProgressBar.Size = new System.Drawing.Size(100, 16);
            // 
            // panel2
            // 
            this.panel2.AutoSize = true;
            this.panel2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.panel2.Controls.Add(this.richTextBox);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 50);
            this.panel2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.panel2.Name = "panel2";
            this.panel2.Padding = new System.Windows.Forms.Padding(23, 0, 23, 25);
            this.panel2.Size = new System.Drawing.Size(716, 436);
            this.panel2.TabIndex = 11;
            // 
            // richTextBox
            // 
            this.richTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.richTextBox.Location = new System.Drawing.Point(23, 0);
            this.richTextBox.Margin = new System.Windows.Forms.Padding(23, 25, 23, 25);
            this.richTextBox.Name = "richTextBox";
            this.richTextBox.Size = new System.Drawing.Size(670, 411);
            this.richTextBox.TabIndex = 10;
            this.richTextBox.Text = "";
            // 
            // panel1
            // 
            this.panel1.AutoSize = true;
            this.panel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.panel1.Name = "panel1";
            this.panel1.Padding = new System.Windows.Forms.Padding(23, 25, 23, 25);
            this.panel1.Size = new System.Drawing.Size(716, 50);
            this.panel1.TabIndex = 12;
            // 
            // MyForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(716, 508);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.myStatus);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "MyForm";
            this.myStatus.ResumeLayout(false);
            this.myStatus.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        public System.Windows.Forms.RichTextBox richTextBox;
        public System.Windows.Forms.Panel panel2;
        public System.Windows.Forms.Panel panel1;

        public StatusStrip myStatus;
        public ToolStripStatusLabel ToolStripStatusLabel;
        public ToolStripStatusLabel ToolStripStatusLabelControlLocation;
        public ToolStripStatusLabel ProcessTimeToolStripStatusLabel;
        public ToolStripProgressBar ToolStripProgressBar;
    }
}