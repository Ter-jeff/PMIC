using System;

namespace AutomationCommon.Controls
{
    partial class MyButtonRunDownload
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public event EventHandler RunButtonClick { add { Run.Click += value; } remove { Run.Click -= value; } }

        public event EventHandler DownloadButtonClick { add { Template.Click += value; } remove { Template.Click -= value; } }

        public string RunText { get { return this.Run.Text ; } set { this.Run.Text = value; } }

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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Run = new System.Windows.Forms.Button();
            this.Template = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Run
            // 
            this.Run.Enabled = false;
            this.Run.Font = new System.Drawing.Font("Microsoft JhengHei UI", 9F);
            this.Run.Location = new System.Drawing.Point(0, 0);
            this.Run.Margin = new System.Windows.Forms.Padding(0);
            this.Run.Name = "Run";
            this.Run.Size = new System.Drawing.Size(110, 40);
            this.Run.TabIndex = 0;
            this.Run.Text = "Run";
            this.Run.Enabled = false;
            // 
            // Template
            // 
            this.Template.Font = new System.Drawing.Font("Microsoft JhengHei UI", 9F);
            this.Template.Location = new System.Drawing.Point(140, 0);
            this.Template.Margin = new System.Windows.Forms.Padding(0);
            this.Template.Name = "Template";
            this.Template.Size = new System.Drawing.Size(110, 40);
            this.Template.TabIndex = 2;
            this.Template.Text = "Template";
            // 
            // ButtonRunDownload
            // 
            this.Controls.Add(this.Template);
            this.Controls.Add(this.Run);
            this.Name = "ButtonRunDownload";
            this.Size = new System.Drawing.Size(250, 40);
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.Button Run;
        public System.Windows.Forms.Button Template;
    }
}
