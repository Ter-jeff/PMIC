using PmicAutomation.MyControls;

namespace PmicAutomation.Utility.VbtGenToolTemplate
{
    partial class VbtGenToolGenerator
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(VbtGenToolGenerator));
            this.FileOpen_TCM = new MyFileOpen();
            this.FileOpen_OutputPath = new MyFileOpen();
            this.Btn_RunDownload = new MyButtonRunDownload();
            this.richTextBox = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // FileOpen_TCM
            // 
            this.FileOpen_TCM.LebalText = "TCM";
            this.FileOpen_TCM.Location = new System.Drawing.Point(50, 20);
            this.FileOpen_TCM.Name = "FileOpen_TCM";
            this.FileOpen_TCM.Size = new System.Drawing.Size(640, 30);
            this.FileOpen_TCM.TabIndex = 0;
            this.FileOpen_TCM.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_TCM_Click);
            // 
            // FileOpen_OutputPath
            // 
            this.FileOpen_OutputPath.LebalText = "OutputPath";
            this.FileOpen_OutputPath.Location = new System.Drawing.Point(50, 60);
            this.FileOpen_OutputPath.Name = "FileOpen_OutputPath";
            this.FileOpen_OutputPath.Size = new System.Drawing.Size(640, 30);
            this.FileOpen_OutputPath.TabIndex = 1;
            this.FileOpen_OutputPath.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_OutputPath_Click);
            // 
            // Btn_RunDownload
            // 
            this.Btn_RunDownload.Location = new System.Drawing.Point(440, 114);
            this.Btn_RunDownload.Name = "Btn_RunDownload";
            this.Btn_RunDownload.RunText = "Run";
            this.Btn_RunDownload.Size = new System.Drawing.Size(250, 40);
            this.Btn_RunDownload.TabIndex = 2;
            this.Btn_RunDownload.RunButtonClick += new System.EventHandler(this.Btn_Run_Click);
            this.Btn_RunDownload.DownloadButtonClick += new System.EventHandler(this.Btn_Download_Click);
            // 
            // richTextBox
            // 
            this.richTextBox.Location = new System.Drawing.Point(30, 187);
            this.richTextBox.Margin = new System.Windows.Forms.Padding(21);
            this.richTextBox.Name = "richTextBox";
            this.richTextBox.Size = new System.Drawing.Size(680, 280);
            this.richTextBox.TabIndex = 3;
            this.richTextBox.Text = "";
            // 
            // VbtGenToolGenerator
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(740, 497);
            this.Controls.Add(this.FileOpen_OutputPath);
            this.Controls.Add(this.FileOpen_TCM);
            this.Controls.Add(this.Btn_RunDownload);
            this.Controls.Add(this.richTextBox);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.HelpButton = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "VbtGenToolGenerator";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Vbt Gen Tool Template Generator";
            this.ResumeLayout(false);

        }
        #endregion

        public MyFileOpen FileOpen_TCM;
        public MyFileOpen FileOpen_OutputPath;
        private MyButtonRunDownload Btn_RunDownload;
        private System.Windows.Forms.RichTextBox richTextBox;
    }
}