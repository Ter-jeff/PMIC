using PmicAutomation.MyControls;

namespace PmicAutomation.Utility.OTPRegisterMap
{
    public partial class OtpRegisterMapFrom
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(OtpRegisterMapFrom));
            this.FileOpen_Yaml = new PmicAutomation.MyControls.MyFileOpen();
            this.FileOpen_OutputPath = new PmicAutomation.MyControls.MyFileOpen();
            this.FileOpen_RegMap = new PmicAutomation.MyControls.MyFileOpen();
            this.Btn_RunDownload = new PmicAutomation.MyControls.MyButtonRunDownload();
            this.richTextBox = new System.Windows.Forms.RichTextBox();
            this.FilesOpen_Otp = new PmicAutomation.MyControls.MyFileOpen();
            this.SuspendLayout();
            // 
            // FileOpen_Yaml
            // 
            this.FileOpen_Yaml.LebalText = "Yaml File";
            this.FileOpen_Yaml.Location = new System.Drawing.Point(50, 20);
            this.FileOpen_Yaml.Name = "FileOpen_Yaml";
            this.FileOpen_Yaml.Size = new System.Drawing.Size(640, 30);
            this.FileOpen_Yaml.TabIndex = 0;
            this.FileOpen_Yaml.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_YamlFile_Click);
            // 
            // FileOpen_OutputPath
            // 
            this.FileOpen_OutputPath.LebalText = "OutputPath";
            this.FileOpen_OutputPath.Location = new System.Drawing.Point(51, 146);
            this.FileOpen_OutputPath.Name = "FileOpen_OutputPath";
            this.FileOpen_OutputPath.Size = new System.Drawing.Size(640, 30);
            this.FileOpen_OutputPath.TabIndex = 1;
            this.FileOpen_OutputPath.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_OutputPath_Click);
            // 
            // FileOpen_RegMap
            // 
            this.FileOpen_RegMap.LebalText = "RegisterMap";
            this.FileOpen_RegMap.Location = new System.Drawing.Point(50, 107);
            this.FileOpen_RegMap.Name = "FileOpen_RegMap";
            this.FileOpen_RegMap.Size = new System.Drawing.Size(640, 33);
            this.FileOpen_RegMap.TabIndex = 2;
            this.FileOpen_RegMap.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_RegMap_Click);
            // 
            // Btn_RunDownload
            // 
            this.Btn_RunDownload.Location = new System.Drawing.Point(460, 230);
            this.Btn_RunDownload.Margin = new System.Windows.Forms.Padding(0);
            this.Btn_RunDownload.Name = "Btn_RunDownload";
            this.Btn_RunDownload.RunText = "Run";
            this.Btn_RunDownload.Size = new System.Drawing.Size(250, 40);
            this.Btn_RunDownload.TabIndex = 3;
            this.Btn_RunDownload.RunButtonClick += new System.EventHandler(this.Btn_Run_Click);
            this.Btn_RunDownload.DownloadButtonClick += new System.EventHandler(this.Btn_Download_Click);
            // 
            // richTextBox
            // 
            this.richTextBox.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.richTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.richTextBox.Location = new System.Drawing.Point(30, 275);
            this.richTextBox.Margin = new System.Windows.Forms.Padding(21);
            this.richTextBox.Name = "richTextBox";
            this.richTextBox.Size = new System.Drawing.Size(680, 260);
            this.richTextBox.TabIndex = 4;
            this.richTextBox.Text = "";
            // 
            // FilesOpen_Otp
            // 
            this.FilesOpen_Otp.LebalText = "OTP Files";
            this.FilesOpen_Otp.Location = new System.Drawing.Point(50, 64);
            this.FilesOpen_Otp.Name = "FilesOpen_Otp";
            this.FilesOpen_Otp.Size = new System.Drawing.Size(640, 30);
            this.FilesOpen_Otp.TabIndex = 5;
            this.FilesOpen_Otp.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_OtpFilesFile_Click);
            // 
            // OtpRegisterMapFrom
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(740, 632);
            this.Controls.Add(this.FilesOpen_Otp);
            this.Controls.Add(this.FileOpen_OutputPath);
            this.Controls.Add(this.FileOpen_Yaml);
            this.Controls.Add(this.FileOpen_RegMap);
            this.Controls.Add(this.Btn_RunDownload);
            this.Controls.Add(this.richTextBox);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.HelpButton = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "OtpRegisterMapFrom";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "OTP Register Map";
            this.Controls.SetChildIndex(this.richTextBox, 0);
            this.Controls.SetChildIndex(this.Btn_RunDownload, 0);
            this.Controls.SetChildIndex(this.FileOpen_RegMap, 0);
            this.Controls.SetChildIndex(this.FileOpen_Yaml, 0);
            this.Controls.SetChildIndex(this.FileOpen_OutputPath, 0);
            this.Controls.SetChildIndex(this.FilesOpen_Otp, 0);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        public MyFileOpen FileOpen_Yaml;
        public MyFileOpen FileOpen_OutputPath;
        public MyFileOpen FileOpen_RegMap;
        private MyButtonRunDownload Btn_RunDownload;
        private System.Windows.Forms.RichTextBox richTextBox;
        public MyFileOpen FilesOpen_Otp;
    }
}