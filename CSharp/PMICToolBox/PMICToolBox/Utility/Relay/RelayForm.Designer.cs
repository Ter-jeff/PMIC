using PmicAutomation.MyControls;

namespace PmicAutomation.Utility.Relay
{
    partial class Relay
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Relay));
            this.FileOpen_ComPin = new PmicAutomation.MyControls.MyFileOpen();
            this.FileOpen_RelayConfig = new PmicAutomation.MyControls.MyFileOpen();
            this.FileOpen_OutputPath = new PmicAutomation.MyControls.MyFileOpen();
            this.Btn_RunDownload = new PmicAutomation.MyControls.MyButtonRunDownload();
            this.richTextBox = new System.Windows.Forms.RichTextBox();
            this.chkboxAdg1414 = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // FileOpen_ComPin
            // 
            this.FileOpen_ComPin.LebalText = "ComPin";
            this.FileOpen_ComPin.Location = new System.Drawing.Point(50, 20);
            this.FileOpen_ComPin.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.FileOpen_ComPin.Name = "FileOpen_ComPin";
            this.FileOpen_ComPin.Size = new System.Drawing.Size(640, 30);
            this.FileOpen_ComPin.TabIndex = 0;
            this.FileOpen_ComPin.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_ComPin_Click);
            // 
            // FileOpen_RelayConfig
            // 
            this.FileOpen_RelayConfig.LebalText = "Relay Config";
            this.FileOpen_RelayConfig.Location = new System.Drawing.Point(50, 60);
            this.FileOpen_RelayConfig.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.FileOpen_RelayConfig.Name = "FileOpen_RelayConfig";
            this.FileOpen_RelayConfig.Size = new System.Drawing.Size(640, 30);
            this.FileOpen_RelayConfig.TabIndex = 1;
            this.FileOpen_RelayConfig.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_RelayConfig_Click);
            // 
            // FileOpen_OutputPath
            // 
            this.FileOpen_OutputPath.LebalText = "OutputPath";
            this.FileOpen_OutputPath.Location = new System.Drawing.Point(50, 100);
            this.FileOpen_OutputPath.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.FileOpen_OutputPath.Name = "FileOpen_OutputPath";
            this.FileOpen_OutputPath.Size = new System.Drawing.Size(640, 30);
            this.FileOpen_OutputPath.TabIndex = 2;
            this.FileOpen_OutputPath.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_OutputPath_Click);
            // 
            // Btn_RunDownload
            // 
            this.Btn_RunDownload.Location = new System.Drawing.Point(440, 152);
            this.Btn_RunDownload.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Btn_RunDownload.Name = "Btn_RunDownload";
            this.Btn_RunDownload.RunText = "One Key";
            this.Btn_RunDownload.Size = new System.Drawing.Size(250, 40);
            this.Btn_RunDownload.TabIndex = 3;
            this.Btn_RunDownload.RunButtonClick += new System.EventHandler(this.Btn_Run_Click);
            this.Btn_RunDownload.DownloadButtonClick += new System.EventHandler(this.Btn_Download_Click);
            // 
            // richTextBox
            // 
            this.richTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.richTextBox.Location = new System.Drawing.Point(30, 217);
            this.richTextBox.Margin = new System.Windows.Forms.Padding(21);
            this.richTextBox.Name = "richTextBox";
            this.richTextBox.Size = new System.Drawing.Size(680, 250);
            this.richTextBox.TabIndex = 4;
            this.richTextBox.Text = "";
            // 
            // chkboxAdg1414
            // 
            this.chkboxAdg1414.AutoSize = true;
            this.chkboxAdg1414.Location = new System.Drawing.Point(30, 170);
            this.chkboxAdg1414.Name = "chkboxAdg1414";
            this.chkboxAdg1414.Size = new System.Drawing.Size(180, 22);
            this.chkboxAdg1414.TabIndex = 5;
            this.chkboxAdg1414.Text = "ADG1414 Matrix Mode";
            this.chkboxAdg1414.UseVisualStyleBackColor = true;
            this.chkboxAdg1414.CheckedChanged += new System.EventHandler(this.chkboxAdg1414_CheckedChanged);
            // 
            // Relay
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(740, 497);
            this.Controls.Add(this.chkboxAdg1414);
            this.Controls.Add(this.FileOpen_RelayConfig);
            this.Controls.Add(this.FileOpen_OutputPath);
            this.Controls.Add(this.FileOpen_ComPin);
            this.Controls.Add(this.Btn_RunDownload);
            this.Controls.Add(this.richTextBox);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.HelpButton = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Relay";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Relay";
            this.Controls.SetChildIndex(this.richTextBox, 0);
            this.Controls.SetChildIndex(this.Btn_RunDownload, 0);
            this.Controls.SetChildIndex(this.FileOpen_ComPin, 0);
            this.Controls.SetChildIndex(this.FileOpen_OutputPath, 0);
            this.Controls.SetChildIndex(this.FileOpen_RelayConfig, 0);
            this.Controls.SetChildIndex(this.chkboxAdg1414, 0);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        public MyFileOpen FileOpen_ComPin;
        public MyFileOpen FileOpen_RelayConfig;
        public MyFileOpen FileOpen_OutputPath;
        private MyButtonRunDownload Btn_RunDownload;
        private System.Windows.Forms.RichTextBox richTextBox;
        public System.Windows.Forms.CheckBox chkboxAdg1414;
    }
}