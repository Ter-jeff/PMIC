using PmicAutomation.MyControls;
using System;
using System.Collections.Generic;

namespace PmicAutomation.Utility.PatSetsAll
{
    public partial class PatSetsAllForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PatSetsAllForm));
            this.FileOpen_InputPath = new PmicAutomation.MyControls.MyFileOpen();
            this.FileOpen_OutputPath = new PmicAutomation.MyControls.MyFileOpen();
            this.Btn_RunDownload = new PmicAutomation.MyControls.MyButtonRunDownload();
            this.richTextBox = new System.Windows.Forms.RichTextBox();
            this.radioButton_relativePath = new System.Windows.Forms.RadioButton();
            this.radioButton_absolutePath = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.comboBoxIgxlVersion = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.rBtnGzOnly = new System.Windows.Forms.RadioButton();
            this.rBtnPatOnly = new System.Windows.Forms.RadioButton();
            this.rBtnAll = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // FileOpen_InputPath
            // 
            this.FileOpen_InputPath.LebalText = "InputPath";
            this.FileOpen_InputPath.Location = new System.Drawing.Point(50, 20);
            this.FileOpen_InputPath.Name = "FileOpen_InputPath";
            this.FileOpen_InputPath.Size = new System.Drawing.Size(640, 30);
            this.FileOpen_InputPath.TabIndex = 0;
            this.FileOpen_InputPath.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_InputPath_Click);
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
            this.Btn_RunDownload.Location = new System.Drawing.Point(440, 185);
            this.Btn_RunDownload.Margin = new System.Windows.Forms.Padding(0);
            this.Btn_RunDownload.Name = "Btn_RunDownload";
            this.Btn_RunDownload.RunText = "Run";
            this.Btn_RunDownload.Size = new System.Drawing.Size(250, 40);
            this.Btn_RunDownload.TabIndex = 2;
            this.Btn_RunDownload.RunButtonClick += new System.EventHandler(this.Btn_Run_Click);
            this.Btn_RunDownload.DownloadButtonClick += new System.EventHandler(this.Btn_Download_Click);
            // 
            // richTextBox
            // 
            this.richTextBox.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.richTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.richTextBox.Location = new System.Drawing.Point(30, 255);
            this.richTextBox.Margin = new System.Windows.Forms.Padding(21);
            this.richTextBox.Name = "richTextBox";
            this.richTextBox.Size = new System.Drawing.Size(680, 260);
            this.richTextBox.TabIndex = 3;
            this.richTextBox.Text = "";
            // 
            // radioButton_relativePath
            // 
            this.radioButton_relativePath.AutoSize = true;
            this.radioButton_relativePath.Location = new System.Drawing.Point(14, 43);
            this.radioButton_relativePath.Name = "radioButton_relativePath";
            this.radioButton_relativePath.Size = new System.Drawing.Size(97, 19);
            this.radioButton_relativePath.TabIndex = 6;
            this.radioButton_relativePath.Text = "Relative Path";
            this.radioButton_relativePath.UseVisualStyleBackColor = true;
            // 
            // radioButton_absolutePath
            // 
            this.radioButton_absolutePath.AutoSize = true;
            this.radioButton_absolutePath.Checked = true;
            this.radioButton_absolutePath.Location = new System.Drawing.Point(14, 22);
            this.radioButton_absolutePath.Name = "radioButton_absolutePath";
            this.radioButton_absolutePath.Size = new System.Drawing.Size(100, 19);
            this.radioButton_absolutePath.TabIndex = 7;
            this.radioButton_absolutePath.TabStop = true;
            this.radioButton_absolutePath.Text = "Absolute Path";
            this.radioButton_absolutePath.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(54, 106);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 15);
            this.label1.TabIndex = 9;
            this.label1.Text = "IGXL Version";
            // 
            // comboBoxIgxlVersion
            // 
            this.comboBoxIgxlVersion.FormattingEnabled = true;
            this.comboBoxIgxlVersion.Location = new System.Drawing.Point(190, 106);
            this.comboBoxIgxlVersion.Name = "comboBoxIgxlVersion";
            this.comboBoxIgxlVersion.Size = new System.Drawing.Size(211, 23);
            this.comboBoxIgxlVersion.TabIndex = 10;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioButton_absolutePath);
            this.groupBox1.Controls.Add(this.radioButton_relativePath);
            this.groupBox1.Location = new System.Drawing.Point(30, 141);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(171, 100);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Path";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.rBtnAll);
            this.groupBox2.Controls.Add(this.rBtnPatOnly);
            this.groupBox2.Controls.Add(this.rBtnGzOnly);
            this.groupBox2.Location = new System.Drawing.Point(207, 141);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(200, 100);
            this.groupBox2.TabIndex = 12;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Pattern Type";
            // 
            // rBtnGzOnly
            // 
            this.rBtnGzOnly.AutoSize = true;
            this.rBtnGzOnly.Checked = true;
            this.rBtnGzOnly.Location = new System.Drawing.Point(16, 20);
            this.rBtnGzOnly.Name = "rBtnGzOnly";
            this.rBtnGzOnly.Size = new System.Drawing.Size(70, 19);
            this.rBtnGzOnly.TabIndex = 0;
            this.rBtnGzOnly.TabStop = true;
            this.rBtnGzOnly.Text = ".Gz Only";
            this.rBtnGzOnly.UseVisualStyleBackColor = true;
            // 
            // rBtnPatOnly
            // 
            this.rBtnPatOnly.AutoSize = true;
            this.rBtnPatOnly.Location = new System.Drawing.Point(16, 44);
            this.rBtnPatOnly.Name = "rBtnPatOnly";
            this.rBtnPatOnly.Size = new System.Drawing.Size(73, 19);
            this.rBtnPatOnly.TabIndex = 1;
            this.rBtnPatOnly.Text = ".Pat Only";
            this.rBtnPatOnly.UseVisualStyleBackColor = true;
            // 
            // rBtnAll
            // 
            this.rBtnAll.AutoSize = true;
            this.rBtnAll.Location = new System.Drawing.Point(16, 68);
            this.rBtnAll.Name = "rBtnAll";
            this.rBtnAll.Size = new System.Drawing.Size(81, 19);
            this.rBtnAll.TabIndex = 2;
            this.rBtnAll.Text = "Include All";
            this.rBtnAll.UseVisualStyleBackColor = true;
            // 
            // PatSetsAllForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(740, 592);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.comboBoxIgxlVersion);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.FileOpen_OutputPath);
            this.Controls.Add(this.FileOpen_InputPath);
            this.Controls.Add(this.Btn_RunDownload);
            this.Controls.Add(this.richTextBox);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.HelpButton = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PatSetsAllForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PatSets All Generator";
            this.Controls.SetChildIndex(this.richTextBox, 0);
            this.Controls.SetChildIndex(this.Btn_RunDownload, 0);
            this.Controls.SetChildIndex(this.FileOpen_InputPath, 0);
            this.Controls.SetChildIndex(this.FileOpen_OutputPath, 0);
            this.Controls.SetChildIndex(this.label1, 0);
            this.Controls.SetChildIndex(this.comboBoxIgxlVersion, 0);
            this.Controls.SetChildIndex(this.groupBox1, 0);
            this.Controls.SetChildIndex(this.groupBox2, 0);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        public MyFileOpen FileOpen_InputPath;
        public MyFileOpen FileOpen_OutputPath;
        private MyButtonRunDownload Btn_RunDownload;
        private System.Windows.Forms.RichTextBox richTextBox;
        public System.Windows.Forms.RadioButton radioButton_relativePath;
        public System.Windows.Forms.RadioButton radioButton_absolutePath;
        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.ComboBox comboBoxIgxlVersion;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        public System.Windows.Forms.RadioButton rBtnAll;
        public System.Windows.Forms.RadioButton rBtnPatOnly;
        public System.Windows.Forms.RadioButton rBtnGzOnly;
    }
}