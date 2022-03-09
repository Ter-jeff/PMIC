
using System.Diagnostics;
using System.Reflection;

namespace PmicAutogen
{
    sealed partial class PmicMainForm
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
            this.myStatus = new AutomationCommon.Controls.MyStatus();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.splitter2 = new System.Windows.Forms.Splitter();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button_Clear = new System.Windows.Forms.Button();
            this.button_Mbist = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button_Basic = new System.Windows.Forms.CheckBox();
            this.button_VBT = new System.Windows.Forms.CheckBox();
            this.button_Otp = new System.Windows.Forms.CheckBox();
            this.button_Scan = new System.Windows.Forms.CheckBox();
            this.splitter1 = new System.Windows.Forms.Splitter();
            this.button_Setting = new System.Windows.Forms.Button();
            this.button_RunAutogen = new System.Windows.Forms.Button();
            this.button_LoadFiles = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.myFileOpen_SettingFile = new AutomationCommon.Controls.MyFileOpen();
            this.myFileOpen_PatternPath = new AutomationCommon.Controls.MyFileOpen();
            this.myFileOpen_TimeSetPath = new AutomationCommon.Controls.MyFileOpen();
            this.myFileOpen_LibraryPath = new AutomationCommon.Controls.MyFileOpen();
            this.groupBox1 = new AutomationCommon.Controls.MyGroupBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.richTextBox = new System.Windows.Forms.RichTextBox();
            this.myFileOpen_ExtraPath = new AutomationCommon.Controls.MyFileOpen();
            this.tabPage1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // myStatus
            // 
            this.myStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.myStatus.Location = new System.Drawing.Point(0, 599);
            this.myStatus.Name = "myStatus";
            this.myStatus.Size = new System.Drawing.Size(934, 22);
            this.myStatus.TabIndex = 1;
            this.myStatus.Text = "myStatus1";
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.SystemColors.Window;
            this.tabPage1.Controls.Add(this.splitter2);
            this.tabPage1.Controls.Add(this.panel1);
            this.tabPage1.Controls.Add(this.splitter1);
            this.tabPage1.Controls.Add(this.button_Setting);
            this.tabPage1.Controls.Add(this.button_RunAutogen);
            this.tabPage1.Controls.Add(this.button_LoadFiles);
            this.tabPage1.ForeColor = System.Drawing.Color.Black;
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(10, 5, 10, 5);
            this.tabPage1.Size = new System.Drawing.Size(926, 134);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Autogen";
            // 
            // splitter2
            // 
            this.splitter2.BackColor = System.Drawing.SystemColors.ControlLight;
            this.splitter2.Location = new System.Drawing.Point(399, 5);
            this.splitter2.Margin = new System.Windows.Forms.Padding(0);
            this.splitter2.Name = "splitter2";
            this.splitter2.Size = new System.Drawing.Size(3, 124);
            this.splitter2.TabIndex = 6;
            this.splitter2.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.button_Clear);
            this.panel1.Controls.Add(this.button_Mbist);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.button_Basic);
            this.panel1.Controls.Add(this.button_VBT);
            this.panel1.Controls.Add(this.button_Otp);
            this.panel1.Controls.Add(this.button_Scan);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel1.Location = new System.Drawing.Point(253, 5);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(146, 124);
            this.panel1.TabIndex = 5;
            // 
            // button_Clear
            // 
            this.button_Clear.BackColor = System.Drawing.SystemColors.ControlLight;
            this.button_Clear.FlatAppearance.BorderColor = System.Drawing.SystemColors.ControlLight;
            this.button_Clear.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Clear.ForeColor = System.Drawing.Color.Black;
            this.button_Clear.Location = new System.Drawing.Point(16, 11);
            this.button_Clear.Name = "button_Clear";
            this.button_Clear.Size = new System.Drawing.Size(60, 23);
            this.button_Clear.TabIndex = 2;
            this.button_Clear.Text = "Clear";
            this.button_Clear.UseVisualStyleBackColor = false;
            this.button_Clear.Click += new System.EventHandler(this.button_Clear_Click);
            // 
            // button_Mbist
            // 
            this.button_Mbist.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.button_Mbist.BackColor = System.Drawing.SystemColors.Window;
            this.button_Mbist.Location = new System.Drawing.Point(82, 11);
            this.button_Mbist.Name = "button_Mbist";
            this.button_Mbist.Size = new System.Drawing.Size(60, 23);
            this.button_Mbist.TabIndex = 2;
            this.button_Mbist.Text = "Mbist";
            this.button_Mbist.UseVisualStyleBackColor = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(57, 106);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(34, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Block";
            // 
            // button_Basic
            // 
            this.button_Basic.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.button_Basic.BackColor = System.Drawing.SystemColors.Window;
            this.button_Basic.Location = new System.Drawing.Point(16, 40);
            this.button_Basic.Name = "button_Basic";
            this.button_Basic.Size = new System.Drawing.Size(60, 23);
            this.button_Basic.TabIndex = 2;
            this.button_Basic.Text = "Basic";
            this.button_Basic.UseVisualStyleBackColor = false;
            // 
            // button_VBT
            // 
            this.button_VBT.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.button_VBT.BackColor = System.Drawing.SystemColors.Window;
            this.button_VBT.Checked = true;
            this.button_VBT.CheckState = System.Windows.Forms.CheckState.Checked;
            this.button_VBT.Enabled = false;
            this.button_VBT.Location = new System.Drawing.Point(82, 69);
            this.button_VBT.Name = "button_VBT";
            this.button_VBT.Size = new System.Drawing.Size(60, 23);
            this.button_VBT.TabIndex = 2;
            this.button_VBT.Text = "VBT";
            this.button_VBT.UseVisualStyleBackColor = false;
            // 
            // button_Otp
            // 
            this.button_Otp.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.button_Otp.BackColor = System.Drawing.SystemColors.Window;
            this.button_Otp.Location = new System.Drawing.Point(82, 40);
            this.button_Otp.Name = "button_Otp";
            this.button_Otp.Size = new System.Drawing.Size(60, 23);
            this.button_Otp.TabIndex = 2;
            this.button_Otp.Text = "OTP";
            this.button_Otp.UseVisualStyleBackColor = false;
            // 
            // button_Scan
            // 
            this.button_Scan.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.button_Scan.BackColor = System.Drawing.SystemColors.Window;
            this.button_Scan.Location = new System.Drawing.Point(16, 69);
            this.button_Scan.Name = "button_Scan";
            this.button_Scan.Size = new System.Drawing.Size(60, 23);
            this.button_Scan.TabIndex = 2;
            this.button_Scan.Text = "Scan";
            this.button_Scan.UseVisualStyleBackColor = false;
            // 
            // splitter1
            // 
            this.splitter1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.splitter1.Location = new System.Drawing.Point(250, 5);
            this.splitter1.Name = "splitter1";
            this.splitter1.Size = new System.Drawing.Size(3, 124);
            this.splitter1.TabIndex = 4;
            this.splitter1.TabStop = false;
            // 
            // button_Setting
            // 
            this.button_Setting.BackColor = System.Drawing.Color.White;
            this.button_Setting.BackgroundImage = global::PmicAutogen.Properties.Resources.settings_4;
            this.button_Setting.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.button_Setting.Dock = System.Windows.Forms.DockStyle.Left;
            this.button_Setting.FlatAppearance.BorderColor = System.Drawing.SystemColors.Window;
            this.button_Setting.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Setting.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.button_Setting.Location = new System.Drawing.Point(170, 5);
            this.button_Setting.Margin = new System.Windows.Forms.Padding(0);
            this.button_Setting.Name = "button_Setting";
            this.button_Setting.Size = new System.Drawing.Size(80, 124);
            this.button_Setting.TabIndex = 1;
            this.button_Setting.Text = "Setting";
            this.button_Setting.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.button_Setting.UseVisualStyleBackColor = false;
            this.button_Setting.Click += new System.EventHandler(this.button_Setting_Click);
            // 
            // button_RunAutogen
            // 
            this.button_RunAutogen.BackColor = System.Drawing.Color.White;
            this.button_RunAutogen.BackgroundImage = global::PmicAutogen.Properties.Resources.play_button_1;
            this.button_RunAutogen.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.button_RunAutogen.Dock = System.Windows.Forms.DockStyle.Left;
            this.button_RunAutogen.Enabled = false;
            this.button_RunAutogen.FlatAppearance.BorderColor = System.Drawing.SystemColors.Window;
            this.button_RunAutogen.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_RunAutogen.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.button_RunAutogen.Location = new System.Drawing.Point(90, 5);
            this.button_RunAutogen.Margin = new System.Windows.Forms.Padding(0);
            this.button_RunAutogen.Name = "button_RunAutogen";
            this.button_RunAutogen.Size = new System.Drawing.Size(80, 124);
            this.button_RunAutogen.TabIndex = 0;
            this.button_RunAutogen.Text = "RunAutogen";
            this.button_RunAutogen.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.button_RunAutogen.UseVisualStyleBackColor = false;
            this.button_RunAutogen.Click += new System.EventHandler(this.button_RunAutogen_Click);
            // 
            // button_LoadFiles
            // 
            this.button_LoadFiles.BackColor = System.Drawing.Color.White;
            this.button_LoadFiles.BackgroundImage = global::PmicAutogen.Properties.Resources.folder_11;
            this.button_LoadFiles.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.button_LoadFiles.Dock = System.Windows.Forms.DockStyle.Left;
            this.button_LoadFiles.FlatAppearance.BorderColor = System.Drawing.SystemColors.Window;
            this.button_LoadFiles.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_LoadFiles.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.button_LoadFiles.Location = new System.Drawing.Point(10, 5);
            this.button_LoadFiles.Margin = new System.Windows.Forms.Padding(0);
            this.button_LoadFiles.Name = "button_LoadFiles";
            this.button_LoadFiles.Size = new System.Drawing.Size(80, 124);
            this.button_LoadFiles.TabIndex = 0;
            this.button_LoadFiles.Text = "LoadFiles";
            this.button_LoadFiles.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.button_LoadFiles.UseVisualStyleBackColor = false;
            this.button_LoadFiles.Click += new System.EventHandler(this.button_LoadFiles_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Top;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(934, 160);
            this.tabControl1.TabIndex = 2;
            // 
            // myFileOpen_SettingFile
            // 
            this.myFileOpen_SettingFile.AutoSize = true;
            this.myFileOpen_SettingFile.BackColor = System.Drawing.SystemColors.Window;
            this.myFileOpen_SettingFile.Dock = System.Windows.Forms.DockStyle.Top;
            this.myFileOpen_SettingFile.LebalText = "SettingFile(*.xlsx)";
            this.myFileOpen_SettingFile.Location = new System.Drawing.Point(0, 160);
            this.myFileOpen_SettingFile.Name = "myFileOpen_SettingFile";
            this.myFileOpen_SettingFile.Padding = new System.Windows.Forms.Padding(10, 5, 10, 5);
            this.myFileOpen_SettingFile.Size = new System.Drawing.Size(934, 33);
            this.myFileOpen_SettingFile.TabIndex = 19;
            this.myFileOpen_SettingFile.ButtonTextBoxButtonClick += new System.EventHandler(this.FileOpen_SettingFile_ButtonTextBoxButtonClick);
            this.myFileOpen_SettingFile.ButtonTextBoxTextChanged += new System.EventHandler(this.FileOpen_SettingFile_ButtonTextBoxTextChanged);
            // 
            // myFileOpen_PatternPath
            // 
            this.myFileOpen_PatternPath.AutoSize = true;
            this.myFileOpen_PatternPath.BackColor = System.Drawing.SystemColors.Window;
            this.myFileOpen_PatternPath.Dock = System.Windows.Forms.DockStyle.Top;
            this.myFileOpen_PatternPath.LebalText = "Pattern Path";
            this.myFileOpen_PatternPath.Location = new System.Drawing.Point(0, 193);
            this.myFileOpen_PatternPath.Name = "myFileOpen_PatternPath";
            this.myFileOpen_PatternPath.Padding = new System.Windows.Forms.Padding(10, 5, 10, 5);
            this.myFileOpen_PatternPath.Size = new System.Drawing.Size(934, 33);
            this.myFileOpen_PatternPath.TabIndex = 20;
            this.myFileOpen_PatternPath.ButtonTextBoxButtonClick += new System.EventHandler(this.myFileOpen_PatternPath_ButtonTextBoxButtonClick);
            this.myFileOpen_PatternPath.ButtonTextBoxTextChanged += new System.EventHandler(this.myFileOpen_PatternPath_ButtonTextBoxTextChanged);
            // 
            // myFileOpen_TimeSetPath
            // 
            this.myFileOpen_TimeSetPath.AutoSize = true;
            this.myFileOpen_TimeSetPath.BackColor = System.Drawing.SystemColors.Window;
            this.myFileOpen_TimeSetPath.Dock = System.Windows.Forms.DockStyle.Top;
            this.myFileOpen_TimeSetPath.LebalText = "TimeSet Path";
            this.myFileOpen_TimeSetPath.Location = new System.Drawing.Point(0, 226);
            this.myFileOpen_TimeSetPath.Name = "myFileOpen_TimeSetPath";
            this.myFileOpen_TimeSetPath.Padding = new System.Windows.Forms.Padding(10, 5, 10, 5);
            this.myFileOpen_TimeSetPath.Size = new System.Drawing.Size(934, 33);
            this.myFileOpen_TimeSetPath.TabIndex = 21;
            this.myFileOpen_TimeSetPath.ButtonTextBoxButtonClick += new System.EventHandler(this.myFileOpen_TimeSetPath_ButtonTextBoxButtonClick);
            this.myFileOpen_TimeSetPath.ButtonTextBoxTextChanged += new System.EventHandler(this.myFileOpen_TimeSetPath_ButtonTextBoxTextChanged);
            // 
            // myFileOpen_LibraryPath
            // 
            this.myFileOpen_LibraryPath.AutoSize = true;
            this.myFileOpen_LibraryPath.BackColor = System.Drawing.SystemColors.Window;
            this.myFileOpen_LibraryPath.Dock = System.Windows.Forms.DockStyle.Top;
            this.myFileOpen_LibraryPath.LebalText = "Library Path";
            this.myFileOpen_LibraryPath.Location = new System.Drawing.Point(0, 259);
            this.myFileOpen_LibraryPath.Name = "myFileOpen_LibraryPath";
            this.myFileOpen_LibraryPath.Padding = new System.Windows.Forms.Padding(10, 5, 10, 5);
            this.myFileOpen_LibraryPath.Size = new System.Drawing.Size(934, 33);
            this.myFileOpen_LibraryPath.TabIndex = 22;
            this.myFileOpen_LibraryPath.ButtonTextBoxButtonClick += new System.EventHandler(this.myFileOpen_LibraryPath_ButtonTextBoxButtonClick);
            this.myFileOpen_LibraryPath.ButtonTextBoxTextChanged += new System.EventHandler(this.myFileOpen_LibraryPath_ButtonTextBoxTextChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.BorderColor = System.Drawing.SystemColors.Window;
            this.groupBox1.Controls.Add(this.panel2);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.groupBox1.Location = new System.Drawing.Point(0, 325);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(10, 3, 10, 3);
            this.groupBox1.Size = new System.Drawing.Size(934, 274);
            this.groupBox1.TabIndex = 23;
            this.groupBox1.TabStop = false;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.richTextBox);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(10, 16);
            this.panel2.Name = "panel2";
            this.panel2.Padding = new System.Windows.Forms.Padding(10);
            this.panel2.Size = new System.Drawing.Size(914, 255);
            this.panel2.TabIndex = 4;
            // 
            // richTextBox
            // 
            this.richTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.richTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.richTextBox.Location = new System.Drawing.Point(10, 10);
            this.richTextBox.Margin = new System.Windows.Forms.Padding(100);
            this.richTextBox.Name = "richTextBox";
            this.richTextBox.Size = new System.Drawing.Size(894, 235);
            this.richTextBox.TabIndex = 4;
            this.richTextBox.Text = "";
            // 
            // myFileOpen_ExtraPath
            // 
            this.myFileOpen_ExtraPath.AutoSize = true;
            this.myFileOpen_ExtraPath.BackColor = System.Drawing.SystemColors.Window;
            this.myFileOpen_ExtraPath.Dock = System.Windows.Forms.DockStyle.Top;
            this.myFileOpen_ExtraPath.LebalText = "ExtraSheets Path";
            this.myFileOpen_ExtraPath.Location = new System.Drawing.Point(0, 292);
            this.myFileOpen_ExtraPath.Name = "myFileOpen_ExtraPath";
            this.myFileOpen_ExtraPath.Padding = new System.Windows.Forms.Padding(10, 5, 10, 5);
            this.myFileOpen_ExtraPath.Size = new System.Drawing.Size(934, 33);
            this.myFileOpen_ExtraPath.TabIndex = 24;
            this.myFileOpen_ExtraPath.ButtonTextBoxButtonClick += new System.EventHandler(this.myFileOpen_ExtraPath_ButtonTextBoxButtonClick);
            this.myFileOpen_ExtraPath.ButtonTextBoxTextChanged += new System.EventHandler(this.myFileOpen_ExtraPath_ButtonTextBoxTextChanged);
            // 
            // PmicMainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(934, 621);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.myFileOpen_ExtraPath);
            this.Controls.Add(this.myFileOpen_LibraryPath);
            this.Controls.Add(this.myFileOpen_TimeSetPath);
            this.Controls.Add(this.myFileOpen_PatternPath);
            this.Controls.Add(this.myFileOpen_SettingFile);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.myStatus);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.HelpButton = true;
            this.KeyPreview = true;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PmicMainForm";
            this.RightToLeftLayout = true;
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.tabPage1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private AutomationCommon.Controls.MyStatus myStatus;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.Button button_LoadFiles;
        private System.Windows.Forms.Button button_RunAutogen;
        private System.Windows.Forms.Button button_Setting;
        public System.Windows.Forms.CheckBox button_VBT;
        public System.Windows.Forms.CheckBox button_Otp;
        public System.Windows.Forms.CheckBox button_Scan;
        public System.Windows.Forms.CheckBox button_Mbist;
        public System.Windows.Forms.CheckBox button_Basic;
        public System.Windows.Forms.Button button_Clear;
        public AutomationCommon.Controls.MyFileOpen myFileOpen_SettingFile;
        public AutomationCommon.Controls.MyFileOpen myFileOpen_PatternPath;
        public AutomationCommon.Controls.MyFileOpen myFileOpen_TimeSetPath;
        public AutomationCommon.Controls.MyFileOpen myFileOpen_LibraryPath;
        public AutomationCommon.Controls.MyFileOpen myFileOpen_ExtraPath;
        private AutomationCommon.Controls.MyGroupBox groupBox1;
        private System.Windows.Forms.Splitter splitter1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.RichTextBox richTextBox;
        private System.Windows.Forms.Splitter splitter2;
        private System.Windows.Forms.Label label1;
    }
}