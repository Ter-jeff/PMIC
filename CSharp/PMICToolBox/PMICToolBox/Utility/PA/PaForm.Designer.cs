using PmicAutomation.MyControls;

namespace PmicAutomation.Utility.PA
{
    partial class PaParser
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PaParser));
            this.FileOpen_PA = new PmicAutomation.MyControls.MyFileOpen();
            this.FileOpen_TesterConfig = new PmicAutomation.MyControls.MyFileOpen();
            this.FileOpen_OutputPath = new PmicAutomation.MyControls.MyFileOpen();
            this.label_Device = new System.Windows.Forms.Label();
            this.comboBox_Device = new System.Windows.Forms.ComboBox();
            this.label_HexVS = new System.Windows.Forms.Label();
            this.textBox_HexVs = new System.Windows.Forms.TextBox();
            this.Btn_RunDownload = new PmicAutomation.MyControls.MyButtonRunDownload();
            this.richTextBox = new System.Windows.Forms.RichTextBox();
            this.label_SharePinEnable = new System.Windows.Forms.Label();
            this.checkBox_SharePinEnable = new System.Windows.Forms.CheckBox();
            this.label_igxlversion = new System.Windows.Forms.Label();
            this.comboBox_igxlversion = new System.Windows.Forms.ComboBox();
            this.FileOpen_DGSReference = new PmicAutomation.MyControls.MyFileOpen();
            this.SuspendLayout();
            // 
            // FileOpen_PA
            // 
            this.FileOpen_PA.LebalText = "PA File(*.csv,*.xls*)";
            this.FileOpen_PA.Location = new System.Drawing.Point(50, 20);
            this.FileOpen_PA.Name = "FileOpen_PA";
            this.FileOpen_PA.Size = new System.Drawing.Size(640, 30);
            this.FileOpen_PA.TabIndex = 0;
            this.FileOpen_PA.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_PA_Click);
            // 
            // FileOpen_TesterConfig
            // 
            this.FileOpen_TesterConfig.LebalText = "TesterConfig";
            this.FileOpen_TesterConfig.Location = new System.Drawing.Point(50, 60);
            this.FileOpen_TesterConfig.Name = "FileOpen_TesterConfig";
            this.FileOpen_TesterConfig.Size = new System.Drawing.Size(640, 30);
            this.FileOpen_TesterConfig.TabIndex = 1;
            this.FileOpen_TesterConfig.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_TesterConfig_Click);
            // 
            // FileOpen_OutputPath
            // 
            this.FileOpen_OutputPath.LebalText = "OutputPath";
            this.FileOpen_OutputPath.Location = new System.Drawing.Point(50, 100);
            this.FileOpen_OutputPath.Name = "FileOpen_OutputPath";
            this.FileOpen_OutputPath.Size = new System.Drawing.Size(640, 30);
            this.FileOpen_OutputPath.TabIndex = 2;
            this.FileOpen_OutputPath.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_OutputPath_Click);
            // 
            // label_Device
            // 
            this.label_Device.AutoSize = true;
            this.label_Device.Font = new System.Drawing.Font("Microsoft JhengHei UI", 9F);
            this.label_Device.Location = new System.Drawing.Point(50, 189);
            this.label_Device.Name = "label_Device";
            this.label_Device.Size = new System.Drawing.Size(76, 15);
            this.label_Device.TabIndex = 3;
            this.label_Device.Text = "Device Type";
            // 
            // comboBox_Device
            // 
            this.comboBox_Device.Font = new System.Drawing.Font("Microsoft JhengHei UI", 9F);
            this.comboBox_Device.FormattingEnabled = true;
            this.comboBox_Device.Items.AddRange(new object[] {
            "AP",
            "PMIC",
            "RF"});
            this.comboBox_Device.Location = new System.Drawing.Point(50, 213);
            this.comboBox_Device.Name = "comboBox_Device";
            this.comboBox_Device.Size = new System.Drawing.Size(90, 23);
            this.comboBox_Device.TabIndex = 4;
            this.comboBox_Device.TextChanged += new System.EventHandler(this.ComboBox_Device_TextChanged);
            // 
            // label_HexVS
            // 
            this.label_HexVS.AutoSize = true;
            this.label_HexVS.Font = new System.Drawing.Font("Microsoft JhengHei UI", 9F);
            this.label_HexVS.Location = new System.Drawing.Point(190, 189);
            this.label_HexVS.Name = "label_HexVS";
            this.label_HexVS.Size = new System.Drawing.Size(74, 15);
            this.label_HexVS.TabIndex = 5;
            this.label_HexVS.Text = "HexVS Slots";
            // 
            // textBox_HexVs
            // 
            this.textBox_HexVs.Font = new System.Drawing.Font("Microsoft JhengHei UI", 9F);
            this.textBox_HexVs.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.textBox_HexVs.Location = new System.Drawing.Point(190, 213);
            this.textBox_HexVs.Name = "textBox_HexVs";
            this.textBox_HexVs.Size = new System.Drawing.Size(500, 23);
            this.textBox_HexVs.TabIndex = 6;
            // 
            // Btn_RunDownload
            // 
            this.Btn_RunDownload.Location = new System.Drawing.Point(440, 254);
            this.Btn_RunDownload.Margin = new System.Windows.Forms.Padding(10);
            this.Btn_RunDownload.Name = "Btn_RunDownload";
            this.Btn_RunDownload.RunText = "Run";
            this.Btn_RunDownload.Size = new System.Drawing.Size(250, 40);
            this.Btn_RunDownload.TabIndex = 8;
            this.Btn_RunDownload.RunButtonClick += new System.EventHandler(this.Btn_Run_Click);
            this.Btn_RunDownload.DownloadButtonClick += new System.EventHandler(this.Btn_Download_Click);
            // 
            // richTextBox
            // 
            this.richTextBox.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.richTextBox.Font = new System.Drawing.Font("Microsoft JhengHei UI", 9F);
            this.richTextBox.Location = new System.Drawing.Point(30, 312);
            this.richTextBox.Margin = new System.Windows.Forms.Padding(21);
            this.richTextBox.Name = "richTextBox";
            this.richTextBox.Size = new System.Drawing.Size(680, 191);
            this.richTextBox.TabIndex = 7;
            this.richTextBox.Text = "";
            // 
            // label_SharePinEnable
            // 
            this.label_SharePinEnable.AutoSize = true;
            this.label_SharePinEnable.Location = new System.Drawing.Point(186, 263);
            this.label_SharePinEnable.Name = "label_SharePinEnable";
            this.label_SharePinEnable.Size = new System.Drawing.Size(56, 15);
            this.label_SharePinEnable.TabIndex = 9;
            this.label_SharePinEnable.Text = "SharePin";
            // 
            // checkBox_SharePinEnable
            // 
            this.checkBox_SharePinEnable.AutoSize = true;
            this.checkBox_SharePinEnable.Location = new System.Drawing.Point(190, 285);
            this.checkBox_SharePinEnable.Name = "checkBox_SharePinEnable";
            this.checkBox_SharePinEnable.Size = new System.Drawing.Size(117, 19);
            this.checkBox_SharePinEnable.TabIndex = 10;
            this.checkBox_SharePinEnable.Text = "SharePin Enable";
            this.checkBox_SharePinEnable.UseVisualStyleBackColor = true;
            // 
            // label_igxlversion
            // 
            this.label_igxlversion.AutoSize = true;
            this.label_igxlversion.Location = new System.Drawing.Point(50, 254);
            this.label_igxlversion.Name = "label_igxlversion";
            this.label_igxlversion.Size = new System.Drawing.Size(78, 15);
            this.label_igxlversion.TabIndex = 11;
            this.label_igxlversion.Text = "IGXL Version";
            // 
            // comboBox_igxlversion
            // 
            this.comboBox_igxlversion.FormattingEnabled = true;
            this.comboBox_igxlversion.Location = new System.Drawing.Point(51, 279);
            this.comboBox_igxlversion.Name = "comboBox_igxlversion";
            this.comboBox_igxlversion.Size = new System.Drawing.Size(90, 23);
            this.comboBox_igxlversion.TabIndex = 12;
            // 
            // FileOpen_DGSReference
            // 
            this.FileOpen_DGSReference.LebalText = "DGS Reference";
            this.FileOpen_DGSReference.Location = new System.Drawing.Point(50, 140);
            this.FileOpen_DGSReference.Name = "FileOpen_DGSReference";
            this.FileOpen_DGSReference.Size = new System.Drawing.Size(640, 30);
            this.FileOpen_DGSReference.TabIndex = 3;
            this.FileOpen_DGSReference.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_DGSReference_Click);
            // 
            // PaParser
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(740, 540);
            this.Controls.Add(this.FileOpen_DGSReference);
            this.Controls.Add(this.comboBox_igxlversion);
            this.Controls.Add(this.label_igxlversion);
            this.Controls.Add(this.checkBox_SharePinEnable);
            this.Controls.Add(this.label_SharePinEnable);
            this.Controls.Add(this.FileOpen_PA);
            this.Controls.Add(this.FileOpen_OutputPath);
            this.Controls.Add(this.FileOpen_TesterConfig);
            this.Controls.Add(this.Btn_RunDownload);
            this.Controls.Add(this.richTextBox);
            this.Controls.Add(this.label_Device);
            this.Controls.Add(this.comboBox_Device);
            this.Controls.Add(this.label_HexVS);
            this.Controls.Add(this.textBox_HexVs);
            this.Font = new System.Drawing.Font("Microsoft JhengHei UI", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.HelpButton = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PaParser";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PinMap & Channel Map Generator";
            this.Controls.SetChildIndex(this.textBox_HexVs, 0);
            this.Controls.SetChildIndex(this.label_HexVS, 0);
            this.Controls.SetChildIndex(this.comboBox_Device, 0);
            this.Controls.SetChildIndex(this.label_Device, 0);
            this.Controls.SetChildIndex(this.richTextBox, 0);
            this.Controls.SetChildIndex(this.Btn_RunDownload, 0);
            this.Controls.SetChildIndex(this.FileOpen_TesterConfig, 0);
            this.Controls.SetChildIndex(this.FileOpen_OutputPath, 0);
            this.Controls.SetChildIndex(this.FileOpen_PA, 0);
            this.Controls.SetChildIndex(this.label_SharePinEnable, 0);
            this.Controls.SetChildIndex(this.checkBox_SharePinEnable, 0);
            this.Controls.SetChildIndex(this.label_igxlversion, 0);
            this.Controls.SetChildIndex(this.comboBox_igxlversion, 0);
            this.Controls.SetChildIndex(this.FileOpen_DGSReference, 0);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        public MyFileOpen FileOpen_PA;
        public MyFileOpen FileOpen_OutputPath;
        public MyFileOpen FileOpen_TesterConfig;
        private System.Windows.Forms.Label label_Device;
        public System.Windows.Forms.ComboBox comboBox_Device;
        private System.Windows.Forms.Label label_HexVS;
        public System.Windows.Forms.TextBox textBox_HexVs;
        private MyButtonRunDownload Btn_RunDownload;
        private System.Windows.Forms.RichTextBox richTextBox;
        private System.Windows.Forms.Label label_SharePinEnable;
        public System.Windows.Forms.CheckBox checkBox_SharePinEnable;
        private System.Windows.Forms.Label label_igxlversion;
        private System.Windows.Forms.ComboBox comboBox_igxlversion;
        public MyFileOpen FileOpen_DGSReference;
    }
}