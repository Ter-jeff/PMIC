namespace PmicAutomation.Utility.TCMID
{
    partial class TCMIDForm
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
            this.richTextBox = new System.Windows.Forms.RichTextBox();
            this.Btn_RunDownload = new PmicAutomation.MyControls.MyButtonRunDownload();
            this.outputPath = new PmicAutomation.MyControls.MyFileOpen();
            this.inputPath = new PmicAutomation.MyControls.MyFileOpen();
            this.label1 = new System.Windows.Forms.Label();
            this.tb_version = new System.Windows.Forms.TextBox();
            this.skipItem = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // richTextBox
            // 
            this.richTextBox.Location = new System.Drawing.Point(21, 186);
            this.richTextBox.Name = "richTextBox";
            this.richTextBox.Size = new System.Drawing.Size(644, 186);
            this.richTextBox.TabIndex = 3;
            this.richTextBox.Text = "";
            // 
            // Btn_RunDownload
            // 
            this.Btn_RunDownload.Location = new System.Drawing.Point(412, 115);
            this.Btn_RunDownload.Name = "Btn_RunDownload";
            this.Btn_RunDownload.RunText = "Run";
            this.Btn_RunDownload.Size = new System.Drawing.Size(253, 42);
            this.Btn_RunDownload.TabIndex = 2;
            this.Btn_RunDownload.RunButtonClick += new System.EventHandler(this.Btn_Run_Click);
            this.Btn_RunDownload.DownloadButtonClick += new System.EventHandler(this.Btn_Download_Click);
            // 
            // outputPath
            // 
            this.outputPath.LebalText = "Output Path";
            this.outputPath.Location = new System.Drawing.Point(21, 60);
            this.outputPath.Name = "outputPath";
            this.outputPath.Size = new System.Drawing.Size(644, 31);
            this.outputPath.TabIndex = 1;
            this.outputPath.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_OutputPath_Click);
            // 
            // inputPath
            // 
            this.inputPath.LebalText = "Input Path";
            this.inputPath.Location = new System.Drawing.Point(21, 23);
            this.inputPath.Name = "inputPath";
            this.inputPath.Size = new System.Drawing.Size(644, 31);
            this.inputPath.TabIndex = 0;
            this.inputPath.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_InputPath_Click);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(21, 104);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(64, 22);
            this.label1.TabIndex = 4;
            this.label1.Text = "T/P Version";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // tb_version
            // 
            this.tb_version.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(254)));
            this.tb_version.Location = new System.Drawing.Point(161, 104);
            this.tb_version.Name = "tb_version";
            this.tb_version.Size = new System.Drawing.Size(106, 21);
            this.tb_version.TabIndex = 5;
            // 
            // skipItem
            // 
            this.skipItem.AutoSize = true;
            this.skipItem.Location = new System.Drawing.Point(161, 141);
            this.skipItem.Name = "skipItem";
            this.skipItem.Size = new System.Drawing.Size(106, 17);
            this.skipItem.TabIndex = 6;
            this.skipItem.Text = "Enable Skip Item";
            this.skipItem.UseVisualStyleBackColor = true;
            // 
            // TCMIDForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(684, 407);
            this.Controls.Add(this.skipItem);
            this.Controls.Add(this.tb_version);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.richTextBox);
            this.Controls.Add(this.Btn_RunDownload);
            this.Controls.Add(this.outputPath);
            this.Controls.Add(this.inputPath);
            this.HelpButton = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "TCMIDForm";
            this.Text = "TCMIDForm";
            this.HelpButtonClicked += new System.ComponentModel.CancelEventHandler(this.TCMIDForm_HelpButtonClicked);
            this.Controls.SetChildIndex(this.inputPath, 0);
            this.Controls.SetChildIndex(this.outputPath, 0);
            this.Controls.SetChildIndex(this.Btn_RunDownload, 0);
            this.Controls.SetChildIndex(this.richTextBox, 0);
            this.Controls.SetChildIndex(this.label1, 0);
            this.Controls.SetChildIndex(this.tb_version, 0);
            this.Controls.SetChildIndex(this.skipItem, 0);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public MyControls.MyFileOpen inputPath;
        public MyControls.MyFileOpen outputPath;
        private MyControls.MyButtonRunDownload Btn_RunDownload;
        private System.Windows.Forms.RichTextBox richTextBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tb_version;
        public System.Windows.Forms.CheckBox skipItem;
    }
}