namespace PmicAutomation.Utility.TCMIDComparator
{
    partial class TCMIDComparatorForm
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
            this.inputFile1 = new PmicAutomation.MyControls.MyFileOpen();
            this.label1 = new System.Windows.Forms.Label();
            this.tb_version = new System.Windows.Forms.TextBox();
            this.inputFile2 = new PmicAutomation.MyControls.MyFileOpen();
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
            this.Btn_RunDownload.Location = new System.Drawing.Point(412, 137);
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
            this.outputPath.Location = new System.Drawing.Point(21, 80);
            this.outputPath.Name = "outputPath";
            this.outputPath.Size = new System.Drawing.Size(644, 31);
            this.outputPath.TabIndex = 1;
            this.outputPath.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_OutputPath_Click);
            // 
            // inputFile1
            // 
            this.inputFile1.LebalText = "Input File 1 (base)";
            this.inputFile1.Location = new System.Drawing.Point(21, 7);
            this.inputFile1.Name = "inputFile1";
            this.inputFile1.Size = new System.Drawing.Size(644, 31);
            this.inputFile1.TabIndex = 0;
            this.inputFile1.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_InputFile1_Click);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(21, 121);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(64, 22);
            this.label1.TabIndex = 4;
            this.label1.Text = "T/P Version";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // tb_version
            // 
            this.tb_version.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(254)));
            this.tb_version.Location = new System.Drawing.Point(161, 121);
            this.tb_version.Name = "tb_version";
            this.tb_version.Size = new System.Drawing.Size(106, 21);
            this.tb_version.TabIndex = 5;
            // 
            // inputFile2
            // 
            this.inputFile2.LebalText = "Input File 2 (new)";
            this.inputFile2.Location = new System.Drawing.Point(21, 42);
            this.inputFile2.Name = "inputFile2";
            this.inputFile2.Size = new System.Drawing.Size(644, 31);
            this.inputFile2.TabIndex = 6;
            this.inputFile2.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_InputFile2_Click);
            // 
            // skipItem
            // 
            this.skipItem.AutoSize = true;
            this.skipItem.Location = new System.Drawing.Point(161, 157);
            this.skipItem.Name = "skipItem";
            this.skipItem.Size = new System.Drawing.Size(106, 17);
            this.skipItem.TabIndex = 7;
            this.skipItem.Text = "Enable Skip Item";
            this.skipItem.UseVisualStyleBackColor = true;
            // 
            // TCMIDComparatorForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(684, 407);
            this.Controls.Add(this.skipItem);
            this.Controls.Add(this.inputFile2);
            this.Controls.Add(this.tb_version);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.richTextBox);
            this.Controls.Add(this.Btn_RunDownload);
            this.Controls.Add(this.outputPath);
            this.Controls.Add(this.inputFile1);
            this.HelpButton = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "TCMIDComparatorForm";
            this.Text = "TCMIDComparatorForm";
            this.HelpButtonClicked += new System.ComponentModel.CancelEventHandler(this.TCMIDForm_HelpButtonClicked);
            this.Controls.SetChildIndex(this.inputFile1, 0);
            this.Controls.SetChildIndex(this.outputPath, 0);
            this.Controls.SetChildIndex(this.Btn_RunDownload, 0);
            this.Controls.SetChildIndex(this.richTextBox, 0);
            this.Controls.SetChildIndex(this.label1, 0);
            this.Controls.SetChildIndex(this.tb_version, 0);
            this.Controls.SetChildIndex(this.inputFile2, 0);
            this.Controls.SetChildIndex(this.skipItem, 0);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public MyControls.MyFileOpen inputFile1;
        public MyControls.MyFileOpen outputPath;
        private MyControls.MyButtonRunDownload Btn_RunDownload;
        private System.Windows.Forms.RichTextBox richTextBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tb_version;
        public MyControls.MyFileOpen inputFile2;
        public System.Windows.Forms.CheckBox skipItem;
    }
}