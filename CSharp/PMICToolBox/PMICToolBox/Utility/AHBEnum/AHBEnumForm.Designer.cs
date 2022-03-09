using PmicAutomation.MyControls;

namespace PmicAutomation.Utility.AHBEnum
{
    public partial class AhbEnum
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AhbEnum));
            this.FileOpen_AhbRegister = new PmicAutomation.MyControls.MyFileOpen();
            this.FileOpen_OutputPath = new PmicAutomation.MyControls.MyFileOpen();
            this.Btn_RunDownload = new PmicAutomation.MyControls.MyButtonRunDownload();
            this.richTextBox = new System.Windows.Forms.RichTextBox();
            this.radioButton_RegNameAndFieldName = new System.Windows.Forms.RadioButton();
            this.radioButton_FieldNameOnly = new System.Windows.Forms.RadioButton();
            this.label_MaxBitWidth = new System.Windows.Forms.Label();
            this.textBox_MaxBitWidth = new System.Windows.Forms.TextBox();
            this.label_FieldName = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // FileOpen_AhbRegister
            // 
            this.FileOpen_AhbRegister.LebalText = "AhbRegister";
            this.FileOpen_AhbRegister.Location = new System.Drawing.Point(50, 20);
            this.FileOpen_AhbRegister.Name = "FileOpen_AhbRegister";
            this.FileOpen_AhbRegister.Size = new System.Drawing.Size(640, 30);
            this.FileOpen_AhbRegister.TabIndex = 0;
            this.FileOpen_AhbRegister.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_AhbRegister_Click);
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
            this.Btn_RunDownload.Location = new System.Drawing.Point(440, 130);
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
            this.richTextBox.Location = new System.Drawing.Point(30, 207);
            this.richTextBox.Margin = new System.Windows.Forms.Padding(21);
            this.richTextBox.Name = "richTextBox";
            this.richTextBox.Size = new System.Drawing.Size(680, 260);
            this.richTextBox.TabIndex = 3;
            this.richTextBox.Text = "";
            // 
            // radioButton_RegNameAndFieldName
            // 
            this.radioButton_RegNameAndFieldName.AutoSize = true;
            this.radioButton_RegNameAndFieldName.Checked = true;
            this.radioButton_RegNameAndFieldName.Location = new System.Drawing.Point(190, 138);
            this.radioButton_RegNameAndFieldName.Name = "radioButton_RegNameAndFieldName";
            this.radioButton_RegNameAndFieldName.Size = new System.Drawing.Size(162, 19);
            this.radioButton_RegNameAndFieldName.TabIndex = 4;
            this.radioButton_RegNameAndFieldName.TabStop = true;
            this.radioButton_RegNameAndFieldName.Text = "Reg Name + Field Name";
            this.radioButton_RegNameAndFieldName.UseVisualStyleBackColor = true;
            // 
            // radioButton_FieldNameOnly
            // 
            this.radioButton_FieldNameOnly.AutoSize = true;
            this.radioButton_FieldNameOnly.Location = new System.Drawing.Point(190, 159);
            this.radioButton_FieldNameOnly.Name = "radioButton_FieldNameOnly";
            this.radioButton_FieldNameOnly.Size = new System.Drawing.Size(114, 19);
            this.radioButton_FieldNameOnly.TabIndex = 4;
            this.radioButton_FieldNameOnly.Text = "Field Name only";
            this.radioButton_FieldNameOnly.UseVisualStyleBackColor = true;
            // 
            // label_MaxBitWidth
            // 
            this.label_MaxBitWidth.AutoSize = true;
            this.label_MaxBitWidth.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.label_MaxBitWidth.Location = new System.Drawing.Point(47, 115);
            this.label_MaxBitWidth.Name = "label_MaxBitWidth";
            this.label_MaxBitWidth.Size = new System.Drawing.Size(82, 15);
            this.label_MaxBitWidth.TabIndex = 5;
            this.label_MaxBitWidth.Text = "Max Bit Width";
            // 
            // textBox_MaxBitWidth
            // 
            this.textBox_MaxBitWidth.Location = new System.Drawing.Point(50, 138);
            this.textBox_MaxBitWidth.Name = "textBox_MaxBitWidth";
            this.textBox_MaxBitWidth.Size = new System.Drawing.Size(98, 21);
            this.textBox_MaxBitWidth.TabIndex = 6;
            this.textBox_MaxBitWidth.Text = "8";
            // 
            // label_FieldName
            // 
            this.label_FieldName.AutoSize = true;
            this.label_FieldName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.label_FieldName.Location = new System.Drawing.Point(190, 115);
            this.label_FieldName.Name = "label_FieldName";
            this.label_FieldName.Size = new System.Drawing.Size(68, 15);
            this.label_FieldName.TabIndex = 5;
            this.label_FieldName.Text = "FieldName";
            // 
            // AhbEnum
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(740, 497);
            this.Controls.Add(this.textBox_MaxBitWidth);
            this.Controls.Add(this.label_FieldName);
            this.Controls.Add(this.label_MaxBitWidth);
            this.Controls.Add(this.radioButton_FieldNameOnly);
            this.Controls.Add(this.radioButton_RegNameAndFieldName);
            this.Controls.Add(this.FileOpen_OutputPath);
            this.Controls.Add(this.FileOpen_AhbRegister);
            this.Controls.Add(this.Btn_RunDownload);
            this.Controls.Add(this.richTextBox);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.HelpButton = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AhbEnum";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Ahb Generator";
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        public MyFileOpen FileOpen_AhbRegister;
        public MyFileOpen FileOpen_OutputPath;
        private MyButtonRunDownload Btn_RunDownload;
        private System.Windows.Forms.RichTextBox richTextBox;
        public System.Windows.Forms.RadioButton radioButton_RegNameAndFieldName;
        public System.Windows.Forms.RadioButton radioButton_FieldNameOnly;
        private System.Windows.Forms.Label label_MaxBitWidth;
        public System.Windows.Forms.TextBox textBox_MaxBitWidth;
        private System.Windows.Forms.Label label_FieldName;
    }
}