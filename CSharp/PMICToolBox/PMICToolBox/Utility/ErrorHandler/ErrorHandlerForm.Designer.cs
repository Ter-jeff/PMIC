using PmicAutomation.MyControls;

namespace PmicAutomation.Utility.ErrorHandler
{
    partial class ErrorHandlerForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ErrorHandlerForm));
            this.richTextBox = new System.Windows.Forms.RichTextBox();
            this.FileOpen_OutputPath = new MyFileOpen();
            this.FileOpen_ErrorHandler = new MyFileOpen();
            this.Btn_RunDownload = new MyButtonRunDownload();
            this.chkAddErrHandler = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // richTextBox
            // 
            this.richTextBox.Location = new System.Drawing.Point(12, 255);
            this.richTextBox.Name = "richTextBox";
            this.richTextBox.Size = new System.Drawing.Size(639, 154);
            this.richTextBox.TabIndex = 3;
            this.richTextBox.Text = "";
            // 
            // FileOpen_OutputPath
            // 
            this.FileOpen_OutputPath.LebalText = "Output Path";
            this.FileOpen_OutputPath.Location = new System.Drawing.Point(12, 85);
            this.FileOpen_OutputPath.Name = "FileOpen_OutputPath";
            this.FileOpen_OutputPath.Size = new System.Drawing.Size(650, 30);
            this.FileOpen_OutputPath.TabIndex = 2;
            this.FileOpen_OutputPath.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_OutputPath_Click);
            // 
            // FileOpen_ErrorHandler
            // 
            this.FileOpen_ErrorHandler.LebalText = "Input Path";
            this.FileOpen_ErrorHandler.Location = new System.Drawing.Point(12, 34);
            this.FileOpen_ErrorHandler.Name = "FileOpen_ErrorHandler";
            this.FileOpen_ErrorHandler.Size = new System.Drawing.Size(650, 30);
            this.FileOpen_ErrorHandler.TabIndex = 1;
            this.FileOpen_ErrorHandler.ButtonTextBoxButtonClick += new System.EventHandler(this.Btn_ErrorHandler_Click);
            // 
            // Btn_RunDownload
            // 
            this.Btn_RunDownload.Location = new System.Drawing.Point(401, 136);
            this.Btn_RunDownload.Name = "Btn_RunDownload";
            this.Btn_RunDownload.RunText = "Run";
            this.Btn_RunDownload.Size = new System.Drawing.Size(250, 40);
            this.Btn_RunDownload.TabIndex = 0;
            this.Btn_RunDownload.RunButtonClick += new System.EventHandler(this.Btn_Run_Click);
            this.Btn_RunDownload.DownloadButtonClick += new System.EventHandler(this.Btn_Download_Click);
            // 
            // chkAddErrHandler
            // 
            this.chkAddErrHandler.AutoSize = true;
            this.chkAddErrHandler.Location = new System.Drawing.Point(154, 136);
            this.chkAddErrHandler.Name = "chkAddErrHandler";
            this.chkAddErrHandler.Size = new System.Drawing.Size(149, 21);
            this.chkAddErrHandler.TabIndex = 4;
            this.chkAddErrHandler.Text = "Add Error Handler ";
            this.chkAddErrHandler.UseVisualStyleBackColor = true;
            // 
            // ErrorHandlerForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(664, 421);
            this.Controls.Add(this.chkAddErrHandler);
            this.Controls.Add(this.richTextBox);
            this.Controls.Add(this.FileOpen_OutputPath);
            this.Controls.Add(this.FileOpen_ErrorHandler);
            this.Controls.Add(this.Btn_RunDownload);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.HelpButton = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ErrorHandlerForm";
            this.Text = "Check Error Handler";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private MyButtonRunDownload Btn_RunDownload;
        public MyFileOpen FileOpen_ErrorHandler;
        public MyFileOpen FileOpen_OutputPath;
        private System.Windows.Forms.RichTextBox richTextBox;
        public System.Windows.Forms.CheckBox chkAddErrHandler;
    }
}