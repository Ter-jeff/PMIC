using PmicAutomation.MyControls;

namespace PmicAutomation.Utility.DatalogComparator
{
    partial class DatalogComparatorForm
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
            this.FileOpen_BaseLog = new PmicAutomation.MyControls.MyFileOpen();
            this.FileOpen_CompareLog = new PmicAutomation.MyControls.MyFileOpen();
            this.FileOpen_Output = new PmicAutomation.MyControls.MyFileOpen();
            this.btnRun = new System.Windows.Forms.Button();
            this.buttonTemplate = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // FileOpen_BaseLog
            // 
            this.FileOpen_BaseLog.LebalText = "Base Datalog:";
            this.FileOpen_BaseLog.Location = new System.Drawing.Point(12, 35);
            this.FileOpen_BaseLog.Name = "FileOpen_BaseLog";
            this.FileOpen_BaseLog.Size = new System.Drawing.Size(646, 42);
            this.FileOpen_BaseLog.TabIndex = 0;
            this.FileOpen_BaseLog.ButtonTextBoxButtonClick += new System.EventHandler(this.FileOpen_BaseLog_ButtonTextBoxButtonClick);
            // 
            // FileOpen_CompareLog
            // 
            this.FileOpen_CompareLog.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FileOpen_CompareLog.LebalText = "Compare Data";
            this.FileOpen_CompareLog.Location = new System.Drawing.Point(12, 83);
            this.FileOpen_CompareLog.Name = "FileOpen_CompareLog";
            this.FileOpen_CompareLog.Size = new System.Drawing.Size(648, 35);
            this.FileOpen_CompareLog.TabIndex = 0;
            this.FileOpen_CompareLog.ButtonTextBoxButtonClick += new System.EventHandler(this.FileOpen_BaseLog_ButtonTextBoxButtonClick);
            // 
            // FileOpen_Output
            // 
            this.FileOpen_Output.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FileOpen_Output.LebalText = "Output";
            this.FileOpen_Output.Location = new System.Drawing.Point(12, 134);
            this.FileOpen_Output.Name = "FileOpen_Output";
            this.FileOpen_Output.Size = new System.Drawing.Size(648, 35);
            this.FileOpen_Output.TabIndex = 0;
            this.FileOpen_Output.ButtonTextBoxButtonClick += new System.EventHandler(this.FileOpen_Output_ButtonTextBoxButtonClick);
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(297, 175);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(113, 38);
            this.btnRun.TabIndex = 1;
            this.btnRun.Text = "Run";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // buttonTemplate
            // 
            this.buttonTemplate.Location = new System.Drawing.Point(545, 175);
            this.buttonTemplate.Name = "buttonTemplate";
            this.buttonTemplate.Size = new System.Drawing.Size(113, 38);
            this.buttonTemplate.TabIndex = 2;
            this.buttonTemplate.Text = "Template";
            this.buttonTemplate.UseVisualStyleBackColor = true;
            this.buttonTemplate.Click += new System.EventHandler(this.buttonTemplate_Click);
            // 
            // DatalogComparatorForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(670, 252);
            this.Controls.Add(this.buttonTemplate);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.FileOpen_Output);
            this.Controls.Add(this.FileOpen_CompareLog);
            this.Controls.Add(this.FileOpen_BaseLog);
            this.MaximizeBox = false;
            this.Name = "DatalogComparatorForm";
            this.Text = "DatalogComparator";
            this.Controls.SetChildIndex(this.FileOpen_BaseLog, 0);
            this.Controls.SetChildIndex(this.FileOpen_CompareLog, 0);
            this.Controls.SetChildIndex(this.FileOpen_Output, 0);
            this.Controls.SetChildIndex(this.btnRun, 0);
            this.Controls.SetChildIndex(this.buttonTemplate, 0);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private MyFileOpen FileOpen_BaseLog;
        private MyFileOpen FileOpen_CompareLog;
        private MyFileOpen FileOpen_Output;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.RichTextBox richTextBox;
        private System.Windows.Forms.Button buttonTemplate;
    }
}