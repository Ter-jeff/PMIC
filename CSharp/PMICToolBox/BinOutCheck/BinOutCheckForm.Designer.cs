namespace BinOutCheck
{
    partial class BinOutCheckForm
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.txtTestProgram = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnSelect = new System.Windows.Forms.Button();
            this.btnStart = new System.Windows.Forms.Button();
            this.buttonTemplate = new System.Windows.Forms.Button();
            this.btnSelectDataLog = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtDataLog = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // txtTestProgram
            // 
            this.txtTestProgram.Location = new System.Drawing.Point(79, 39);
            this.txtTestProgram.Name = "txtTestProgram";
            this.txtTestProgram.Size = new System.Drawing.Size(172, 20);
            this.txtTestProgram.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 43);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(73, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Test Program:";
            // 
            // btnSelect
            // 
            this.btnSelect.Location = new System.Drawing.Point(258, 37);
            this.btnSelect.Name = "btnSelect";
            this.btnSelect.Size = new System.Drawing.Size(47, 25);
            this.btnSelect.TabIndex = 2;
            this.btnSelect.Text = "...";
            this.btnSelect.UseVisualStyleBackColor = true;
            this.btnSelect.Click += new System.EventHandler(this.btnSelect_Click);
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(115, 127);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(66, 28);
            this.btnStart.TabIndex = 2;
            this.btnStart.Text = "Start";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // buttonTemplate
            // 
            this.buttonTemplate.Location = new System.Drawing.Point(249, 127);
            this.buttonTemplate.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.buttonTemplate.Name = "buttonTemplate";
            this.buttonTemplate.Size = new System.Drawing.Size(56, 28);
            this.buttonTemplate.TabIndex = 3;
            this.buttonTemplate.Text = "Template";
            this.buttonTemplate.UseVisualStyleBackColor = true;
            this.buttonTemplate.Click += new System.EventHandler(this.buttonTemplate_Click);
            // 
            // btnSelectDataLog
            // 
            this.btnSelectDataLog.Location = new System.Drawing.Point(258, 68);
            this.btnSelectDataLog.Name = "btnSelectDataLog";
            this.btnSelectDataLog.Size = new System.Drawing.Size(47, 25);
            this.btnSelectDataLog.TabIndex = 6;
            this.btnSelectDataLog.Text = "...";
            this.btnSelectDataLog.UseVisualStyleBackColor = true;
            this.btnSelectDataLog.Click += new System.EventHandler(this.btnSelectDataLog_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(28, 74);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(51, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "DataLog:";
            // 
            // txtDataLog
            // 
            this.txtDataLog.Location = new System.Drawing.Point(79, 70);
            this.txtDataLog.Name = "txtDataLog";
            this.txtDataLog.Size = new System.Drawing.Size(172, 20);
            this.txtDataLog.TabIndex = 4;
            // 
            // BinOutCheckForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(335, 186);
            this.Controls.Add(this.btnSelectDataLog);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtDataLog);
            this.Controls.Add(this.buttonTemplate);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.btnSelect);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtTestProgram);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "BinOutCheckForm";
            this.Text = "BinOut Check";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtTestProgram;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnSelect;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Button buttonTemplate;
        private System.Windows.Forms.Button btnSelectDataLog;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtDataLog;
    }
}

