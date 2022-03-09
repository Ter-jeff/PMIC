
namespace AlarmChekc
{
    partial class AlarmCheck
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.B_IGXLTestProgram = new System.Windows.Forms.Button();
            this.T_TestProgram = new System.Windows.Forms.TextBox();
            this.l_TestProgram = new System.Windows.Forms.Label();
            this.l_OutputPath = new System.Windows.Forms.Label();
            this.T_OutputPath = new System.Windows.Forms.TextBox();
            this.B_OutputPath = new System.Windows.Forms.Button();
            this.B_Strart = new System.Windows.Forms.Button();
            this.buttonTemplate = new System.Windows.Forms.Button();
            this.chkAlarmCheck = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // B_IGXLTestProgram
            // 
            this.B_IGXLTestProgram.Location = new System.Drawing.Point(258, 21);
            this.B_IGXLTestProgram.Name = "B_IGXLTestProgram";
            this.B_IGXLTestProgram.Size = new System.Drawing.Size(64, 20);
            this.B_IGXLTestProgram.TabIndex = 0;
            this.B_IGXLTestProgram.Text = "...";
            this.B_IGXLTestProgram.UseVisualStyleBackColor = true;
            this.B_IGXLTestProgram.Click += new System.EventHandler(this.B_IGXLTestProgram_Click);
            // 
            // T_TestProgram
            // 
            this.T_TestProgram.Location = new System.Drawing.Point(78, 22);
            this.T_TestProgram.Name = "T_TestProgram";
            this.T_TestProgram.Size = new System.Drawing.Size(175, 20);
            this.T_TestProgram.TabIndex = 1;
            // 
            // l_TestProgram
            // 
            this.l_TestProgram.AutoSize = true;
            this.l_TestProgram.Location = new System.Drawing.Point(10, 24);
            this.l_TestProgram.Name = "l_TestProgram";
            this.l_TestProgram.Size = new System.Drawing.Size(67, 13);
            this.l_TestProgram.TabIndex = 2;
            this.l_TestProgram.Text = "TestProgram";
            // 
            // l_OutputPath
            // 
            this.l_OutputPath.AutoSize = true;
            this.l_OutputPath.Location = new System.Drawing.Point(10, 58);
            this.l_OutputPath.Name = "l_OutputPath";
            this.l_OutputPath.Size = new System.Drawing.Size(61, 13);
            this.l_OutputPath.TabIndex = 2;
            this.l_OutputPath.Text = "OutputPath";
            // 
            // T_OutputPath
            // 
            this.T_OutputPath.Location = new System.Drawing.Point(78, 55);
            this.T_OutputPath.Name = "T_OutputPath";
            this.T_OutputPath.Size = new System.Drawing.Size(175, 20);
            this.T_OutputPath.TabIndex = 1;
            // 
            // B_OutputPath
            // 
            this.B_OutputPath.Location = new System.Drawing.Point(258, 58);
            this.B_OutputPath.Name = "B_OutputPath";
            this.B_OutputPath.Size = new System.Drawing.Size(64, 20);
            this.B_OutputPath.TabIndex = 0;
            this.B_OutputPath.Text = "...";
            this.B_OutputPath.UseVisualStyleBackColor = true;
            this.B_OutputPath.Click += new System.EventHandler(this.B_OutputPath_Click);
            // 
            // B_Strart
            // 
            this.B_Strart.Location = new System.Drawing.Point(178, 147);
            this.B_Strart.Name = "B_Strart";
            this.B_Strart.Size = new System.Drawing.Size(66, 28);
            this.B_Strart.TabIndex = 0;
            this.B_Strart.Text = "Start";
            this.B_Strart.UseVisualStyleBackColor = true;
            this.B_Strart.Click += new System.EventHandler(this.B_Strart_Click);
            // 
            // buttonTemplate
            // 
            this.buttonTemplate.Location = new System.Drawing.Point(249, 147);
            this.buttonTemplate.Margin = new System.Windows.Forms.Padding(2);
            this.buttonTemplate.Name = "buttonTemplate";
            this.buttonTemplate.Size = new System.Drawing.Size(73, 28);
            this.buttonTemplate.TabIndex = 4;
            this.buttonTemplate.Text = "Template";
            this.buttonTemplate.UseVisualStyleBackColor = true;
            this.buttonTemplate.Click += new System.EventHandler(this.buttonTemplate_Click);
            // 
            // chkAlarmCheck
            // 
            this.chkAlarmCheck.AutoSize = true;
            this.chkAlarmCheck.Location = new System.Drawing.Point(78, 92);
            this.chkAlarmCheck.Margin = new System.Windows.Forms.Padding(2);
            this.chkAlarmCheck.Name = "chkAlarmCheck";
            this.chkAlarmCheck.Size = new System.Drawing.Size(176, 17);
            this.chkAlarmCheck.TabIndex = 5;
            this.chkAlarmCheck.Text = "Automatical Add Alarm Function";
            this.chkAlarmCheck.UseVisualStyleBackColor = true;
            // 
            // AlarmCheck
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(335, 186);
            this.Controls.Add(this.chkAlarmCheck);
            this.Controls.Add(this.buttonTemplate);
            this.Controls.Add(this.l_OutputPath);
            this.Controls.Add(this.l_TestProgram);
            this.Controls.Add(this.T_OutputPath);
            this.Controls.Add(this.T_TestProgram);
            this.Controls.Add(this.B_Strart);
            this.Controls.Add(this.B_OutputPath);
            this.Controls.Add(this.B_IGXLTestProgram);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "AlarmCheck";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "AlarmCheck";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button B_IGXLTestProgram;
        private System.Windows.Forms.TextBox T_TestProgram;
        private System.Windows.Forms.Label l_TestProgram;
        private System.Windows.Forms.Label l_OutputPath;
        private System.Windows.Forms.TextBox T_OutputPath;
        private System.Windows.Forms.Button B_OutputPath;
        private System.Windows.Forms.Button B_Strart;
        private System.Windows.Forms.Button buttonTemplate;
        public System.Windows.Forms.CheckBox chkAlarmCheck;
    }
}

