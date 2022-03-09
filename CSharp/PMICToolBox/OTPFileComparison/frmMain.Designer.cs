
namespace OTPFileComparison
{
    partial class frmMain
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
            this.btnLoad = new System.Windows.Forms.Button();
            this.txtOutput = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnOutput = new System.Windows.Forms.Button();
            this.btnRun = new System.Windows.Forms.Button();
            this.comparisonOTPTbl1 = new CommonLib.Controls.ComparisonOTPTbl();
            this.SuspendLayout();
            // 
            // btnLoad
            // 
            this.btnLoad.Location = new System.Drawing.Point(11, 12);
            this.btnLoad.Margin = new System.Windows.Forms.Padding(2);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(92, 40);
            this.btnLoad.TabIndex = 3;
            this.btnLoad.Text = "Load OTP Compare File";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
            // 
            // txtOutput
            // 
            this.txtOutput.Location = new System.Drawing.Point(177, 22);
            this.txtOutput.Margin = new System.Windows.Forms.Padding(2);
            this.txtOutput.Name = "txtOutput";
            this.txtOutput.Size = new System.Drawing.Size(491, 21);
            this.txtOutput.TabIndex = 11;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(127, 24);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(43, 15);
            this.label1.TabIndex = 12;
            this.label1.Text = "Output";
            // 
            // btnOutput
            // 
            this.btnOutput.Location = new System.Drawing.Point(672, 17);
            this.btnOutput.Margin = new System.Windows.Forms.Padding(2);
            this.btnOutput.Name = "btnOutput";
            this.btnOutput.Size = new System.Drawing.Size(44, 30);
            this.btnOutput.TabIndex = 14;
            this.btnOutput.Text = "...";
            this.btnOutput.UseVisualStyleBackColor = true;
            this.btnOutput.Click += new System.EventHandler(this.btnOutput_Click);
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(306, 353);
            this.btnRun.Margin = new System.Windows.Forms.Padding(2);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(100, 51);
            this.btnRun.TabIndex = 16;
            this.btnRun.Text = "Run";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // comparisonOTPTbl1
            // 
            this.comparisonOTPTbl1.AutoScroll = true;
            this.comparisonOTPTbl1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.comparisonOTPTbl1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel, ((byte)(134)));
            this.comparisonOTPTbl1.Location = new System.Drawing.Point(6, 56);
            this.comparisonOTPTbl1.Margin = new System.Windows.Forms.Padding(2);
            this.comparisonOTPTbl1.Name = "comparisonOTPTbl1";
            this.comparisonOTPTbl1.Size = new System.Drawing.Size(707, 293);
            this.comparisonOTPTbl1.TabIndex = 17;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(720, 435);
            this.Controls.Add(this.comparisonOTPTbl1);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.btnOutput);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtOutput);
            this.Controls.Add(this.btnLoad);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "OTPFileComparison";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.Controls.SetChildIndex(this.btnLoad, 0);
            this.Controls.SetChildIndex(this.txtOutput, 0);
            this.Controls.SetChildIndex(this.label1, 0);
            this.Controls.SetChildIndex(this.btnOutput, 0);
            this.Controls.SetChildIndex(this.btnRun, 0);
            this.Controls.SetChildIndex(this.comparisonOTPTbl1, 0);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnLoad;
        private System.Windows.Forms.TextBox txtOutput;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnOutput;
        private System.Windows.Forms.Button btnRun;
        private CommonLib.Controls.ComparisonOTPTbl comparisonOTPTbl1;
    }
}

