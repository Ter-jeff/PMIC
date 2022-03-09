namespace AutomationCommon.Controls
{
    partial class ProcessForm
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
            this.circularProgressBar1 = new AutomationCommon.Controls.CircularProgressBar();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.SuspendLayout();
            // 
            // circularProgressBar1
            // 
            this.circularProgressBar1.BackColor = System.Drawing.SystemColors.Control;
            this.circularProgressBar1.BarColor1 = System.Drawing.Color.Orange;
            this.circularProgressBar1.BarColor2 = System.Drawing.Color.Orange;
            this.circularProgressBar1.BarWidth = 14F;
            this.circularProgressBar1.Font = new System.Drawing.Font("Segoe UI", 15F);
            this.circularProgressBar1.ForeColor = System.Drawing.Color.DimGray;
            this.circularProgressBar1.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.ForwardDiagonal;
            this.circularProgressBar1.LineColor = System.Drawing.Color.DimGray;
            this.circularProgressBar1.LineWidth = 1;
            this.circularProgressBar1.Location = new System.Drawing.Point(142, 11);
            this.circularProgressBar1.Maximum = ((long)(100));
            this.circularProgressBar1.MinimumSize = new System.Drawing.Size(100, 100);
            this.circularProgressBar1.Name = "circularProgressBar1";
            this.circularProgressBar1.ProgressShape = AutomationCommon.Controls.CircularProgressBar._ProgressShape.Flat;
            this.circularProgressBar1.Size = new System.Drawing.Size(100, 100);
            this.circularProgressBar1.TabIndex = 0;
            this.circularProgressBar1.Text = "57";
            this.circularProgressBar1.TextMode = AutomationCommon.Controls.CircularProgressBar._TextMode.Percentage;
            this.circularProgressBar1.Value = ((long)(57));
            // 
            // statusStrip1
            // 
            this.statusStrip1.Location = new System.Drawing.Point(0, 119);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(384, 22);
            this.statusStrip1.TabIndex = 1;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // ProcessForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit;
            this.ClientSize = new System.Drawing.Size(384, 141);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.circularProgressBar1);
            this.Name = "ProcessForm";
            this.ShowIcon = false;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public AutomationCommon.Controls.CircularProgressBar circularProgressBar1;
        public System.Windows.Forms.StatusStrip statusStrip1;
    }
}