namespace CLBistDataConverter
{
    partial class frmAbout
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
            this.btnOK = new System.Windows.Forms.Button();
            this.txtProductInfo = new System.Windows.Forms.TextBox();
            this.labTeradyne = new System.Windows.Forms.Label();
            this.labProductionInformation = new System.Windows.Forms.Label();
            this.labEnv = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(429, 341);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(119, 54);
            this.btnOK.TabIndex = 3;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // txtProductInfo
            // 
            this.txtProductInfo.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.txtProductInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtProductInfo.Location = new System.Drawing.Point(12, 106);
            this.txtProductInfo.Multiline = true;
            this.txtProductInfo.Name = "txtProductInfo";
            this.txtProductInfo.ReadOnly = true;
            this.txtProductInfo.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtProductInfo.Size = new System.Drawing.Size(536, 101);
            this.txtProductInfo.TabIndex = 2;
            this.txtProductInfo.Text = "Tool Name:\r\nDescription:";
            // 
            // labTeradyne
            // 
            this.labTeradyne.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.labTeradyne.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labTeradyne.ForeColor = System.Drawing.Color.White;
            this.labTeradyne.Location = new System.Drawing.Point(-2, -1);
            this.labTeradyne.Name = "labTeradyne";
            this.labTeradyne.Size = new System.Drawing.Size(564, 72);
            this.labTeradyne.TabIndex = 5;
            this.labTeradyne.Text = " Teradyne";
            this.labTeradyne.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // labProductionInformation
            // 
            this.labProductionInformation.AutoSize = true;
            this.labProductionInformation.Location = new System.Drawing.Point(12, 86);
            this.labProductionInformation.Name = "labProductionInformation";
            this.labProductionInformation.Size = new System.Drawing.Size(135, 17);
            this.labProductionInformation.TabIndex = 6;
            this.labProductionInformation.Text = "Product Information:";
            // 
            // labEnv
            // 
            this.labEnv.AutoSize = true;
            this.labEnv.Location = new System.Drawing.Point(12, 214);
            this.labEnv.Name = "labEnv";
            this.labEnv.Size = new System.Drawing.Size(187, 17);
            this.labEnv.TabIndex = 8;
            this.labEnv.Text = "Environmental Requirement:";
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox1.Location = new System.Drawing.Point(12, 234);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox1.Size = new System.Drawing.Size(536, 101);
            this.textBox1.TabIndex = 7;
            this.textBox1.Text = ".Net Framework4.0";
            // 
            // frmAbout
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(561, 407);
            this.Controls.Add(this.labEnv);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.labProductionInformation);
            this.Controls.Add(this.labTeradyne);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.txtProductInfo);
            this.MinimizeBox = false;
            this.Name = "frmAbout";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "About";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.TextBox txtProductInfo;
        private System.Windows.Forms.Label labTeradyne;
        private System.Windows.Forms.Label labProductionInformation;
        private System.Windows.Forms.Label labEnv;
        private System.Windows.Forms.TextBox textBox1;
    }
}