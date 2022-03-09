namespace CLBistDataConverter
{
    partial class frmReleaseNote
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
            this.txtReleaseNote = new System.Windows.Forms.TextBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtReleaseNote
            // 
            this.txtReleaseNote.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.txtReleaseNote.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtReleaseNote.Location = new System.Drawing.Point(12, 12);
            this.txtReleaseNote.Multiline = true;
            this.txtReleaseNote.Name = "txtReleaseNote";
            this.txtReleaseNote.ReadOnly = true;
            this.txtReleaseNote.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtReleaseNote.Size = new System.Drawing.Size(536, 249);
            this.txtReleaseNote.TabIndex = 0;
            this.txtReleaseNote.Text = " - V1.0.0 Create";
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(429, 267);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(119, 54);
            this.btnOK.TabIndex = 1;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // FrmReleaseNote
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(560, 330);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.txtReleaseNote);
            this.MinimizeBox = false;
            this.Name = "FrmReleaseNote";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ReleaseNote";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtReleaseNote;
        private System.Windows.Forms.Button btnOK;
    }
}