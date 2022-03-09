
namespace CommonLib.Controls
{
    partial class ComparisonOTPRow
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.txtFile = new System.Windows.Forms.TextBox();
            this.chkHC = new System.Windows.Forms.CheckBox();
            this.chkVC = new System.Windows.Forms.CheckBox();
            this.btndel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtFile
            // 
            this.txtFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtFile.BackColor = System.Drawing.SystemColors.Window;
            this.txtFile.Location = new System.Drawing.Point(38, 1);
            this.txtFile.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtFile.Name = "txtFile";
            this.txtFile.ReadOnly = true;
            this.txtFile.Size = new System.Drawing.Size(527, 21);
            this.txtFile.TabIndex = 0;
            // 
            // chkHC
            // 
            this.chkHC.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.chkHC.AutoSize = true;
            this.chkHC.Checked = true;
            this.chkHC.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkHC.Location = new System.Drawing.Point(592, 4);
            this.chkHC.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.chkHC.Name = "chkHC";
            this.chkHC.Size = new System.Drawing.Size(15, 14);
            this.chkHC.TabIndex = 1;
            this.chkHC.UseVisualStyleBackColor = true;
            // 
            // chkVC
            // 
            this.chkVC.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.chkVC.AutoSize = true;
            this.chkVC.Checked = true;
            this.chkVC.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkVC.Location = new System.Drawing.Point(648, 4);
            this.chkVC.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.chkVC.Name = "chkVC";
            this.chkVC.Size = new System.Drawing.Size(15, 14);
            this.chkVC.TabIndex = 2;
            this.chkVC.UseVisualStyleBackColor = true;
            // 
            // btndel
            // 
            this.btndel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel, ((byte)(134)));
            this.btndel.Location = new System.Drawing.Point(0, 0);
            this.btndel.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btndel.Name = "btndel";
            this.btndel.Size = new System.Drawing.Size(33, 22);
            this.btndel.TabIndex = 3;
            this.btndel.Text = "del";
            this.btndel.UseVisualStyleBackColor = true;
            this.btndel.Click += new System.EventHandler(this.btndel_Click);
            // 
            // ComparisonOTPRow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoScroll = true;
            this.Controls.Add(this.btndel);
            this.Controls.Add(this.chkVC);
            this.Controls.Add(this.chkHC);
            this.Controls.Add(this.txtFile);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "ComparisonOTPRow";
            this.Size = new System.Drawing.Size(684, 24);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtFile;
        private System.Windows.Forms.CheckBox chkHC;
        private System.Windows.Forms.CheckBox chkVC;
        private System.Windows.Forms.Button btndel;
    }
}
