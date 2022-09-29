namespace PmicAutogen
{
    partial class VddRefForm
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label_Cancel = new System.Windows.Forms.Label();
            this.label_Ok = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.DomainName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Voltage = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ReferencePin = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dataGridView1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(10, 10);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(626, 212);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AllowUserToResizeColumns = false;
            this.dataGridView1.AllowUserToResizeRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.DomainName,
            this.Voltage,
            this.ReferencePin});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(3, 16);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(620, 193);
            this.dataGridView1.TabIndex = 6;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label_Cancel);
            this.panel1.Controls.Add(this.label_Ok);
            this.panel1.Controls.Add(this.btnCancel);
            this.panel1.Controls.Add(this.btnOK);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(10, 222);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(626, 57);
            this.panel1.TabIndex = 7;
            // 
            // label_Cancel
            // 
            this.label_Cancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.label_Cancel.AutoSize = true;
            this.label_Cancel.Location = new System.Drawing.Point(392, 15);
            this.label_Cancel.Name = "label_Cancel";
            this.label_Cancel.Size = new System.Drawing.Size(143, 26);
            this.label_Cancel.TabIndex = 10;
            this.label_Cancel.Text = "The domain value will \r\nfollow IO Levels sheet value.";
            // 
            // label_Ok
            // 
            this.label_Ok.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.label_Ok.AutoSize = true;
            this.label_Ok.Location = new System.Drawing.Point(172, 15);
            this.label_Ok.Name = "label_Ok";
            this.label_Ok.Size = new System.Drawing.Size(120, 26);
            this.label_Ok.TabIndex = 9;
            this.label_Ok.Text = "Base on refrernce pin to\r\ndefine doman value";
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(312, 12);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 32);
            this.btnCancel.TabIndex = 8;
            this.btnCancel.Text = "Cancel Ref";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnOK.Location = new System.Drawing.Point(95, 12);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 32);
            this.btnOK.TabIndex = 7;
            this.btnOK.Text = "Ref Vdd";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // DomainName
            // 
            this.DomainName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.DomainName.FillWeight = 101.5228F;
            this.DomainName.HeaderText = "DomainName";
            this.DomainName.Name = "DomainName";
            this.DomainName.ReadOnly = true;
            this.DomainName.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.DomainName.Width = 160;
            // 
            // Voltage
            // 
            this.Voltage.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.Voltage.HeaderText = "Voltage";
            this.Voltage.Name = "Voltage";
            this.Voltage.ReadOnly = true;
            // 
            // ReferencePin
            // 
            this.ReferencePin.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.ReferencePin.FillWeight = 98.47716F;
            this.ReferencePin.HeaderText = "Reference Pin";
            this.ReferencePin.Name = "ReferencePin";
            this.ReferencePin.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.ReferencePin.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // VDDRefForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(646, 289);
            this.ControlBox = false;
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.panel1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "VddRefForm";
            this.Padding = new System.Windows.Forms.Padding(10);
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "\"IO_Level_Sheet Domain\" Define Reference Pin";
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label_Cancel;
        private System.Windows.Forms.Label label_Ok;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn DomainName;
        private System.Windows.Forms.DataGridViewTextBoxColumn Voltage;
        private System.Windows.Forms.DataGridViewComboBoxColumn ReferencePin;
    }
}