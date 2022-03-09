namespace PinNameRemoval
{
    partial class PinNameRemovalForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PinNameRemovalForm));
            this.label1 = new System.Windows.Forms.Label();
            this.tbInputPath = new System.Windows.Forms.TextBox();
            this.btnSelectInput = new System.Windows.Forms.Button();
            this.btnRun = new System.Windows.Forms.Button();
            this.btnSelectPinmap = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.tbPinmapPath = new System.Windows.Forms.TextBox();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.btnDeletePins = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label3 = new System.Windows.Forms.Label();
            this.tbOutputPath = new System.Windows.Forms.TextBox();
            this.tbPinName = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.btnSelectOutput = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "Input Folder";
            // 
            // tbInputPath
            // 
            this.tbInputPath.Location = new System.Drawing.Point(122, 15);
            this.tbInputPath.Name = "tbInputPath";
            this.tbInputPath.Size = new System.Drawing.Size(281, 21);
            this.tbInputPath.TabIndex = 1;
            // 
            // btnSelectInput
            // 
            this.btnSelectInput.Location = new System.Drawing.Point(409, 15);
            this.btnSelectInput.Name = "btnSelectInput";
            this.btnSelectInput.Size = new System.Drawing.Size(36, 25);
            this.btnSelectInput.TabIndex = 2;
            this.btnSelectInput.Text = "...";
            this.btnSelectInput.UseVisualStyleBackColor = true;
            this.btnSelectInput.Click += new System.EventHandler(this.btnSelectInput_Click);
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(244, 141);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(78, 58);
            this.btnRun.TabIndex = 3;
            this.btnRun.Text = "Run";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // btnSelectPinmap
            // 
            this.btnSelectPinmap.Location = new System.Drawing.Point(409, 46);
            this.btnSelectPinmap.Name = "btnSelectPinmap";
            this.btnSelectPinmap.Size = new System.Drawing.Size(36, 25);
            this.btnSelectPinmap.TabIndex = 4;
            this.btnSelectPinmap.Text = "...";
            this.btnSelectPinmap.UseVisualStyleBackColor = true;
            this.btnSelectPinmap.Click += new System.EventHandler(this.btnSelectPinmap_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 49);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 15);
            this.label2.TabIndex = 5;
            this.label2.Text = "Pinmap";
            // 
            // tbPinmapPath
            // 
            this.tbPinmapPath.Location = new System.Drawing.Point(122, 46);
            this.tbPinmapPath.Name = "tbPinmapPath";
            this.tbPinmapPath.Size = new System.Drawing.Size(281, 21);
            this.tbPinmapPath.TabIndex = 6;
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.HorizontalScrollbar = true;
            this.listBox1.ItemHeight = 15;
            this.listBox1.Location = new System.Drawing.Point(12, 205);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(433, 184);
            this.listBox1.TabIndex = 7;
            // 
            // btnDeletePins
            // 
            this.btnDeletePins.Location = new System.Drawing.Point(328, 141);
            this.btnDeletePins.Name = "btnDeletePins";
            this.btnDeletePins.Size = new System.Drawing.Size(75, 58);
            this.btnDeletePins.TabIndex = 8;
            this.btnDeletePins.Text = "Delete Pins";
            this.btnDeletePins.UseVisualStyleBackColor = true;
            this.btnDeletePins.Visible = false;
            this.btnDeletePins.Click += new System.EventHandler(this.btnDeletePins_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 402);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(433, 31);
            this.progressBar1.TabIndex = 9;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 81);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(81, 15);
            this.label3.TabIndex = 10;
            this.label3.Text = "Output Folder";
            // 
            // tbOutputPath
            // 
            this.tbOutputPath.Location = new System.Drawing.Point(122, 78);
            this.tbOutputPath.Name = "tbOutputPath";
            this.tbOutputPath.Size = new System.Drawing.Size(281, 21);
            this.tbOutputPath.TabIndex = 11;
            // 
            // tbPinName
            // 
            this.tbPinName.Location = new System.Drawing.Point(122, 110);
            this.tbPinName.Name = "tbPinName";
            this.tbPinName.Size = new System.Drawing.Size(281, 21);
            this.tbPinName.TabIndex = 12;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 113);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(58, 15);
            this.label4.TabIndex = 13;
            this.label4.Text = "Pin Name";
            // 
            // btnSelectOutput
            // 
            this.btnSelectOutput.Location = new System.Drawing.Point(409, 77);
            this.btnSelectOutput.Name = "btnSelectOutput";
            this.btnSelectOutput.Size = new System.Drawing.Size(36, 26);
            this.btnSelectOutput.TabIndex = 14;
            this.btnSelectOutput.Text = "...";
            this.btnSelectOutput.UseVisualStyleBackColor = true;
            this.btnSelectOutput.Click += new System.EventHandler(this.btnSelectOutput_Click);
            // 
            // PinNameRemovalForm
            // 
            this.AcceptButton = this.btnRun;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(457, 434);
            this.Controls.Add(this.btnSelectOutput);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.tbPinName);
            this.Controls.Add(this.tbOutputPath);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnDeletePins);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.tbPinmapPath);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnSelectPinmap);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.btnSelectInput);
            this.Controls.Add(this.tbInputPath);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Microsoft JhengHei", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "PinNameRemovalForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Remove Pin Name v.[version]";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbInputPath;
        private System.Windows.Forms.Button btnSelectInput;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.Button btnSelectPinmap;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbPinmapPath;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Button btnDeletePins;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tbOutputPath;
        private System.Windows.Forms.TextBox tbPinName;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnSelectOutput;
    }
}

