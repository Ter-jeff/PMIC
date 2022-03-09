namespace CLBistDataConverter
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
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.optionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.releaseNotesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.statusStripBar = new System.Windows.Forms.StatusStrip();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.toolStripStatusLabelBlank = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabelMsg = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabelBlank2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabelState = new System.Windows.Forms.ToolStripStatusLabel();
            this.gBoxInput = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnInput = new System.Windows.Forms.Button();
            this.txtDatalogFilePath = new System.Windows.Forms.TextBox();
            this.gBoxLog = new System.Windows.Forms.GroupBox();
            this.rTxtBoxLog = new System.Windows.Forms.RichTextBox();
            this.gOutputBox = new System.Windows.Forms.GroupBox();
            this.btnOutput = new System.Windows.Forms.Button();
            this.txtOutputFolder = new System.Windows.Forms.TextBox();
            this.btnRun = new System.Windows.Forms.Button();
            this.button_download = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            this.statusStripBar.SuspendLayout();
            this.gBoxInput.SuspendLayout();
            this.gBoxLog.SuspendLayout();
            this.gOutputBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.optionsToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(811, 28);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(44, 24);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(108, 26);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // optionsToolStripMenuItem
            // 
            this.optionsToolStripMenuItem.Name = "optionsToolStripMenuItem";
            this.optionsToolStripMenuItem.Size = new System.Drawing.Size(73, 24);
            this.optionsToolStripMenuItem.Text = "Options";
            this.optionsToolStripMenuItem.Visible = false;
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.releaseNotesToolStripMenuItem,
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(53, 24);
            this.helpToolStripMenuItem.Text = "Help";
            // 
            // releaseNotesToolStripMenuItem
            // 
            this.releaseNotesToolStripMenuItem.Name = "releaseNotesToolStripMenuItem";
            this.releaseNotesToolStripMenuItem.Size = new System.Drawing.Size(174, 26);
            this.releaseNotesToolStripMenuItem.Text = "ReleaseNotes";
            this.releaseNotesToolStripMenuItem.Click += new System.EventHandler(this.releaseNotesToolStripMenuItem_Click);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(174, 26);
            this.aboutToolStripMenuItem.Text = "About";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // statusStripBar
            // 
            this.statusStripBar.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.statusStripBar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripProgressBar1,
            this.toolStripStatusLabelBlank,
            this.toolStripStatusLabelMsg,
            this.toolStripStatusLabelBlank2,
            this.toolStripStatusLabelState});
            this.statusStripBar.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow;
            this.statusStripBar.Location = new System.Drawing.Point(0, 462);
            this.statusStripBar.Name = "statusStripBar";
            this.statusStripBar.Size = new System.Drawing.Size(811, 25);
            this.statusStripBar.TabIndex = 1;
            this.statusStripBar.Text = "statusStrip1";
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(100, 19);
            // 
            // toolStripStatusLabelBlank
            // 
            this.toolStripStatusLabelBlank.Name = "toolStripStatusLabelBlank";
            this.toolStripStatusLabelBlank.Size = new System.Drawing.Size(13, 20);
            this.toolStripStatusLabelBlank.Text = " ";
            // 
            // toolStripStatusLabelMsg
            // 
            this.toolStripStatusLabelMsg.Name = "toolStripStatusLabelMsg";
            this.toolStripStatusLabelMsg.Size = new System.Drawing.Size(13, 20);
            this.toolStripStatusLabelMsg.Text = " ";
            // 
            // toolStripStatusLabelBlank2
            // 
            this.toolStripStatusLabelBlank2.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.toolStripStatusLabelBlank2.Name = "toolStripStatusLabelBlank2";
            this.toolStripStatusLabelBlank2.Size = new System.Drawing.Size(13, 20);
            this.toolStripStatusLabelBlank2.Text = " ";
            // 
            // toolStripStatusLabelState
            // 
            this.toolStripStatusLabelState.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.toolStripStatusLabelState.Name = "toolStripStatusLabelState";
            this.toolStripStatusLabelState.Size = new System.Drawing.Size(13, 20);
            this.toolStripStatusLabelState.Text = " ";
            // 
            // gBoxInput
            // 
            this.gBoxInput.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gBoxInput.Controls.Add(this.label1);
            this.gBoxInput.Controls.Add(this.btnInput);
            this.gBoxInput.Controls.Add(this.txtDatalogFilePath);
            this.gBoxInput.Location = new System.Drawing.Point(12, 31);
            this.gBoxInput.Name = "gBoxInput";
            this.gBoxInput.Size = new System.Drawing.Size(787, 71);
            this.gBoxInput.TabIndex = 2;
            this.gBoxInput.TabStop = false;
            this.gBoxInput.Text = "Input";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(92, 17);
            this.label1.TabIndex = 2;
            this.label1.Text = "DataLog File:";
            // 
            // btnInput
            // 
            this.btnInput.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnInput.Location = new System.Drawing.Point(728, 23);
            this.btnInput.Name = "btnInput";
            this.btnInput.Size = new System.Drawing.Size(53, 31);
            this.btnInput.TabIndex = 1;
            this.btnInput.Text = "...";
            this.btnInput.UseVisualStyleBackColor = true;
            this.btnInput.Click += new System.EventHandler(this.btnInput_Click);
            // 
            // txtDatalogFilePath
            // 
            this.txtDatalogFilePath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDatalogFilePath.Location = new System.Drawing.Point(104, 27);
            this.txtDatalogFilePath.Name = "txtDatalogFilePath";
            this.txtDatalogFilePath.Size = new System.Drawing.Size(618, 22);
            this.txtDatalogFilePath.TabIndex = 0;
            // 
            // gBoxLog
            // 
            this.gBoxLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gBoxLog.Controls.Add(this.rTxtBoxLog);
            this.gBoxLog.Location = new System.Drawing.Point(12, 275);
            this.gBoxLog.Name = "gBoxLog";
            this.gBoxLog.Size = new System.Drawing.Size(787, 178);
            this.gBoxLog.TabIndex = 3;
            this.gBoxLog.TabStop = false;
            this.gBoxLog.Text = "Log";
            // 
            // rTxtBoxLog
            // 
            this.rTxtBoxLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.rTxtBoxLog.Location = new System.Drawing.Point(6, 21);
            this.rTxtBoxLog.Name = "rTxtBoxLog";
            this.rTxtBoxLog.Size = new System.Drawing.Size(775, 151);
            this.rTxtBoxLog.TabIndex = 0;
            this.rTxtBoxLog.Text = "";
            // 
            // gOutputBox
            // 
            this.gOutputBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gOutputBox.Controls.Add(this.btnOutput);
            this.gOutputBox.Controls.Add(this.txtOutputFolder);
            this.gOutputBox.Location = new System.Drawing.Point(12, 126);
            this.gOutputBox.Name = "gOutputBox";
            this.gOutputBox.Size = new System.Drawing.Size(787, 71);
            this.gOutputBox.TabIndex = 3;
            this.gOutputBox.TabStop = false;
            this.gOutputBox.Text = "Output";
            // 
            // btnOutput
            // 
            this.btnOutput.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOutput.Location = new System.Drawing.Point(728, 22);
            this.btnOutput.Name = "btnOutput";
            this.btnOutput.Size = new System.Drawing.Size(53, 31);
            this.btnOutput.TabIndex = 2;
            this.btnOutput.Text = "...";
            this.btnOutput.UseVisualStyleBackColor = true;
            this.btnOutput.Click += new System.EventHandler(this.btnOutput_Click);
            // 
            // txtOutputFolder
            // 
            this.txtOutputFolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtOutputFolder.Location = new System.Drawing.Point(6, 26);
            this.txtOutputFolder.Name = "txtOutputFolder";
            this.txtOutputFolder.Size = new System.Drawing.Size(716, 22);
            this.txtOutputFolder.TabIndex = 2;
            // 
            // btnRun
            // 
            this.btnRun.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRun.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnRun.Location = new System.Drawing.Point(306, 203);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(199, 66);
            this.btnRun.TabIndex = 4;
            this.btnRun.Text = "Start";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // button_download
            // 
            this.button_download.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button_download.Location = new System.Drawing.Point(577, 203);
            this.button_download.Name = "button_download";
            this.button_download.Size = new System.Drawing.Size(163, 65);
            this.button_download.TabIndex = 5;
            this.button_download.Text = "Template";
            this.button_download.UseVisualStyleBackColor = true;
            this.button_download.Click += new System.EventHandler(this.button_download_Click);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(811, 487);
            this.Controls.Add(this.button_download);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.gOutputBox);
            this.Controls.Add(this.gBoxLog);
            this.Controls.Add(this.gBoxInput);
            this.Controls.Add(this.statusStripBar);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.statusStripBar.ResumeLayout(false);
            this.statusStripBar.PerformLayout();
            this.gBoxInput.ResumeLayout(false);
            this.gBoxInput.PerformLayout();
            this.gBoxLog.ResumeLayout(false);
            this.gOutputBox.ResumeLayout(false);
            this.gOutputBox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem optionsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem releaseNotesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.StatusStrip statusStripBar;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelBlank;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelMsg;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelState;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelBlank2;
        private System.Windows.Forms.GroupBox gBoxInput;
        private System.Windows.Forms.GroupBox gBoxLog;
        private System.Windows.Forms.RichTextBox rTxtBoxLog;
        private System.Windows.Forms.GroupBox gOutputBox;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.TextBox txtDatalogFilePath;
        private System.Windows.Forms.Button btnInput;
        private System.Windows.Forms.Button btnOutput;
        private System.Windows.Forms.TextBox txtOutputFolder;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button_download;
    }
}

