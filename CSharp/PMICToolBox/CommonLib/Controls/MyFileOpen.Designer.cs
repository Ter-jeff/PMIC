using System;

namespace CommonLib.Controls
{
    partial class MyFileOpen
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public string LebalText { get { return Label.Text; } set { Label.Text = value; } }
        public event EventHandler ButtonTextBoxButtonClick { add { ButtonTextBox.ButtonClick += value; } remove { ButtonTextBox.ButtonClick -= value; } }

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
            this.Label = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.ButtonTextBox = new CommonLib.Controls.MyButtonTextBox();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // Label
            // 
            this.Label.CausesValidation = false;
            this.Label.Dock = System.Windows.Forms.DockStyle.Left;
            this.Label.Font = new System.Drawing.Font("Microsoft JhengHei UI", 9F);
            this.Label.Location = new System.Drawing.Point(10, 10);
            this.Label.Name = "Label";
            this.Label.Size = new System.Drawing.Size(130, 23);
            this.Label.TabIndex = 5;
            this.Label.Text = "Label";
            this.Label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.ButtonTextBox);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(140, 10);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(500, 23);
            this.panel1.TabIndex = 9;
            // 
            // ButtonTextBox
            // 
            this.ButtonTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ButtonTextBox.Font = new System.Drawing.Font("Microsoft JhengHei UI", 9F);
            this.ButtonTextBox.ImeMode = System.Windows.Forms.ImeMode.On;
            this.ButtonTextBox.Location = new System.Drawing.Point(0, 0);
            this.ButtonTextBox.Name = "ButtonTextBox";
            this.ButtonTextBox.Size = new System.Drawing.Size(500, 23);
            this.ButtonTextBox.TabIndex = 9;
            // 
            // MyFileOpen
            // 
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.Label);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.Name = "MyFileOpen";
            this.Padding = new System.Windows.Forms.Padding(10);
            this.Size = new System.Drawing.Size(650, 43);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.Label Label;
        private System.Windows.Forms.Panel panel1;
        public MyButtonTextBox ButtonTextBox;
    }
}
