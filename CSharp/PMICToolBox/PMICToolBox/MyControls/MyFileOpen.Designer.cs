using System;
namespace PmicAutomation.MyControls
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
            this.ButtonTextBox = new MyButtonTextBox();
            this.SuspendLayout();
            // 
            // Label
            // 
            this.Label.Font = new System.Drawing.Font("Microsoft JhengHei UI", 9F);
            this.Label.Location = new System.Drawing.Point(0, 0);
            this.Label.Name = "Label";
            this.Label.Size = new System.Drawing.Size(120, 30);
            this.Label.TabIndex = 5;
            this.Label.Text = "Label";
            this.Label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // ButtonTextBox
            // 
            this.ButtonTextBox.Font = new System.Drawing.Font("Microsoft JhengHei UI", 9F);
            this.ButtonTextBox.Location = new System.Drawing.Point(140, 4);
            this.ButtonTextBox.Name = "ButtonTextBox";
            this.ButtonTextBox.Size = new System.Drawing.Size(500, 23);
            this.ButtonTextBox.TabIndex = 0;
            // 
            // FileOpen
            // 
            this.Controls.Add(this.Label);
            this.Controls.Add(this.ButtonTextBox);
            this.Name = "MyFileOpen";
            this.Size = new System.Drawing.Size(650, 30);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.Label Label;
        public MyButtonTextBox ButtonTextBox;
    }
}
