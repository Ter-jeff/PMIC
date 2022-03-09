using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using CommonLib.Properties;

namespace CommonLib.Controls
{
    public sealed class MyButtonTextBox : TextBox
    {
        public MyButtonTextBox()
        {
            InitializeComponent();

            Button = new Button { Cursor = Cursors.Default };
            Button.SizeChanged += (o, e) => OnResize(e);
            Button.BackgroundImage = Resources.shell32_3191.ToBitmap();
            Button.BackgroundImageLayout = ImageLayout.Zoom;
            Button.TabStop = false;
            Button.FlatStyle = FlatStyle.Flat;
            Button.FlatAppearance.BorderSize = 0;
            Controls.Add(Button);

            Font = new Font("Microsoft JhengHei UI", 9F);
            Size = new Size(500, 30);
        }

        public Button Button
        {
            get;
            set;
        }

        public event EventHandler ButtonClick
        {
            add { Button.Click += value; }
            remove { Button.Click -= value; }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            Button.Size = new Size(ClientSize.Height + 4, ClientSize.Height + 4);
            Button.Location = new Point(ClientSize.Width - ClientSize.Height - 2, -2);
            SendMessage(Handle, 0xd3, (IntPtr)2, (IntPtr)(Button.Width << 16));
        }

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wp, IntPtr lp);

        private void InitializeComponent()
        {
            SuspendLayout();
            ResumeLayout(false);
        }
    }
}