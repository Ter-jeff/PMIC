using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace CommonLib.Controls
{
    public class CheckedListBoxForm : Form
    {
        private MyCheckededListBox _myCheckededListBox;
        private List<string> _checkedEnables;

        public CheckedListBoxForm()
        {
            InitializeComponent();
            Font = new Font("Microsoft Sans Serif", 9F);
        }

        public void BindingCheckedListBox(List<string> enables, List<string> checkedList)
        {
            foreach (var enable in enables)
            {
                if (checkedList == null || checkedList.Exists(x => x.Equals(enable, StringComparison.CurrentCultureIgnoreCase)))
                    _myCheckededListBox.Items.Add(enable, true);
                else
                    _myCheckededListBox.Items.Add(enable, false);
            }

            AdjustSize(_myCheckededListBox);
        }

        private void AdjustSize(MyCheckededListBox checkedListBox)
        {
            int h = checkedListBox.ItemHeight * checkedListBox.Items.Count;
            checkedListBox.Height = h + checkedListBox.Height - checkedListBox.ClientSize.Height;
            Height = checkedListBox.Height + 70;
        }

        private void InitializeComponent()
        {
            _myCheckededListBox = new MyCheckededListBox();
            SuspendLayout();
            // 
            // _checkedListBox1
            // 
            _myCheckededListBox.Dock = DockStyle.Fill;
            _myCheckededListBox.FormattingEnabled = true;
            _myCheckededListBox.Location = new Point(20, 20);
            _myCheckededListBox.Name = "_checkedListBox1";
            _myCheckededListBox.Size = new Size(244, 222);
            _myCheckededListBox.TabIndex = 0;
            // 
            // CheckedListBoxForm
            // 
            ClientSize = new Size(284, 262);
            Controls.Add(_myCheckededListBox);
            Name = "CheckedListBoxForm";
            Padding = new Padding(20);
            FormClosed += new FormClosedEventHandler(CheckedListBoxForm_FormClosed);
            ResumeLayout(false);

        }

        public List<string> GetCheckedList()
        {

            return _checkedEnables;
        }

        private void CheckedListBoxForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            _checkedEnables = new List<string>();
            foreach (var item in _myCheckededListBox.CheckedItems)
            {
                _checkedEnables.Add(item.ToString());
            }
        }

    }

}