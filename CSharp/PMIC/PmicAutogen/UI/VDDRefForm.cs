using PmicAutogen.Inputs.TestPlan.Reader;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace PmicAutogen
{
    public partial class VddRefForm : Form
    {
        private readonly Dictionary<string, string> _pinList = new Dictionary<string, string>();
        private readonly Dictionary<string, VddLevelsRow> _vddPinInfoList = new Dictionary<string, VddLevelsRow>();

        public VddRefForm(Dictionary<string, string> pinList, Dictionary<string, VddLevelsRow> vDdPinList)
        {
            InitializeComponent();
            _pinList = pinList;
            _vddPinInfoList = vDdPinList;

            var comboBoxColumn = (DataGridViewComboBoxColumn)dataGridView1.Columns[2];
            comboBoxColumn.Items.Add("No Reference");
            comboBoxColumn.Items.AddRange(_vddPinInfoList.Keys.ToArray());

            for (var i = 0; i < _pinList.Count; i++)
            {
                var pinInfo = _pinList.ElementAt(i);
                string[] row = { pinInfo.Key, pinInfo.Value };
                dataGridView1.Rows.Add(row);
                var dataGridViewComboBoxCell = (DataGridViewComboBoxCell)dataGridView1.Rows[i].Cells[2];
                dataGridViewComboBoxCell.Value = "No Reference";
            }

            dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;
        }

        public Dictionary<string, VddLevelsRow> RefVddPins { get; set; }

        public void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Columns[e.ColumnIndex].Name == "ReferencePin")
            {
                var cell = (DataGridViewComboBoxCell)dataGridView1.Rows[e.RowIndex].Cells[2];
                var pin = cell.Value.ToString();
                if (pin == "No Reference")
                    return;

                var pinVolt = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                var vddLevelRow = _vddPinInfoList[pin];

                if (!CheckMapping(pinVolt, vddLevelRow))
                {
                    var msg = string.Format("Voltage Not Match!\n  IO Pin {0} : {1}\n Vdd Pin {2} : {3}.",
                        dataGridView1.Rows[e.RowIndex].Cells[0].Value,
                        pinVolt,
                        pin, vddLevelRow.Nv);
                    MessageBox.Show(msg, "Warnning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private bool CheckMapping(string pinVolt, VddLevelsRow vddLevelRow)
        {
            if (vddLevelRow.Nv != vddLevelRow.Lv ||
                vddLevelRow.Nv != vddLevelRow.Hv ||
                (vddLevelRow.ULv != "" && vddLevelRow.Nv != vddLevelRow.ULv) ||
                (vddLevelRow.UHv != "" && vddLevelRow.Nv != vddLevelRow.UHv))
                return true;

            if (vddLevelRow.Nv != pinVolt) return false;
            return true;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            Click_Ok();
        }

        public void Click_Ok()
        {
            RefVddPins = new Dictionary<string, VddLevelsRow>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                var combox = row.Cells[2];
                if (combox.Value.ToString() != "No Reference")
                {
                    var ioPin = row.Cells[0].Value;
                    var vddLevelRow = _vddPinInfoList[combox.Value.ToString()];

                    if (vddLevelRow.Nv != vddLevelRow.Lv ||
                        vddLevelRow.Nv != vddLevelRow.Hv ||
                        (vddLevelRow.ULv != "" && vddLevelRow.Nv != vddLevelRow.ULv) ||
                        (vddLevelRow.UHv != "" && vddLevelRow.Nv != vddLevelRow.UHv))
                        RefVddPins.Add(ioPin.ToString(), vddLevelRow);
                }
            }

            DialogResult = DialogResult.OK;
        }
    }
}