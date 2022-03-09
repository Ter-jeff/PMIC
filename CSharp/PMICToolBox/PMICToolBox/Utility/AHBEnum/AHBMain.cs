using Library.Function;
using Library.Function.ErrorReport;
using OfficeOpenXml;
using PmicAutomation.Utility.AHBEnum.Input;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using PmicAutomation.MyControls;

namespace PmicAutomation.Utility.AHBEnum
{
    public class AhbMain
    {
        private readonly MyForm.RichTextBoxAppend _appendText;
        private readonly string _ahbRegister;
        private readonly string _outputPath;
        private readonly int _maxBitWidth;
        private readonly bool _regNameAndFieldName;
        private bool _regNameOnly;

        private AhbRegisterMapSheet _ahbRegSheet;

        public AhbMain(string ahbRegister, string outputPath, string maxBitWidth, bool regNameAndFieldName, bool regNameOnly)
        {
            _ahbRegister = ahbRegister;
            int defaultBitWidth;
            int.TryParse(maxBitWidth, out defaultBitWidth);
            _maxBitWidth = defaultBitWidth;
            _regNameAndFieldName = regNameAndFieldName;
            _regNameOnly = regNameOnly;
        }

        public AhbMain(AhbEnum ahbEnum)
        {
            _appendText = ahbEnum.AppendText;
            _ahbRegister = ahbEnum.FileOpen_AhbRegister.ButtonTextBox.Text;
            _outputPath = ahbEnum.FileOpen_OutputPath.ButtonTextBox.Text;
            int defaultBitWidth;
            int.TryParse(ahbEnum.textBox_MaxBitWidth.Text, out defaultBitWidth);
            _maxBitWidth = defaultBitWidth;
            _regNameAndFieldName = ahbEnum.radioButton_RegNameAndFieldName.Checked;
            _regNameOnly = ahbEnum.radioButton_FieldNameOnly.Checked;
            ErrorManager.ResetError();
        }

        public void WorkFlow()
        {
            ReadFiles();

            Check();

            GenFiles();

            _appendText.Invoke("All processes were completed !!!", Color.Black);
        }

        private void ReadFiles()
        {
            using (ExcelPackage inputExcel =
                new ExcelPackage(new FileInfo(_ahbRegister)))
            {
                foreach (ExcelWorksheet sheet in inputExcel.Workbook.Worksheets)
                {
                    _appendText.Invoke("Starting to read Ahb RegisterMap " + sheet.Name + " ...", Color.Black);
                    AhbRegisterMapReader ahbRegisterMapReader = new AhbRegisterMapReader();
                    _ahbRegSheet = ahbRegisterMapReader.ReadSheet(sheet);
                    if (_ahbRegSheet != null)
                    {
                        break;
                    }
                }
            }
        }

        private void Check()
        {
            int maxBitWidth = _ahbRegSheet.GetMaxBitWidth();
            if (maxBitWidth == _maxBitWidth)
            {
                _appendText.Invoke("The max of bit width is 8 ...", Color.Black);
            }
            else
            {
                _appendText.Invoke("The max of bit width " + maxBitWidth + " is not " + _maxBitWidth + " !!!", Color.Red);
            }

            AhbChecker AhbChecker = new AhbChecker();
            AhbChecker.CheckDuplicateAhbEnum(_ahbRegSheet);
        }

        private void GenFiles()
        {
            GenAhbEnum();

            GenErrorReport();
        }

        private void GenAhbEnum()
        {
            List<string> lines = _ahbRegSheet.GenAhbEnum(_regNameAndFieldName, _maxBitWidth);
            _appendText.Invoke("Starting to generate AHB_REG_MAP.bas ...", Color.Black);
            BasFileManage.GenBasFile(Path.Combine(_outputPath, "AHB_REG_MAP.bas"), lines);
        }

        private void GenErrorReport()
        {
            if (ErrorManager.GetErrorCount() <= 0)
            {
                return;
            }

            _appendText.Invoke("Starting to print error report ...", Color.Red);
            string outputFile = Path.Combine(_outputPath, "Error.xlsx");
            List<string> files = new List<string> { _ahbRegister };
            ErrorManager.GenErrorReport(outputFile, files);
        }
    }
}