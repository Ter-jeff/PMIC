using System;
using System.IO;
using System.Linq;
using VBTPOPGenPreCheckBusiness.DataStore;
using OfficeOpenXml;
using System.Windows.Forms;

namespace VBTPOPGenPreCheckBusiness.Business
{
    public class TestplanReader
    {
        private string _testplan;
        private string _filename;
        private bool _isValidTestplan;
        private ExcelPackage _excelPackage;
        private ExcelWorkbook _workbook;       
        public static string[] ExcludingSheets = new string[] { "Status", "COMMANDS OVERVIEW", "Command_Category", "CALL_FUNCTION_LIST", "COMMENT_LOOKUP", "Sheet1", "VbtGen_template" };

        public TestplanReader(string tp)
        {
            _testplan = tp;
            _filename = Path.GetFileNameWithoutExtension(tp);
            _isValidTestplan = CheckValidTestplan();
        }

        public bool IsValidTestplan
        {
            get { return _isValidTestplan; }
        }

        private bool CheckValidTestplan()
        {
            bool isValid = false;
            _excelPackage = new ExcelPackage(new FileInfo(_testplan));
            _workbook = _excelPackage.Workbook;
            if (_workbook.Worksheets.ToList().Exists(s => s.Name.IndexOf("_TestParameter", StringComparison.OrdinalIgnoreCase) != -1))
                isValid = true;

            return isValid;
        }

        public TestPlanFile Read()
        {
            TestPlanFile testPlan = new TestPlanFile(_filename);
            int headerRowIndex = -1;

            int parameterFunctionNameIndex = -1;
            int blockNameIndex = -1;
            int otpRegisterIndex = -1;
            int trimRegisterIndex = -1;
            int trimBitFieldIndex = -1;
            int measPinIndex = -1;
            int powerPinIndex = -1;
            int analogPinIndex = -1;
            int numbitsIndex = -1;

            int topListIndex = -1;
            int commandIndex = -1;
            int functionNameIndex = -1;
            int registerIndex = -1;
            int bitfieldIndex = -1;
            int valuesIndex = -1;
            int pinIndex = -1;

            string parameterFunctionName = string.Empty;
            string blockName = string.Empty;
            string otpRegister = string.Empty;
            string trimRegister = string.Empty;
            string trimBitField = string.Empty;
            string numbits = string.Empty;
            string measPin = string.Empty;
            string powerPin = string.Empty;
            string analogPin = string.Empty;

            string topList = string.Empty;
            string command = string.Empty;
            string functionName = string.Empty;
            string register = string.Empty;
            string bitfieldName = string.Empty;
            string values = string.Empty;
            string pin = string.Empty;

            foreach (ExcelWorksheet sheet in _workbook.Worksheets)
            {
                if (ExcludingSheets.Contains(sheet.Name))
                    continue;
                //TestParameter Sheet
                if (sheet.Name.IndexOf("_TestParameter", StringComparison.OrdinalIgnoreCase) != -1)
                {
                    TestParameterSheet parameterSheet = new TestParameterSheet();
                    parameterSheet.SheetName = sheet.Name;
                    testPlan.ParameterSheet = parameterSheet;
                    FetchParameterHeaderInfo(sheet, ref headerRowIndex, ref parameterFunctionNameIndex, ref blockNameIndex, ref otpRegisterIndex, ref trimRegisterIndex, ref trimBitFieldIndex, ref numbitsIndex, ref measPinIndex, ref powerPinIndex, ref analogPinIndex);
                    for (int r = headerRowIndex + 1; r <= sheet.Dimension.Rows; ++r)
                    {
                        parameterFunctionName = GetCellValue(sheet, r, parameterFunctionNameIndex);
                        blockName = GetCellValue(sheet, r, blockNameIndex);
                        otpRegister = GetCellValue(sheet, r, otpRegisterIndex);
                        trimRegister = GetCellValue(sheet, r, trimRegisterIndex);
                        trimBitField = GetCellValue(sheet, r, trimBitFieldIndex);
                        numbits = GetCellValue(sheet, r, numbitsIndex);
                        measPin = GetCellValue(sheet, r, measPinIndex);
                        powerPin = GetCellValue(sheet, r, powerPinIndex);
                        analogPin = GetCellValue(sheet, r, analogPinIndex);

                        //if (otpRegister.StartsWith("OTP_", StringComparison.OrdinalIgnoreCase) && !trimRegister.ToUpper().Equals("N/A") && !trimBitField.ToUpper().Equals("N/A") && !numbits.ToUpper().Equals("N/A"))
                        //if (otpRegister.StartsWith("OTP_", StringComparison.OrdinalIgnoreCase))
                        //{
                            var tokens = otpRegister.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                            bool multiBlock = tokens.Count() > 1 ? true : false;
                            for (int i = 0; i < tokens.Count(); ++i)
                            {
                                parameterSheet.TestParameterRows.Add(new TestParameterRow(r, parameterFunctionName, blockName, numbits, trimRegister, trimBitField, tokens[i], measPin, powerPin, analogPin, multiBlock));
                            }
                        //}
                    }
                }
                else
                {
                    TestFlowSheet testFlowSheet = new TestFlowSheet(sheet.Name);
                    testPlan.TestFlowSheetlst.Add(testFlowSheet);
                    if (!FetchTestFlowSheetrHeaderInfo(sheet, ref headerRowIndex, ref topListIndex, ref commandIndex, ref functionNameIndex, ref registerIndex, ref bitfieldIndex, ref valuesIndex, ref pinIndex))
                        continue;
                    for (int r = headerRowIndex + 1; r <= sheet.Dimension.Rows; ++r)
                    {
                        command = GetCellValue(sheet, r, commandIndex);
                        if (!string.IsNullOrEmpty(command))
                        {
                            topList = GetCellValue(sheet, r, topListIndex);
                            functionName = GetCellValue(sheet, r, functionNameIndex);
                            register = GetCellValue(sheet, r, registerIndex);
                            bitfieldName = GetCellValue(sheet, r, bitfieldIndex);
                            values = GetCellValue(sheet, r, valuesIndex);
                            pin = GetCellValue(sheet, r, pinIndex);
                            testFlowSheet.CommandRows.Add(new CommandRow(r, topList, command, functionName, register, bitfieldName, values, pin));
                        }

                    }
                }
            }

            return testPlan;
        }

        private void FetchParameterHeaderInfo(ExcelWorksheet sheet, ref int headerRowIndex, ref int parameterFunctionNameIndex, ref int blockNameIndex, ref int otpRegisterIndex, ref int trimRegisterIndex, ref int trimBitFieldIndex, ref int numbitsIndex, ref int measPinIndex, ref int powerPinIndex, ref int analogPinIndex)
        {
            const string FunctionName = "FunctionName";
            const string BlockName = "BlockName";
            const string OTPRegister = "OTPRegister";
            const string TrimRegister = "TrimRegister";
            const string TrimBitField = "TrimBitField";
            const string Numbits = "Numbits";
            const string MeasPin = "MeasPin";
            const string PowerPin = "PowerPin";
            const string AnalogPin = "AnalogSweepPin";
            string context;
            for (int r = 1; r <= sheet.Dimension.Rows; ++r)
            {
                if (sheet.Cells[r, 1].Value != null)
                {
                    string token = sheet.Cells[r, 1].Value.ToString();
                    if (!string.IsNullOrEmpty(token))
                    {
                        headerRowIndex = r;
                        for (int c = 1; c <= sheet.Dimension.Columns; ++c)
                        {
                            if (sheet.Cells[r, c].Value == null)
                                break;
                            context = sheet.Cells[r, c].Value.ToString().Trim();
                            if (context.Equals(FunctionName, StringComparison.OrdinalIgnoreCase))
                                parameterFunctionNameIndex = c;
                            else if (context.Equals(BlockName, StringComparison.OrdinalIgnoreCase))
                                blockNameIndex = c;
                            else if (context.Equals(TrimRegister, StringComparison.OrdinalIgnoreCase))
                                trimRegisterIndex = c;
                            else if (context.Equals(OTPRegister, StringComparison.OrdinalIgnoreCase))
                                otpRegisterIndex = c;
                            else if (context.Equals(TrimBitField, StringComparison.OrdinalIgnoreCase))
                                trimBitFieldIndex = c;
                            else if (context.Equals(Numbits, StringComparison.OrdinalIgnoreCase))
                                numbitsIndex = c;
                            else if (context.Equals(MeasPin, StringComparison.OrdinalIgnoreCase))
                                measPinIndex = c;
                            else if (context.Equals(PowerPin, StringComparison.OrdinalIgnoreCase))
                                powerPinIndex = c;
                            else if (context.Equals(AnalogPin, StringComparison.OrdinalIgnoreCase))
                                analogPinIndex = c;
                        }
                        if (headerRowIndex != -1 && otpRegisterIndex != -1 && trimRegisterIndex != -1 && trimBitFieldIndex != -1 && numbitsIndex != -1 &&
                            measPinIndex != -1 && powerPinIndex != -1 && analogPinIndex != -1)
                            break;
                    }
                }
            }

            if (headerRowIndex == -1 || otpRegisterIndex == -1 || trimRegisterIndex == -1 || trimBitFieldIndex == -1 || numbitsIndex == -1 ||
                            measPinIndex == -1 || powerPinIndex == -1 || analogPinIndex == -1)
                throw new Exception("Can not find needed header in parameter sheet: " + sheet.Name);
        }

        private bool FetchTestFlowSheetrHeaderInfo(ExcelWorksheet sheet, ref int headerRowIndex, ref int topListIndex, ref int commandIndex, ref int functionNameIndex, ref int registerIndex, ref int bitfieldIndex, ref int valuesIndex, ref int pinIndex)
        {
            const string Command = "COMMAND";
            const string TopList = "TOP_LIST";
            const string FunctionName = "FUNCTION_NAME";
            const string Register = "REGISTER/MACRO NAME";
            const string BitfieldName = "BITFIELD NAME";
            const string Values = "VALUE(S)";
            const string Pin = "PIN";
            string context;
            for (int r = 1; r <= sheet.Dimension.Rows; ++r)
            {
                headerRowIndex = r;
                for (int c = 1; c <= sheet.Dimension.Columns; ++c)
                {
                    context = GetCellValue(sheet, r, c);
                    if (string.IsNullOrEmpty(context))
                        continue;
                    if (context.Equals(TopList, StringComparison.OrdinalIgnoreCase))
                        topListIndex = c;
                    else if (context.Equals(Command, StringComparison.OrdinalIgnoreCase))
                        commandIndex = c;
                    else if (context.Equals(FunctionName, StringComparison.OrdinalIgnoreCase))
                        functionNameIndex = c;
                    else if (context.Equals(Register, StringComparison.OrdinalIgnoreCase))
                        registerIndex = c;
                    else if (context.Equals(BitfieldName, StringComparison.OrdinalIgnoreCase))
                        bitfieldIndex = c;
                    else if (context.Equals(Values, StringComparison.OrdinalIgnoreCase))
                        valuesIndex = c;
                    else if (context.Equals(Pin, StringComparison.OrdinalIgnoreCase))
                        pinIndex = c;
                }
                if (headerRowIndex != -1 && topListIndex != -1 && commandIndex != -1 && functionNameIndex != -1 && registerIndex != -1 && bitfieldIndex != -1 && valuesIndex != -1 && pinIndex != -1)
                    break;
            }

            if (headerRowIndex == -1 || topListIndex == -1 || commandIndex == -1 || functionNameIndex == -1 || registerIndex == -1 || bitfieldIndex == -1 || valuesIndex == -1 || pinIndex == -1)
            {
                MessageBox.Show("Can not find needed header in test flow sheet: " + sheet.Name + ", will skip to check this sheet.", "Test Flow Sheet Format Error", MessageBoxButtons.OK);
                return false;
            }
            return true;
        }

        private string GetCellValue(ExcelWorksheet sheet, int rowIndex, int columnIndex)
        {
            return sheet.Cells[rowIndex, columnIndex].Value == null ? string.Empty : sheet.Cells[rowIndex, columnIndex].Value.ToString().Trim();
        }
    }
}
