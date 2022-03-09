using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBTPOPGenPreCheckBusiness.DataStore
{
    public class TestPlanFile
    {
        public string FileName;
        public List<TestFlowSheet> TestFlowSheetlst;
        public TestParameterSheet ParameterSheet;

        public TestPlanFile(string fileName)
        {
            FileName = fileName;
            TestFlowSheetlst = new List<TestFlowSheet>();
        }

        public List<TestFlowSheet> FindValidTestFlowSheets()
        {
            List<string> functionNamelstInParameter = new List<string>();
            if (ParameterSheet != null)
                functionNamelstInParameter = ParameterSheet.TestParameterRows.Select(s => s.FunctionName).ToList().Distinct().ToList();
            return TestFlowSheetlst.FindAll(s => functionNamelstInParameter.Exists(f => f.Equals(s.GetMainFunctionName(), StringComparison.OrdinalIgnoreCase)));
        }
        
    }

    public class TestFlowSheet
    {
        public string SheetName;
        public List<CommandRow> CommandRows;

        public static List<string> commandNamelstToCheckRegister = new List<string>() { "OTP_WRITE", "AHB_READ", "AHB_WRITE", "VM_WRITE" };
        public static List<string> commandNamelstToCheckPin = new List<string>() { "SETUP_MV", "SETUP_FI_MV", "SETUP_FV_MI", "SETUP_FI_MI","SETUP_DIFF_DC30","SETUP_DIFF_UVI80",
            "SETUP_MEAS_FREQ","SETUP_DIG_PIN_INITIALSTATE","CHANGE_FORCE_V","CHANGE_FORCE_I","CHANGE_RANGE_V","CHANGE_RANGE_I","CHANGE_CURRENTANDRANGE","CHANGE_EDGETIME",
            "STROBE_V_I","STROBE_V_DIFF","STROBE_FREQ","MEAS_FAILCOUNT","RESET_SETUPFREQ","CONNECT_RELAY","DISCONNECT_RELAY","CONNECT_DIG","DISCONNECT_DIG","DISCONNECT","READ_I","READ_V"
        };

        public TestFlowSheet(string sheetName)
        {
            SheetName = sheetName;
            CommandRows = new List<CommandRow>();
        }

        public string GetMainFunctionName()
        {
            CommandRow startOfTestCommand = CommandRows.Find(s => s.CommandName.Equals("START_OF_TEST", StringComparison.OrdinalIgnoreCase) ||
                                                            s.CommandName.Equals("START_OF_TEST_DTB", StringComparison.OrdinalIgnoreCase));
            if (startOfTestCommand == null)
                return "";
            return startOfTestCommand.FunctionName;
        }

        public List<CommandRow> GetCommandlistToCheckRegister()
        {
            return CommandRows.FindAll(s => commandNamelstToCheckRegister.Exists(c => c.Equals(s.CommandName, StringComparison.OrdinalIgnoreCase)));
        }

        public List<CommandRow> GetCommandlistToCheckPin()
        {
            return CommandRows.FindAll(s => commandNamelstToCheckPin.Exists(c => c.Equals(s.CommandName, StringComparison.OrdinalIgnoreCase)));
        }
    }

    public class CommandRow
    {
        public int RowIndex;
        public string TopList;
        public string CommandName;
        public string FunctionName;
        public string RegisterName;
        public string BitFieldName;
        public string Values;
        public string Pin;

        public CommandRow(int rowIndex, string topList, string commandName, string functionName, string registerName, string bitfieldName, string values, string pin)
        {
            RowIndex = rowIndex;
            TopList = topList;
            CommandName = commandName;
            FunctionName = functionName;
            RegisterName = registerName;
            BitFieldName = bitfieldName;
            Values = values;
            Pin = pin;
        }
    }

    public class TestParameterSheet
    {
        public List<TestParameterRow> TestParameterRows;
        public string SheetName;
        public TestParameterSheet()
        {
            TestParameterRows = new List<TestParameterRow>();
        }
    }
    public class TestParameterRow
    {
        private int _row;
        private string _funtionName;
        private string _blockName;
        private string _numbits;
        private string _trimRegister;
        private string _trimBitField;
        private string _otpRegister;
        private string _measPin;
        private string _powerPin;
        private string _analogPin;

        private bool _multiBlock;

        public TestParameterRow(int row, string functionName, string blockName, string numbits, string trimRegister, string trimBitField, string otpRegister, string measPin, string powerPin, string analogPin, bool multiBlock)
        {
            _row = row;
            _funtionName = functionName;
            _blockName = blockName;
            _numbits = numbits;
            _trimRegister = trimRegister;
            _trimBitField = trimBitField;
            _otpRegister = otpRegister;
            _measPin = measPin;
            _powerPin = powerPin;
            _analogPin = analogPin;
            _multiBlock = multiBlock;
        }


        public int Row
        {
            get { return _row; }
        }

        public string FunctionName
        {
            get { return _funtionName; }
        }

        public string BlockName
        {
            get { return _blockName; }
        }
        public string Numbits
        {
            get { return _numbits; }
        }

        public string TrimRegister
        {
            get { return _trimRegister; }
        }

        public string TrimBitField
        {
            get { return _trimBitField; }
        }

        public string MeasPin
        {
            get { return _measPin; }
        }

        public string PowerPin
        {
            get { return _powerPin; }
        }

        public string AnalogPin
        {
            get { return _analogPin; }
        }

        public string OtpRegister
        {
            get { return _otpRegister; }
        }

        public bool MultiBlock
        {
            get { return _multiBlock; }
        }
    }
}
