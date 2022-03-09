using System;
using System.Text.RegularExpressions;

namespace PmicAutogen.Inputs.Setting.BinNumber
{
    public class SoftBinRangeRow
    {
        private bool _isExceed;
        private int _softBinCurrent;
        private int _softBinEnd;
        private int _softBinStart;

        public SoftBinRangeRow()
        {
            Description = "";
            Start = "";
            End = "";
            State = "";
            Bin = "";

            HardBin = "";
            HardHvBin = "";
            HardLvBin = "";
            HardNvBin = "";
            HardHlvBin = "";

            Block = "";
            Condition = "";

            _softBinStart = -1;
            _softBinEnd = -1;
            _softBinCurrent = -1;
        }

        public string Description { get; set; }
        public string Block { get; set; }
        public string Condition { get; set; }
        public string Start { get; set; }
        public string End { get; set; }
        public string State { get; set; }
        public string Bin { get; set; }
        public string HardBin { get; set; }
        public string HardHvBin { get; set; }
        public string HardLvBin { get; set; }
        public string HardNvBin { get; set; }
        public string HardHlvBin { get; set; }

        private void ConvertNumber()
        {
            var convert1 = int.TryParse(Start, out _softBinStart);
            var convert2 = int.TryParse(End, out _softBinEnd);

            if (!convert1 || !convert2)
                throw new Exception(string.Format("The Bin Number range: {0} is not a number, start:{1} ,end:{2}",
                    Description, Start, End));

            if (_softBinStart > _softBinEnd)
                throw new Exception(string.Format(
                    "Error for Bin Number range of:{0} ,start number:{1} can not be larger than end number:{2}!",
                    Description, Start, End));
            _softBinCurrent = _softBinStart;
        }

        public int GetSoftBinNumber()
        {
            if (_softBinStart == -1)
            {
                ConvertNumber();
                return _softBinCurrent;
            }

            if (_softBinCurrent < _softBinEnd)
            {
                _isExceed = false;
                return ++_softBinCurrent;
            }

            _isExceed = true;
            return _softBinEnd;
        }

        public int GetSoftBinStart()
        {
            if (_softBinStart == -1) ConvertNumber();
            return _softBinStart;
        }

        public int GetSoftBinEnd()
        {
            if (_softBinStart == -1) ConvertNumber();
            return _softBinEnd;
        }

        public bool CheckExceed()
        {
            return _isExceed;
        }

        public string GetStatus()
        {
            if (State.Equals("B", StringComparison.OrdinalIgnoreCase)) return "Fail";
            if (State.Equals("G", StringComparison.OrdinalIgnoreCase)) return "Pass";
            if (State.Equals("O", StringComparison.OrdinalIgnoreCase)) return "Fail-Stop";
            return "Fail";
        }

        public bool Match(string inputCondition)
        {
            var conditions = Regex.Replace(Condition, @"\s*", "").Split(',');
            inputCondition = Regex.Replace(inputCondition, @"\s+", "");
            foreach (var s in conditions)
                if (s.Equals(inputCondition, StringComparison.OrdinalIgnoreCase))
                    return true;
            return false;
        }
    }
}