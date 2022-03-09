using Library.Function;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text.RegularExpressions;

namespace PmicAutomation.Utility.PA.Input
{
    [Serializable]
    public class PaRow
    {
        private string _bumpName;
        public int RowNum;
        public string SourceSheetName;

        public PaRow()
        {
            No = "";
            Site = "";
            //BumpNumber = "";
            BumpName = "";
            //BumpX = "";
            //BumpY = "";
            //Ball = "";
            //BallName = "";
            //Channel = "";
            Assignment = "";
            Pogo = "";
            Ps = "";
            PaType = "";
            InstrumentType = "";
            ChannelMapName = "";
            GenPattern = "";
        }

        public string No { get; set; }
        public string Site { get; set; }

        public string BumpName
        {
            set { _bumpName = value; }
            get { return string.IsNullOrEmpty(ChannelMapName) ? _bumpName : ChannelMapName; }
        }

        public string Assignment { get; set; }
        public string Pogo { get; set; }
        public string Ps { get; set; }
        public string PaType { get; set; }
        public string PinMapType { get; set; }
        public string InstrumentType { get; set; }
        public string ChannelMapName { get; set; }
        public string GenPattern { get; set; }
        public string PinType { get; set; }

        public bool IsPower()
        {
            return Regex.IsMatch(Ps, "power", RegexOptions.IgnoreCase);
        }

        public PaRow DeepClone()
        {
            using (MemoryStream stream = new MemoryStream())
            {
                BinaryFormatter formatter = new BinaryFormatter();
                formatter.Serialize(stream, this);
                stream.Seek(0, SeekOrigin.Begin);
                return (PaRow)formatter.Deserialize(stream);
            }
        }
    }

    public class PaSheet
    {
        #region Constructor

        public PaSheet()
        {
            Rows = new List<PaRow>();
        }

        #endregion

        #region Properity

        public List<PaRow> Rows { get; set; }

        #endregion

        #region Field

        public string Name;
        public int NoIndex;
        public int SiteIndex;
        public int BumpNumberIndex;
        public int BumpNameIndex;
        public int BumpXIndex;
        public int BumpYIndex;
        public int BallIndex;
        public int BallNameIndex;
        public int ChannelIndex;
        public int AssignmentIndex;
        public int PogoIndex;
        public int PsIndex;
        public int TypeIndex;
        public int ChannelMapNameIndex;
        public int GenPatternIndex;
        public int PinTypeIndex;
        #endregion
    }

    public class PaCsvReader
    {
        private const string NoHeader = @"No";
        private const string SiteHeader = @"Site";
        private const string BumpNumberHeader = @"Bump\s*Number";
        private const string BumpNameHeader = @"Bump\s*Name";
        private const string BumpXHeader = @"BumpX";
        private const string BumpYHeader = @"BumpY";
        private const string BallHeader = @"Ball";
        private const string BallNameHeader = @"Ball\s*Name";
        private const string ChannelHeader = @"Channel";
        private const string AssignmentHeader = @"Assignment";
        private const string PogoHeader = @"Pogo";
        private const string PsHeader = @"Ps";
        private const string TypeHeader = @"Type";
        private const string ChannelMapNameHeader = @"ChannelMap Name";
        private const string GenPatternHeader = @"GenPattern";
        private const string PinTypeHeader = @"PinType";

        private readonly UflexConfig _uflexConfig;
        private string _file;

        public PaCsvReader(UflexConfig uflexConfig)
        {
            _uflexConfig = uflexConfig;
        }

        public PaSheet Read(string file)
        {
            try
            {
                PaSheet paSheet = new PaSheet();

                _file = file;
                paSheet.Name = Path.GetFileName(_file);

                GetHeader(file, paSheet);

                GetRows(file, paSheet);

                return paSheet;
            }
            catch (Exception e)
            {
                throw new Exception("Reading PA failed, may be caused by wrong format. " + e.Message);
            }
        }

        private void GetRows(string file, PaSheet paSheet)
        {
            int lineCount = 2;
            List<PaRow> paRows = new List<PaRow>();
            using (StreamReader fileReader =
                new StreamReader(new FileStream(file, FileMode.Open, FileAccess.ReadWrite)))
            {
                string line;
                while ((line = fileReader.ReadLine()) != null)
                {
                    if (!Regex.IsMatch(line, @"^[\*|\d].*"))
                    {
                        continue;
                    }

                    line = line.Replace(" ", "");
                    Regex csvParser = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");
                    string[] strArray = csvParser.Split(line);
                    PaRow row = new PaRow { SourceSheetName = file, RowNum = lineCount++ };
                    if (paSheet.NoIndex != -1)
                    {
                        row.No = strArray[paSheet.NoIndex];
                    }

                    if (paSheet.SiteIndex != -1)
                    {
                        row.Site = strArray[paSheet.SiteIndex];
                    }

                    //if (paSheet.BumpNumberIndex != -1)
                    //    newItem.BumpNumber = strArray[paSheet.BumpNumberIndex];
                    if (paSheet.BumpNameIndex != -1)
                    {
                        row.BumpName = strArray[paSheet.BumpNameIndex].Replace("[", "").Replace("]", "").ToUpper();
                    }

                    //if (paSheet.BumpXIndex != -1)
                    //    newItem.BumpX = strArray[paSheet.BumpXIndex];
                    //if (paSheet.BumpYIndex != -1)
                    //    newItem.BumpY = strArray[paSheet.BumpYIndex];
                    //if (paSheet.BallIndex != -1)
                    //    newItem.Ball = strArray[paSheet.BallIndex];
                    //if (paSheet.BallNameIndex != -1)
                    //    newItem.BallName = strArray[paSheet.BallNameIndex];
                    //if (paSheet.ChannelIndex != -1)
                    //    newItem.Channel = strArray[paSheet.ChannelIndex];
                    if (paSheet.AssignmentIndex != -1)
                    {
                        row.Assignment = strArray[paSheet.AssignmentIndex];
                    }

                    if (paSheet.PogoIndex != -1)
                    {
                        row.Pogo = strArray[paSheet.PogoIndex];
                    }

                    if (paSheet.PsIndex != -1)
                    {
                        row.Ps = strArray[paSheet.PsIndex];
                    }

                    if (paSheet.TypeIndex != -1)
                    {
                        row.PaType = strArray[paSheet.TypeIndex] == "IO" ? "I/O" : strArray[paSheet.TypeIndex];
                    }

                    if (paSheet.ChannelMapNameIndex != -1)
                    {
                        row.ChannelMapName = strArray[paSheet.ChannelMapNameIndex].Replace("[", "").Replace("]", "")
                            .ToUpper();
                    }

                    if (paSheet.GenPatternIndex != -1)
                    {
                        row.GenPattern = strArray[paSheet.GenPatternIndex].ToUpper();
                    }

                    if (paSheet.PinTypeIndex != -1)
                    {
                        row.PinType = strArray[paSheet.PinTypeIndex].ToUpper();
                    }

                    row.InstrumentType = _uflexConfig.GetToolTypeByChannelAssignment(row.Assignment);

                    if (row.BumpName != "" && row.Assignment != "" && row.Assignment.ToUpper() != "#N/A")
                    {
                        if (row.BumpName.Contains("/"))
                        {
                            row.BumpName = row.BumpName.Split('/')[0];
                        }

                        paRows.Add(row);
                    }
                }
            }

            paSheet.Rows.AddRange(paRows);
        }

        private void GetHeader(string file, PaSheet paSheet)
        {
            string header = "";
            using (StreamReader fileReader =
                new StreamReader(new FileStream(file, FileMode.Open, FileAccess.ReadWrite)))
            {
                string line;
                while ((line = fileReader.ReadLine()) != null)
                {
                    if (Regex.IsMatch(line, @"^\d"))
                    {
                        break;
                    }

                    header += line;
                }

                Regex csvParser = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");
                List<string> headerList = csvParser.Split(header).ToList();
                paSheet.NoIndex = headerList.FindIndex(a => Regex.IsMatch(a, NoHeader, RegexOptions.IgnoreCase));
                paSheet.SiteIndex = headerList.FindIndex(a => Regex.IsMatch(a, SiteHeader, RegexOptions.IgnoreCase));
                paSheet.BumpNumberIndex =
                    headerList.FindIndex(a => Regex.IsMatch(a, BumpNumberHeader, RegexOptions.IgnoreCase));
                paSheet.BumpNameIndex =
                    headerList.FindIndex(a => Regex.IsMatch(a, BumpNameHeader, RegexOptions.IgnoreCase));
                paSheet.BumpXIndex = headerList.FindIndex(a => Regex.IsMatch(a, BumpXHeader, RegexOptions.IgnoreCase));
                paSheet.BumpYIndex = headerList.FindIndex(a => Regex.IsMatch(a, BumpYHeader, RegexOptions.IgnoreCase));
                paSheet.BallIndex = headerList.FindIndex(a => Regex.IsMatch(a, BallHeader, RegexOptions.IgnoreCase));
                paSheet.BallNameIndex =
                    headerList.FindIndex(a => Regex.IsMatch(a, BallNameHeader, RegexOptions.IgnoreCase));
                paSheet.ChannelIndex =
                    headerList.FindIndex(a => a.Equals(ChannelHeader, StringComparison.OrdinalIgnoreCase));
                paSheet.AssignmentIndex =
                    headerList.FindIndex(a => Regex.IsMatch(a, AssignmentHeader, RegexOptions.IgnoreCase));
                paSheet.PogoIndex = headerList.FindIndex(a => Regex.IsMatch(a, PogoHeader, RegexOptions.IgnoreCase));
                paSheet.PsIndex = headerList.FindIndex(a => Regex.IsMatch(a, PsHeader, RegexOptions.IgnoreCase));
                paSheet.TypeIndex = headerList.FindIndex(a => Regex.IsMatch(a, TypeHeader, RegexOptions.IgnoreCase));
                paSheet.ChannelMapNameIndex = headerList.FindIndex(a =>
                    a.Equals(ChannelMapNameHeader, StringComparison.OrdinalIgnoreCase));
                paSheet.GenPatternIndex = headerList.FindIndex(a =>
                    a.Equals(GenPatternHeader, StringComparison.OrdinalIgnoreCase));
                paSheet.PinTypeIndex = headerList.FindIndex(a =>
                    a.Equals(PinTypeHeader, StringComparison.OrdinalIgnoreCase));
            }
        }
    }

    public class PaExcelReader
    {
        private const string HeaderNo = "No";
        private const string HeaderSite = "Site";
        private const string HeaderBumpName = "Bump Name";
        private const string HeaderAssignment = "Assignment";
        private const string HeaderPogo = "POGO";
        private const string HeaderPs = "PS";
        private const string HeaderType = "Type";
        private const string HeaderChannelMapName = "ChannelMap Name";
        private const string GenPatternHeader = @"GenPattern";
        private const string PinTypeHeader = @"PinType";
        private readonly UflexConfig _uflexConfig;
        private int _assignmentIndex = -1;
        private int _bumpNameIndex = -1;
        private int _channelMapNameIndex = -1;
        private int _genPatternIndex = -1;
        private int _pinTypeIndex = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;
        private int _noIndex = -1;
        private PaSheet _pAExcelSheet;
        private int _pogoIndex = -1;
        private int _psIndex = -1;
        private int _siteIndex = -1;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _typeIndex = -1;

        public PaExcelReader(UflexConfig uflexConfig)
        {
            _uflexConfig = uflexConfig;
        }

        public PaSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null)
            {
                return null;
            }

            _pAExcelSheet = new PaSheet();

            _excelWorksheet = worksheet;

            _pAExcelSheet.Name = worksheet.Name;

            Reset();

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            if (!GetHeaderIndex())
            {
                return null;
            }

            _pAExcelSheet = ReadSheetData();

            return _pAExcelSheet;
        }

        private PaSheet ReadSheetData()
        {
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                PaRow row = new PaRow { RowNum = i };
                if (_noIndex != -1)
                {
                    row.No = _excelWorksheet.GetMergeCellValue(i, _noIndex).Trim();
                }

                if (_siteIndex != -1)
                {
                    row.Site = _excelWorksheet.GetMergeCellValue(i, _siteIndex).Trim();
                }

                if (_bumpNameIndex != -1)
                {
                    row.BumpName = _excelWorksheet.GetMergeCellValue(i, _bumpNameIndex).Trim().Replace("[", "")
                        .Replace("]", "").ToUpper();
                }

                if (_assignmentIndex != -1)
                {
                    row.Assignment = _excelWorksheet.GetMergeCellValue(i, _assignmentIndex).Trim();
                }

                if (_pogoIndex != -1)
                {
                    row.Pogo = _excelWorksheet.GetMergeCellValue(i, _pogoIndex).Trim();
                }

                if (_psIndex != -1)
                {
                    row.Ps = _excelWorksheet.GetMergeCellValue(i, _psIndex).Trim();
                }

                if (_typeIndex != -1)
                {
                    row.PaType = _excelWorksheet.GetMergeCellValue(i, _typeIndex).Trim();
                }

                if (_channelMapNameIndex != -1)
                {
                    row.ChannelMapName = _excelWorksheet.GetMergeCellValue(i, _channelMapNameIndex).Trim()
                        .Replace("[", "").Replace("]", "").ToUpper();
                }

                if (_genPatternIndex != -1)
                {
                    row.ChannelMapName = _excelWorksheet.GetMergeCellValue(i, _channelMapNameIndex).Trim().ToUpper();
                }

                if (_pinTypeIndex != -1)
                {
                    row.PinType = _excelWorksheet.GetMergeCellValue(i, _pinTypeIndex).Trim().ToUpper();
                }

                row.InstrumentType = _uflexConfig.GetToolTypeByChannelAssignment(row.Assignment);

                if (row.BumpName != "" && row.Assignment != "" && row.Assignment.ToUpper() != "#N/A")
                {
                    if (row.BumpName.Contains("/"))
                    {
                        row.BumpName = row.BumpName.Split('/')[0];
                    }

                    _pAExcelSheet.Rows.Add(row);
                }
            }

            _pAExcelSheet.NoIndex = _noIndex;
            _pAExcelSheet.SiteIndex = _siteIndex;
            _pAExcelSheet.BumpNameIndex = _bumpNameIndex;
            _pAExcelSheet.AssignmentIndex = _assignmentIndex;
            _pAExcelSheet.PogoIndex = _pogoIndex;
            _pAExcelSheet.PsIndex = _psIndex;
            _pAExcelSheet.TypeIndex = _typeIndex;
            _pAExcelSheet.ChannelMapNameIndex = _channelMapNameIndex;
            _pAExcelSheet.GenPatternIndex = _genPatternIndex;
            _pAExcelSheet.PinTypeIndex = _pinTypeIndex;
            return _pAExcelSheet;
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = _excelWorksheet.GetMergeCellValue(_startRowNumber, i).Trim();
                if (lStrHeader.Equals(HeaderNo, StringComparison.OrdinalIgnoreCase))
                {
                    _noIndex = i;
                    continue;
                }

                if (Regex.IsMatch(lStrHeader, HeaderSite, RegexOptions.IgnoreCase))
                {
                    _siteIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderBumpName, StringComparison.OrdinalIgnoreCase))
                {
                    _bumpNameIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderAssignment, StringComparison.OrdinalIgnoreCase))
                {
                    _assignmentIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderPogo, StringComparison.OrdinalIgnoreCase))
                {
                    _pogoIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderPs, StringComparison.OrdinalIgnoreCase))
                {
                    _psIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderType, StringComparison.OrdinalIgnoreCase))
                {
                    _typeIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderChannelMapName, StringComparison.OrdinalIgnoreCase))
                {
                    _channelMapNameIndex = i;
                }

                if (lStrHeader.Equals(GenPatternHeader, StringComparison.OrdinalIgnoreCase))
                {
                    _genPatternIndex = i;
                }

                if (lStrHeader.Equals(PinTypeHeader, StringComparison.OrdinalIgnoreCase))
                {
                    _pinTypeIndex = i;
                }
            }

            return true;
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            int colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (int i = 1; i <= rowNum; i++)
                for (int j = 1; j <= colNum; j++)
                {
                    if (_excelWorksheet.GetMergeCellValue(i, j).Trim().Equals(HeaderNo, StringComparison.OrdinalIgnoreCase))
                    {
                        _startRowNumber = i;
                        return true;
                    }
                }

            return false;
        }

        private bool GetDimensions()
        {
            if (_excelWorksheet.Dimension != null)
            {
                _startColNumber = _excelWorksheet.Dimension.Start.Column;
                _startRowNumber = _excelWorksheet.Dimension.Start.Row;
                _endColNumber = _excelWorksheet.Dimension.End.Column;
                _endRowNumber = _excelWorksheet.Dimension.End.Row;
                return true;
            }

            return false;
        }

        private void Reset()
        {
            _startColNumber = -1;
            _startRowNumber = -1;
            _endColNumber = -1;
            _endRowNumber = -1;
            _noIndex = -1;
            _siteIndex = -1;
            _bumpNameIndex = -1;
            _assignmentIndex = -1;
            _pogoIndex = -1;
            _psIndex = -1;
            _typeIndex = -1;
            _channelMapNameIndex = -1;
            _genPatternIndex = -1;
            _pinTypeIndex = -1;
        }
    }
}