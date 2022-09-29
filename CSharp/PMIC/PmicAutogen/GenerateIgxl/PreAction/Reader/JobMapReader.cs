using CommonLib.Extension;
using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.PreAction.Reader
{
    public class JobTemperatureMap
    {
        public const string Rt = "RT";
        public const string Ht = "HT";
        public const string Cp = "CP";
        public const string Ft = "FT";

        public JobTemperatureMap(string jobName, string temperature, string testStage)
        {
            JobName = jobName;
            Temperature = temperature;
            TestStage = testStage;
        }

        public string JobName { set; get; }

        public string Temperature { set; get; }

        public string TestStage { set; get; }
    }


    public class JobMapReader
    {
        private readonly Dictionary<string, List<string>> _jobMapDictionary;
        private int _endColIndex;
        private int _endRowIndex;
        private ExcelWorksheet _excelWorksheet;
        private int _startColIndex;
        private int _startRowIndex;

        #region Constructor

        public JobMapReader()
        {
            _jobMapDictionary = new Dictionary<string, List<string>>();
            _startRowIndex = _startColIndex = _endColIndex = _endRowIndex = 1;
        }

        #endregion

        public List<JobTemperatureMap> JobTempMap { get; set; } = new List<JobTemperatureMap>();

        #region Member Function

        public Dictionary<string, List<string>> ReadFlow(ExcelWorksheet jobMappingSheet)
        {
            try
            {
                _excelWorksheet = jobMappingSheet;
                InitIndex();
                ReadData();
            }
            catch (Exception e)
            {
                throw new Exception("Error in reading job mapping sheet! " + e.Message);
            }

            return _jobMapDictionary;
        }

        private void InitIndex()
        {
            var startAddress = _excelWorksheet.Dimension.Start;
            _startColIndex = startAddress.Column;
            _startRowIndex = startAddress.Row;
            var endAddress = _excelWorksheet.Dimension.End;
            _endColIndex = endAddress.Column;
            _endRowIndex = endAddress.Row;
        }

        private void ReadData()
        {
            for (var j = _startColIndex; j <= _endColIndex; j++)
            {
                var i = _startRowIndex;
                var testSetting = _excelWorksheet.GetCellValue(i, j).Trim();
                if (!testSetting.Equals(""))
                {
                    //string stageName = this.GetStage(testSetting);
                    var jobList = new List<string>();
                    for (i++; i <= _endRowIndex; i++)
                    {
                        var jobName = _excelWorksheet.GetCellValue(i, j);
                        if (!jobName.Equals(""))
                        {
                            if (jobName.Split(':').Length > 1)
                            {
                                var temp = jobName.Split(':')[1];
                                var job = jobName.Split(':')[0];
                                jobList.Add(job);
                                JobTempMap.Add(new JobTemperatureMap(job, temp, testSetting)); //("CP1", "RT", "CP")
                            }
                            else
                            {
                                jobList.Add(jobName);
                            }
                        }
                    }

                    _jobMapDictionary.Add(testSetting, jobList);
                }
            }
        }

        #endregion
    }
}