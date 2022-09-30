using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text.RegularExpressions;

namespace IgxlData.IgxlBase
{
    [DebuggerDisplay("{TestName}")]
    [Serializable]
    public class InstanceRow : IgxlRow
    {
        public InstanceRow()
        {
            SheetName = "";
            RowNum = 0;
            TestName = "";
            Type = "";
            Name = "";
            CalledAs = "";
            DcCategory = "";
            DcSelector = "";
            AcCategory = "";
            AcSelector = "";
            TimeSets = "";
            EdgeSets = "";
            PinLevels = "";
            MixedSignalTiming = "";
            Overlay = "";
            ArgList = "";
            Args = new List<string>();
            InitList = new List<string>();
            PayloadList = new List<string>();
            Comment = "";
        }

        public string GetPropertyValue(string propertyName)
        {
            var myType = typeof(InstanceRow);
            var myPropInfo = myType.GetProperty(propertyName);
            return myPropInfo.GetValue(this, null).ToString();
        }

        public InstanceRow DeepClone()
        {
            using (Stream objectStream = new MemoryStream())
            {
                IFormatter formatter = new BinaryFormatter();
                formatter.Serialize(objectStream, this);
                objectStream.Seek(0, SeekOrigin.Begin);
                return formatter.Deserialize(objectStream) as InstanceRow;
            }
        }

        public InstanceRow SetInstanceRow(string[] arr)
        {
            var instanceRow = new InstanceRow();
            if (arr.Length < 130) return instanceRow;
            instanceRow.TestName = arr[1];
            instanceRow.Type = arr[2];
            instanceRow.Name = arr[3];
            instanceRow.CalledAs = arr[4];
            instanceRow.DcCategory = arr[5];
            instanceRow.DcSelector = arr[6];
            instanceRow.AcCategory = arr[7];
            instanceRow.AcSelector = arr[8];
            instanceRow.TimeSets = arr[9];
            instanceRow.EdgeSets = arr[10];
            instanceRow.PinLevels = arr[11];
            instanceRow.MixedSignalTiming = arr[12];
            instanceRow.Overlay = arr[13];
            instanceRow.ArgList = arr[14];
            for (var i = 15; i < 145; i++)
                instanceRow.Args.Add(arr[i]);
            instanceRow.Comment = arr[145];
            return instanceRow;
        }

        public void SetArgument(string argument, string value)
        {
            var index = ArgList.Split(',').ToList().IndexOf(argument);
            if (index != -1)
            {
                if (index < Args.Count)
                {
                    Args[index] = value;
                }
                else
                {
                    Args.AddRange(Enumerable.Repeat("", index - Args.Count).ToList());
                    Args.Add(value);
                }
            }
        }

        public List<string> GetHardipMeasurePin()
        {
            var arr = ArgList.Split(',');
            for (var i = 0; i < arr.Length; i++)
                if (arr[i].Equals("MeasI_pinS", StringComparison.CurrentCultureIgnoreCase))
                    if (i < Args.Count)
                        return Regex.Split(Args.GetRange(i, 1).First(), @"[,|+|;|:]").Select(x => x.Split('=').Last())
                            .ToList();
            return null;
        }

        public string TestName { get; set; }
        public string Type { get; set; }
        public string Name { get; set; }
        public string CalledAs { get; set; }
        public string DcCategory { get; set; }
        public string DcSelector { get; set; }
        public string AcCategory { get; set; }
        public string AcSelector { get; set; }
        public string TimeSets { get; set; }
        public string EdgeSets { get; set; }
        public string PinLevels { get; set; }
        public string MixedSignalTiming { get; set; }
        public string Overlay { get; set; }
        public string ArgList { get; set; }
        public List<string> Args { get; set; }
        public string Comment { get; set; }
        public List<string> InitList { get; set; }
        public List<string> PayloadList { get; set; }
    }
}