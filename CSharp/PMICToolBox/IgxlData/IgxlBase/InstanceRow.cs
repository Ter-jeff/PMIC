using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text.RegularExpressions;

namespace IgxlData.IgxlBase
{
    [Serializable]
    public class InstanceRow : IgxlRow
    {
        #region Field
        #endregion

        #region Property
        public string SheetName { get; set; }
        public int RowNum { get; set; }
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
        #endregion

        #region Constructor

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
        #endregion

        public bool IsEmptyRow
        {
            get
            {
                return string.IsNullOrEmpty(Name) &&
                       string.IsNullOrEmpty(TestName) &&
                       string.IsNullOrEmpty(ArgList) &&
                       string.IsNullOrEmpty(DcCategory) &&
                       string.IsNullOrEmpty(DcSelector) &&
                       string.IsNullOrEmpty(AcCategory) &&
                       string.IsNullOrEmpty(AcSelector) &&
                       string.IsNullOrEmpty(TimeSets) &&
                        string.IsNullOrEmpty(PinLevels);
            }
        }




        public string GetPropertyValue(string propertyName)
        {
            Type myType = typeof(InstanceRow);
            PropertyInfo myPropInfo = myType.GetProperty(propertyName);
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
            InstanceRow instanceRow = new InstanceRow();
            if (arr.Length < 15) return instanceRow;
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
            for (int i = 15; i < arr.Length; i++)
                instanceRow.Args.Add(arr[i]);
            if (arr.Length > 145)
                instanceRow.Comment = arr[145];
            return instanceRow;
        }

        public void SetArgument(string argument, string value)
        {
            int index = ArgList.Split(',').ToList().IndexOf(argument);
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
            for (int i = 0; i < arr.Count(); i++)
            {
                if (arr[i].Equals("MeasI_pinS", StringComparison.CurrentCultureIgnoreCase))
                {
                    if (i < Args.Count())
                    {
                        return Regex.Split(Args.GetRange(i, 1).First(), @"[,|+|;|:]").Select(x => x.Split('=').Last()).ToList();
                    }
                }
            }
            return null;
        }

        public bool InstanceRowCompare(InstanceRow row1, InstanceRow row2)
        {
            if (!(string.IsNullOrEmpty(row1.AcCategory) && string.IsNullOrEmpty(row2.AcCategory)))
                if (!row1.AcCategory.Equals(row2.AcCategory))
                    return false;
            if (!(string.IsNullOrEmpty(row1.AcSelector) && string.IsNullOrEmpty(row2.AcSelector)))
                if (!row1.AcSelector.Equals(row2.AcSelector))
                    return false;
            if (!(string.IsNullOrEmpty(row1.ArgList) && string.IsNullOrEmpty(row2.ArgList)))
                if (!row1.ArgList.Equals(row2.ArgList))
                    return false;
            if (!(string.IsNullOrEmpty(row1.CalledAs) && string.IsNullOrEmpty(row2.CalledAs)))
                if (!row1.CalledAs.Equals(row2.CalledAs))
                    return false;
            if (!(string.IsNullOrEmpty(row1.ColumnA) && string.IsNullOrEmpty(row2.ColumnA)))
                if (!row1.ColumnA.Equals(row2.ColumnA))
                    return false;
            if (!(string.IsNullOrEmpty(row1.Comment) && string.IsNullOrEmpty(row2.Comment)))
                if (!row1.Comment.Equals(row2.Comment))
                    return false;
            if (!(string.IsNullOrEmpty(row1.DcCategory) && string.IsNullOrEmpty(row2.DcCategory)))
                if (!row1.DcCategory.Equals(row2.DcCategory))
                    return false;
            if (!(string.IsNullOrEmpty(row1.DcSelector) && string.IsNullOrEmpty(row2.DcSelector)))
                if (!row1.DcSelector.Equals(row2.DcSelector))
                    return false;
            if (!(string.IsNullOrEmpty(row1.EdgeSets) && string.IsNullOrEmpty(row2.EdgeSets)))
                if (!row1.EdgeSets.Equals(row2.EdgeSets))
                    return false;
            if (!(string.IsNullOrEmpty(row1.MixedSignalTiming) && string.IsNullOrEmpty(row2.MixedSignalTiming)))
                if (!row1.MixedSignalTiming.Equals(row2.MixedSignalTiming))
                    return false;
            if (!(string.IsNullOrEmpty(row1.Name) && string.IsNullOrEmpty(row2.Name)))
                if (!row1.Name.Equals(row2.Name))
                    return false;
            if (!(string.IsNullOrEmpty(row1.Overlay) && string.IsNullOrEmpty(row2.Overlay)))
                if (!row1.Overlay.Equals(row2.Overlay))
                    return false;
            if (!(string.IsNullOrEmpty(row1.PinLevels) && string.IsNullOrEmpty(row2.PinLevels)))
                if (!row1.PinLevels.Equals(row2.PinLevels))
                    return false;
            if (!(string.IsNullOrEmpty(row1.SheetName) && string.IsNullOrEmpty(row2.SheetName)))
                if (!row1.SheetName.Equals(row2.SheetName))
                    return false;
            if (!(string.IsNullOrEmpty(row1.TestName) && string.IsNullOrEmpty(row2.TestName)))
                if (!row1.TestName.Equals(row2.TestName))
                    return false;
            if (!(string.IsNullOrEmpty(row1.TimeSets) && string.IsNullOrEmpty(row2.TimeSets)))
                if (!row1.TimeSets.Equals(row2.TimeSets))
                    return false;
            if (!(string.IsNullOrEmpty(row1.Type) && string.IsNullOrEmpty(row2.Type)))
                if (!row1.Type.Equals(row2.Type))
                    return false;
            if (row1.InitList.Count != 0 && row2.InitList.Count != 0)
            {
                if (row1.InitList.Count != row2.InitList.Count) return false;
                for (int i = 0; i < row1.InitList.Count; i++)
                {
                    if (!row1.InitList[i].Equals(row2.InitList[i]))
                        return false;
                }
            }

            if (row1.PayloadList.Count != 0 && row2.PayloadList.Count != 0)
            {
                if (row1.PayloadList.Count != row2.PayloadList.Count) return false;
                for (int i = 0; i < row1.PayloadList.Count; i++)
                {
                    if (!row1.PayloadList[i].Equals(row2.PayloadList[i]))
                        return false;
                }
            }

            var row1Args = row1.ArgList.Split(',').ToList();
            var row2Args = row2.ArgList.Split(',').ToList();
            if (row1Args.Count != row2Args.Count)
                return false;

            for (int i = 0; i < row1Args.Count; i++)
            {
                if (!(string.IsNullOrEmpty(row1Args[i]) && string.IsNullOrEmpty(row2Args[i])))
                    if (!row1Args[i].Equals(row2Args[i]))
                        return false;
            }
            return true;
        }
    }
}