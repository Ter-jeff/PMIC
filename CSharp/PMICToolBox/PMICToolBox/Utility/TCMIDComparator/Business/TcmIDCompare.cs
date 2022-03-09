using CommonLib.Controls;
using PmicAutomation.Utility.TCMID.Business;
using PmicAutomation.Utility.TCMID.DataStructure;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CompareStatus = PmicAutomation.Utility.TCMIDComparator.DataStructure.EnumStore.CompareStatus;

namespace PmicAutomation.Utility.TCMIDComparator.Business
{
    public class TcmIDCompare
    {
        private List<TcmIDGenBase> _tcmIdObjList;
        private MyForm.RichTextBoxAppend _AppendText;
        private List<Tuple<TcmIdEntry, TcmIdEntry>> _reportList; // Tuple<base, new>
        //private string _skipItemFile = Properties.Settings.Default.SkipItemConfig;
        //private List<string> _skipLines;

        public TcmIDCompare(TCMIDComparatorForm form, List<TcmIDGenBase> tcmIdObjList)
        {
            _AppendText = form.AppendText;
            _tcmIdObjList = tcmIdObjList;
            _reportList = new List<Tuple<TcmIdEntry, TcmIdEntry>>();
            //_skipLines = new List<string>();

            //if (form.skipItem.Checked)
                //SetupSkipItem();
        }

        //private void SetupSkipItem()
        //{
        //    if (File.Exists(_skipItemFile))
        //        _skipLines = File.ReadAllLines(_skipItemFile).ToList();
        //}

        /*
         * baseId = null => ADD 
         * newId = null => REMOVE
         */
        public void Process()
        {
            TcmIDGenBase objBase = _tcmIdObjList[0];
            TcmIDGenBase objNew = _tcmIdObjList[1];

            var listRemove = objBase.TcmIdList.Select(p => p.Testname).Except(objNew.TcmIdList.Select(s => s.Testname)).ToList();
            if (listRemove.Any())
            {
                foreach (string testname in listRemove)
                {
                    var target = objBase.TcmIdList.FirstOrDefault(p => p.Testname.Equals(testname));
                    target.Status = CompareStatus.REMOVE;
                    _reportList.Add(new Tuple<TcmIdEntry, TcmIdEntry>(target, null));
                }
            }

            // check skip item
            if (objNew.skipLines.Any())
            {
                foreach (TcmIdEntry newId in objNew.TcmIdList)
                {
                    if (objNew.skipLines.Exists(p => newId.Testname.IndexOf(p, StringComparison.OrdinalIgnoreCase) != -1))
                    {
                        newId.Status = CompareStatus.TCMID_REMOVE;
                        newId.TcmId = string.Empty;
                        _reportList.Add(new Tuple<TcmIdEntry, TcmIdEntry>(null, newId));
                        newId.Resettable = false;
                    }
                }
            }

            foreach (TcmIdEntry newId in objNew.TcmIdList)
            {
                if (newId.Resettable == false)
                    continue;

                TcmIdEntry baseId = objBase.TcmIdList.FirstOrDefault(p => p.Testname.ToUpper().Equals(newId.Testname.ToUpper()));
                if (baseId != null)
                {
                    if (!newId.TcmId.ToUpper().Equals(baseId.TcmId.ToUpper()))
                    {
                        newId.Status = CompareStatus.MODIFY;
                        newId.TcmId = baseId.TcmId;
                    }
                    else if (DoubleCompareNotEqual(newId, baseId))
                    {
                        newId.Status = CompareStatus.MODIFY;
                        _reportList.Add(new Tuple<TcmIdEntry, TcmIdEntry>(baseId, newId));
                    }
                    newId.Resettable = false;
                }
                else
                {
                    newId.Status = CompareStatus.ADD;
                    _reportList.Add(new Tuple<TcmIdEntry, TcmIdEntry>(baseId, newId));
                }
            }

            CalNewTcmID(objNew);

            objNew.GenCompareReport(objNew.TcmIdList, bCompare:true);
            objNew.GenLimitSheet(objNew.TcmIdList, bCompare: true);
            objNew.GenDiffReport(_reportList);
        }

        private bool DoubleCompareNotEqual(TcmIdEntry newId, TcmIdEntry baseId)
        {
            DataTable dt = new DataTable();
            var newLowLim = dt.Compute(newId.LowLim.Replace("=", ""), "");
            var newHiLim = dt.Compute(newId.HiLim.Replace("=", ""), "");
            var baseLowLim = dt.Compute(baseId.LowLim.Replace("=", ""), "");
            var baseHiLim = dt.Compute(baseId.HiLim.Replace("=", ""), "");

            if (Convert.ToDouble(newLowLim) != Convert.ToDouble(baseLowLim) || Convert.ToDouble(newHiLim) != Convert.ToDouble(baseHiLim))
                return true;
            else
                return false;
        }

        private void CalNewTcmID(TcmIDGenBase obj)
        {
            List<TcmIdEntry> listId = obj.TcmIdList;
            Dictionary<char, int> dicId = new Dictionary<char, int>();

            // ex: T0001, C0003
            List<string> items = listId.FindAll(s => s.Resettable == true && !string.IsNullOrEmpty(s.TcmId))
                .Select(p => p.TcmId.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries).Last()).ToList();
            foreach (string item in items)
            {
                char key = item.First();
                int value = Convert.ToInt32(item.Substring(1));
                if (!dicId.ContainsKey(key))
                    dicId.Add(key, value);
                else
                    dicId[key]++;
            }

            foreach (var id in listId)
            {
                if (id.Resettable && !string.IsNullOrEmpty(id.TcmId))
                {
                    string item = id.TcmId.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries).Last();
                    dicId[item.First()]++;
                    id.TcmId = id.TcmId.Replace(item, string.Format("{0}{1:D4}", item.First(), dicId[item.First()]));
                    id.OriginalTcmId = id.TcmId;
                    var target = _reportList.Find(p => p.Item2 != null && p.Item2.Equals(id));
                    _reportList.Remove(target);
                    _reportList.Add(new Tuple<TcmIdEntry, TcmIdEntry>(null, id));
                }
            }
        }
    }
}
