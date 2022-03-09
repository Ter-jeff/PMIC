using PmicAutomation.Utility.TCMID.DataStructure;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PmicAutomation.Utility.TCMID.Business
{
    class TcmIDGenConti : TcmIDGenBase
    {
        public TcmIDGenConti()
        {
        }

        public override void Gen(bool bCompare = false, bool bGenFlag = true)
        {
            List<DataRow> targetList = GetTestnameSortList();
            _tcmIdList = GenTCMID(targetList);
            if (bGenFlag)
            {
                GenCompareReport(_tcmIdList, bCompare);
                GenLimitSheet(_tcmIdList, bCompare);
            }
        }

        private List<DataRow> GetTestnameSortList()
        {
            IEnumerable<DataRow> collection = _limitDT.Rows.Cast<DataRow>();
            List<DataRow> targetList = collection.ToList().OrderBy(p => p[_idxTestname].ToString(), StringComparer.OrdinalIgnoreCase)
                .Where(s => !(Regex.IsMatch(s[_idxLowlim].ToString(), @"N/A", RegexOptions.IgnoreCase) || Regex.IsMatch(s[_idxHilim].ToString(), @"N/A", RegexOptions.IgnoreCase)))
                .Where(s => Regex.IsMatch(s[_idxTestname].ToString(), @"^open-*\w*_|^short-*\w*_", RegexOptions.IgnoreCase)).ToList();
            return targetList;
        }

        protected override string FetchTcmID(string testname)
        {
            return FetchTcmID(testname, "P");
        }

        protected override string GetTestType(string testname)
        {
            return "P";
        }
    }
}
