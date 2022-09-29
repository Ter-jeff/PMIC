using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Inputs.ScghFile.Reader;
using System.Collections.Generic;

namespace PmicAutogen.Inputs.ScghFile.ProChar.Business
{
    public class MbistPatSetConstructor : ProdCharPatSetConstructorBase
    {
        public MbistPatSetConstructor(IEnumerable<IProdCharSheetRow> inputRows) : base(inputRows)
        {
            Block = "Mbist";
        }

        public List<ProdCharRowMbist> WorkFlow(bool removeNonUsage = false)
        {
            if (removeNonUsage)
            {
                PayloadList = FilterProChar(PayloadList);
                InitList = FilterProChar(InitList);
            }

            var prodCharRowMbists = new List<ProdCharRowMbist>();
            var prodCharRows = GetPatSetFromProdChar(InitList, PayloadList);

            foreach (var row in prodCharRows)
            {
                var prodCharRowMbist = row.NewProdCharRowMbist();
                prodCharRowMbist.Prefix = GetPrefix();
                prodCharRowMbist.InitPatSetNameByNamingRule = "";
                prodCharRowMbist.PatSetName = prodCharRowMbist.Prefix + "_" + GetPayLoadName(prodCharRowMbist);
                prodCharRowMbist.InstanceName = prodCharRowMbist.Prefix + "_" + GetPayLoadName(prodCharRowMbist);
                prodCharRowMbist.PerformanceMode = GetPerformanceMode(prodCharRowMbist, PerformanceModeList);
                prodCharRowMbist.RowNum = row.RowNum;
                CheckNop(prodCharRowMbist);
                prodCharRowMbists.Add(prodCharRowMbist);
            }

            return prodCharRowMbists;
        }

        private void CheckNop(ProdCharRowMbist prodCharRowMbist)
        {
            var prodCharRow = (ProdCharSheetRow)prodCharRowMbist.ProdCharItem;

            if (prodCharRowMbist.InitPatternMissing)
                prodCharRowMbist.Nop = true;

            if (prodCharRow.Usage != "1")
                prodCharRowMbist.Nop = true;
        }
    }
}