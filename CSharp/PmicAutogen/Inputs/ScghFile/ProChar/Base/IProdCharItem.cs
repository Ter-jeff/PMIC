using System.Collections.Generic;

namespace PmicAutogen.Inputs.ScghFile.ProChar.Base
{
    public interface IProdCharSheetRow
    {
        int RowNum { set; get; }
        string PayloadValue { get; set; }
        string Block { set; get; }
        string Mode { set; get; }
        string Item { get; set; }
        string Usage { get; set; }
        string Inits { set; get; }
        string PayLoads { set; get; }
        List<string> GetInitList();
        List<string> GetPayloadList();
    }
}