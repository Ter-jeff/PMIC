using System.IO;
using System.Text;
using unvell.ReoGrid.IO;

namespace SpreedSheet.Interface
{
    internal interface IPersistenceWorkbook
    {
        void Save(string path, FileFormat format = FileFormat._Auto, Encoding encoding = null);
        void Save(Stream stream, FileFormat format = FileFormat._Auto, Encoding encoding = null);
        void Load(string path, FileFormat format = FileFormat._Auto, Encoding encoding = null);
        void Load(Stream stream, FileFormat format = FileFormat._Auto, Encoding encoding = null, string sheetName = "");
    }
}