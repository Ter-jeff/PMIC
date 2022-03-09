using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OTPFileComparison
{
    public class OTPFile
    {
        public Dictionary<string, int> Headers;
        public List<List<string>> OTPRows;
        private string _fileName;

        public string FileName
        {
            get { return _fileName; }
        }
        

        public OTPFile(string fileName)
        {
            _fileName = fileName;
            Headers = new Dictionary<string, int>();
            OTPRows = new List<List<string>>();
        }
    }
}
