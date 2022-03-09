using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OTPFileComparison
{
    public class OtpFileReader
    {
        public static string HeaderKey = "Deterministic";
        public OTPFile Read(string fileName)
        {
            if (!File.Exists(fileName))
                return null;
            OTPFile otpFile = new OTPFile(fileName);
            List<string> fileLines = File.ReadAllLines(fileName).ToList();
            int headerRowIndex = fileLines.FindIndex(s => s.Split(',')[0].Equals(HeaderKey, StringComparison.OrdinalIgnoreCase));
            if(headerRowIndex <0)
            {
                throw new Exception("Can not find header row in OTP file: " + fileName);
            }

            //Read header
            int columnIndex = 1;
            foreach(string header in fileLines[headerRowIndex].Split(','))
            {
                if (otpFile.Headers.ContainsKey(header.Trim()))
                    throw new Exception(string.Format("Header '{0}' duplicate in OTP file: {1}", header, fileName));
                otpFile.Headers.Add(header.Trim(), columnIndex);
                columnIndex++;
            }

            //Read data
            for(int row = headerRowIndex + 1; row<fileLines.Count; row++)
            {
                otpFile.OTPRows.Add(fileLines[row].Split(',').ToList());
            }

            return otpFile;
        }

    }
}
