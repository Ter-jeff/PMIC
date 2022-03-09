using CLBistDataConverter.DataStructures;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CLBistDataConverter
{
    public class FilerWriter
    {
        public void Write(List<CLBistDie> clBistDieDatalst, string outputFileName)
        {
            try
            {
                string rowPrefix, row;
                using (StreamWriter sw = new StreamWriter(outputFileName))
                {
                    sw.WriteLine("Lot ID, wafer ID, die location X, die location Y, DFT Group,Site, CL#, freq [MHz], clk_cfg5, bDAC1, bDAC2, bIREF, i1, i12, i2, i21, L11, L22, L12, L21, k1, k2, R11, R22, R12, R21, Rdc1, Rdc2, RrefA, RrefB, L11-L22, L12-L21, R11-R22, R12-R21, Ave L self, Ave L mutual, VDDH" + "\t");
                    foreach (CLBistDie dieData in clBistDieDatalst)
                    {
                        foreach (CLBistSite siteData in dieData.CLBistSitelst)
                        {
                            rowPrefix = siteData.LotId + "," + siteData.WaferId + "," + siteData.DieLocationX + ","
                                + siteData.DieLocationY + "," + siteData.DftGroup + "," + siteData.Site + ",";
                            foreach (CLBistSiteOutputRow outputRow in siteData.OutputRowlist)
                            {
                                row = rowPrefix + outputRow.ClNumber + "," + outputRow.Freq + "," + outputRow.Clk_cfg5 + "," +
                                    outputRow.BDac1 + "," + outputRow.BDac2 + "," + outputRow.BlRef + "," + outputRow.I1 + "," +
                                    outputRow.I12 + "," + outputRow.I2 + "," + outputRow.I21 + "," + outputRow.L11 + "," + outputRow.L22 + "," +
                                    outputRow.L12 + "," + outputRow.L21 + "," + outputRow.K1 + "," + outputRow.K2 + "," + outputRow.R11 + "," + outputRow.R22 + "," +
                                    outputRow.R12 + "," + outputRow.R21 + "," + outputRow.Rdc1 + "," + outputRow.Rdc2 + "," + outputRow.RrefA + "," +
                                    outputRow.RrefB + "," + outputRow.L11SubL22 + "," + outputRow.L12SubL21 + "," + outputRow.R11SubR22 + "," +
                                    outputRow.R12SubR21 + "," + outputRow.AveLself + "," + outputRow.AveLmutual + "," + outputRow.Vddh;
                                sw.WriteLine(row + "\t");
                            }

                        }
                    }
                }
            }catch(Exception ex)
            {
                throw new Exception("Error when writing csv file. " + ex.Message.ToString());
            }
        }
    }
}
