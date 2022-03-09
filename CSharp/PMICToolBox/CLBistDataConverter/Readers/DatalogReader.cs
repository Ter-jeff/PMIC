using System;
using System.Collections.Generic;
using System.IO;
using Library.DataStruct;
using CLBistDataConverter.DataStructures;
using CLBistDataConverter;

namespace CLBistDataConverter
{
    public class DatalogReader
    {        
        private FileStream _fs;
        private StreamReader _sr;

        public DatalogReader()
        {

        }

        public List<CLBistDie> Read(string filePath)
        {
            List<CLBistDie> clBistDieDatalist = new List<CLBistDie>();

            if (!File.Exists(filePath))
                return clBistDieDatalist;

            OpenFile(filePath);
            CLBistDie currentDie = null;
            CLBistSite currentSite = null;
            string lotId, waferId, dieLocationX, dieLocationY, dftGroup, instanceName = "", site="";
            bool readWaferDataStart = false;
            bool readCLBistDataStart = false;        
            try
            {
                int row = 1;
                string line;
                while ((line = _sr.ReadLine()) != null)
                {
                    line = line.Trim();
                    DataLogRowType rowContextType = CheckLogRowType(line, readWaferDataStart, readCLBistDataStart);
                    switch (rowContextType)
                    {
                        case DataLogRowType.IgnoredRow:
                            break;

                        case DataLogRowType.InstanceNameRow:
                            instanceName = RegStore.RegInstanceNameRow.Match(line).Groups["instancename"].ToString();
                            if (instanceName.Equals("OTP_UpdateECID", StringComparison.OrdinalIgnoreCase))
                            {
                                readWaferDataStart = true;
                                currentDie = new CLBistDie();
                                clBistDieDatalist.Add(currentDie);
                            }
                            else
                            {
                                if(instanceName.Equals("CLBIST_Rdc_FW",StringComparison.OrdinalIgnoreCase) || instanceName.Equals("CLBIST_FW",StringComparison.OrdinalIgnoreCase))
                                {
                                    readCLBistDataStart = true;
                                }else
                                {
                                    if (readCLBistDataStart)
                                        readCLBistDataStart = false;
                                }
                                if (readWaferDataStart)
                                    readWaferDataStart = false;
                            }
                            break;
                        case DataLogRowType.CLBistMeasureLog:
                            CLBistDataLogRow logRow = new CLBistDataLogRow(row, line);
                            if(currentDie != null)
                            {
                                currentSite = currentDie.CLBistSitelst.Find(s => s.Site.Equals(logRow.Site));
                                if (currentSite != null)
                                    currentSite.DatalogRows.Add(logRow);
                            }
                            break;
                        case DataLogRowType.Clk_Cfg5:
                            string dacNumber = RegStore.RegClk_Cfg5.Match(line).Groups["bdacnumber"].ToString();
                            string clk_cfg5 = RegStore.RegClk_Cfg5.Match(line).Groups["clk_cfg5"].ToString();
                            currentSite = currentDie.CLBistSitelst.Find(s => s.Site.Equals(site));
                            if (currentSite != null)
                            {
                                currentDie.CLBistSitelst.ForEach(s => s.clkCfg5Dic[dacNumber] = clk_cfg5);
                            }
                            break;
                        case DataLogRowType.ReadLotId:
                            site = RegStore.RegReadLotId.Match(line).Groups["site"].ToString();
                            string lotid = RegStore.RegReadLotId.Match(line).Groups["lotid"].ToString();
                            currentSite = currentDie.CLBistSitelst.Find(s => s.Site.Equals(site));
                            if (currentSite == null)
                            {
                                currentSite = new CLBistSite(site);
                                currentDie.CLBistSitelst.Add(currentSite);
                            }
                            currentSite.LotId = lotid;
                            break;
                        case DataLogRowType.ReadWaferId:
                            site = RegStore.RegReadWaferId.Match(line).Groups["site"].ToString();
                            string waferid = RegStore.RegReadWaferId.Match(line).Groups["waferid"].ToString();
                            currentSite = currentDie.CLBistSitelst.Find(s => s.Site.Equals(site));
                            if (currentSite == null)
                            {
                                currentSite = new CLBistSite(site);
                                currentDie.CLBistSitelst.Add(currentSite);
                            }
                            currentSite.WaferId = waferid;
                            break;
                        case DataLogRowType.ReadXCoord:
                            site = RegStore.RegReadXCoord.Match(line).Groups["site"].ToString();
                            string xcoord = RegStore.RegReadXCoord.Match(line).Groups["xcoord"].ToString();
                            currentSite = currentDie.CLBistSitelst.Find(s => s.Site.Equals(site));
                            if (currentSite == null)
                            {
                                currentSite = new CLBistSite(site);
                                currentDie.CLBistSitelst.Add(currentSite);
                            }
                            currentSite.DieLocationX = xcoord;
                            break;
                        case DataLogRowType.ReadYCoord:
                            site = RegStore.RegReadYCoord.Match(line).Groups["site"].ToString();
                            string ycoord = RegStore.RegReadYCoord.Match(line).Groups["ycoord"].ToString();
                            currentSite = currentDie.CLBistSitelst.Find(s => s.Site.Equals(site));
                            if (currentSite == null)
                            {
                                currentSite = new CLBistSite(site);
                                currentDie.CLBistSitelst.Add(currentSite);
                            }
                            currentSite.DieLocationY = ycoord;
                            break;
                        case DataLogRowType.ReadActStrM:
                            site = RegStore.RegReadActStrM.Match(line).Groups["site"].ToString();
                            string actstrm = RegStore.RegReadActStrM.Match(line).Groups["actstrm"].ToString();
                            currentSite = currentDie.CLBistSitelst.Find(s => s.Site.Equals(site));
                            if (currentSite == null)
                            {
                                currentSite = new CLBistSite(site);
                                currentDie.CLBistSitelst.Add(currentSite);
                            }
                            currentSite.DftGroup = actstrm;
                            break;
                    }

                    row++;
                }
            }
            catch (Exception e)
            {
                throw new Exception("Error when reading datalog file. " + e.Message.ToString());
            }
            finally
            {
                CloseFile();
            }

            return clBistDieDatalist;
        }

        private DataLogRowType CheckLogRowType(string lineContext, bool readWaferDataFlag, bool readCLBistDataStart)
        {
            if (readWaferDataFlag)
            {
                if (RegStore.RegReadLotId.IsMatch(lineContext))
                    return DataLogRowType.ReadLotId;
                if (RegStore.RegReadWaferId.IsMatch(lineContext))
                    return DataLogRowType.ReadWaferId;
                if (RegStore.RegReadXCoord.IsMatch(lineContext))
                    return DataLogRowType.ReadXCoord;
                if (RegStore.RegReadYCoord.IsMatch(lineContext))
                    return DataLogRowType.ReadYCoord;
                if (RegStore.RegReadActStrM.IsMatch(lineContext))
                    return DataLogRowType.ReadActStrM;
            }

            if (readCLBistDataStart)
            {
                if (RegStore.RegClk_Cfg5.IsMatch(lineContext))
                    return DataLogRowType.Clk_Cfg5;
                if (RegStore.RegClBistDatalogRow.IsMatch(lineContext))
                    return DataLogRowType.CLBistMeasureLog;
            }

            if (RegStore.RegInstanceNameRow.IsMatch(lineContext))
                return DataLogRowType.InstanceNameRow;

            return DataLogRowType.IgnoredRow;

        }

        private void OpenFile(string filePath)
        {
            _fs = new FileStream(filePath, FileMode.Open);
            _sr = new StreamReader(_fs);
        }

        private void CloseFile()
        {
            if (_sr != null)
            {
                _sr.Close();
                _sr.Dispose();
            }
            if (_fs != null)
            {
                _fs.Close();
                _fs.Dispose();
            }
        }
    }
}

public enum DataLogRowType
{
    ReadLotId,
    ReadWaferId,
    ReadXCoord,
    ReadYCoord,
    ReadActStrM,
    Clk_Cfg5,
    CLBistMeasureLog,
    InstanceNameRow,
    IgnoredRow
}
