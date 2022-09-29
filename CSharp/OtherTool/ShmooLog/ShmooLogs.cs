using ShmooLog.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ShmooLog
{
    public class ShmooLogs : List<ShmooLog>
    {
        public string ConvertExcel(string output)
        {
            var summaryXlsx = "";
            var shmSets = new ShmooSets();
            foreach (var shmooLog in this)
            {
                var dt = new DataTable("All"); //把那個檔案內每個Device的Shmoo集中~
                foreach (var currDevice in shmooLog.CurrentDeviceNoList)
                {
                    var currShmooDt = shmooLog.ConcurrentDataTable[currDevice].Copy();
                    currShmooDt.Columns.Add(new DataColumn("File Name", typeof(string))
                    { DefaultValue = shmooLog.FileName }); //Report用的到
                    currShmooDt.Columns.Add(new DataColumn("Sort", typeof(string))
                    { DefaultValue = shmooLog.ConcurrentDicDevice[currDevice].Sort }); //1D Shmoo未來要Print 但原始資料沒有
                    currShmooDt.AcceptChanges();

                    if (currShmooDt.Rows.Count == 0) continue;

                    dt.TableName = this.First().FilePath; //以第一個檔案檔名為準去建立Xlsx
                    dt.Merge(currShmooDt);
                }


                shmSets.DsShmooSets.Tables.Add(dt);
                if (shmooLog.SramDef.HasSramDef)
                {
                    shmooLog.SramDef.BulitCompressStr();
                    shmSets.SramDataSet.AddRange(shmooLog.SramDef.SramDataSet);
                }
            }

            //Cook Shmoo Report for Epplus
            for (var tableNo = 0; tableNo < shmSets.DsShmooSets.Tables.Count; tableNo++) //代表一個個File
            {
                if (shmSets.DsShmooSets.Tables[tableNo].Rows.Count == 0) continue;

                //解析Shmoo Setup
                shmSets.ParseShmooSetupIdPerTable(tableNo);
                if (shmSets.ListAllCategories.Count == 0) //表示都被濾掉了 也不用Gen了
                    continue;

                try
                {
                    #region Cook 要放到Shmoo Report的sheets

                    shmSets.CookCurrentShmooReport();
                }
                catch (Exception)
                {
                    return "";
                }

                #endregion

                summaryXlsx = Path.Combine(output,
                    Path.GetFileName(Regex.Replace(shmSets.DsShmooSets.Tables[tableNo].TableName, @"\.txt",
                        "_Shmoo.xlsx")));
                HandleExcel.GenerateShmooReport(summaryXlsx, shmSets);
                //if (File.Exists(summaryXlsx))
                //    System.Diagnostics.Process.Start(summaryXlsx);
            }

            return summaryXlsx;
        }
    }
}