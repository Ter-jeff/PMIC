using System.Text.RegularExpressions;

namespace IgxlData.Others.PatternListCsvFile
{
    public class Pattern
    {
        //Only Support N20 Product's Naming Rule!!

        public string OriginalUpperName = "N/A";
        public string FullNameWoMod = "N/A"; //原始名稱去副檔名 然後轉全大寫 去掉MOD
        public string GenericName = "N/A"; //去掉PatVersion, SiliconVersion, TimeStamp

        public string Header = "N/A"; //PP DD CZ FA
        public string TimeStamp = "N/A"; //不能用int, 因為有些白爛會用錯的0xxx開頭

        //由Org + TypeSpec 判斷TpCategory
        public string Organization = "N/A"; //Organization : 'A:HARD_IP,C:CPU,L:GFX,P:HARD_IP,S:SOC,V:HARD_IP'

        public string TypeSpec = "N/A";
        //TypeSpec : 'AN:HARD_IP,BI:BIST,CH:SCAN,EF:HARD_IP,FU:HARD_IP,IO:HARD_IP,JT:HARD_IP,SC:SCAN,PW:HARD_IP'

        public string ProjectCode = "N/A";

        public string SiliconVersion = "N/A"; //A0 B0
        public int PatternVersion;

        public string TpCategory = "N/A"; //只用Category不好 

        public bool IsMod;

        public string OpCode = "N/A";


        //抽取資訊的Method寫在父層 勢必已經檢查完符合Naming Rule
        public Pattern(string patName, string regDateCode = @"\d{10}")
        {
            if (patName == "N/A" || patName == string.Empty) return;
            //patName= patName.Replace('/', '\\');
            //var elements = patName.Split('\\').Last().ToUpper().Split('.').First();
            //去副檔名 然後轉全大寫
            patName = Regex.Replace(patName, @"\..+\\", "", RegexOptions.IgnoreCase).ToUpper().Trim(); //去尾巴
            patName = Regex.Replace(patName, @"\..+|\.", "", RegexOptions.IgnoreCase).ToUpper().Trim(); //去尾巴
            patName = Regex.Replace(patName, @".+/|.+\\", "", RegexOptions.IgnoreCase).ToUpper().Trim(); //去Full Path 


            OriginalUpperName = patName;

            if (Regex.IsMatch(patName, @"_MOD.?")) //必須放在Full Name前 因為Alignment比的是Full Name, 加上.?因為有可能_MOD_XXX
            {
                patName = Regex.Replace(patName, @"_MOD.+|_MOD", "", RegexOptions.IgnoreCase);
                IsMod = true;
            }


            if (Regex.IsMatch(patName, "_DM_", RegexOptions.IgnoreCase))
            {
                OpCode = "DUAL";
            }
            if (Regex.IsMatch(patName, "_SI_", RegexOptions.IgnoreCase))
            {
                OpCode = "SINGLE";
            }


            FullNameWoMod = patName;
            Header = patName.Split('_')[0];
            Organization = patName.Split('_').Length > 2 ? patName.Split('_')[2] : "";
            TypeSpec = patName.Split('_').Length > 4 ? patName.Split('_')[4] : "N/A";
            ProjectCode = patName.Split('_').Length > 1 ? patName.Split('_')[1].ToUpper() : "";

            //Organization : 'A:HARD_IP,C:CPU,L:GFX,P:HARD_IP,S:SOC,V:HARD_IP'
            //TypeSpec : 'AN:HARD_IP,BI:BIST,CH:SCAN,EF:HARD_IP,FU:HARD_IP,IO:HARD_IP,JT:HARD_IP,SC:SCAN,PW:HARD_IP'
            var orgAry = "A:HARD_IP,C:CPU,L:GFX,P:HARD_IP,S:SOC,V:HARD_IP,H:SOC".Split(','); //H:SOC是權宜之計
            foreach (var s in orgAry)
            {
                var org = s.Split(':')[0];
                var cat = s.Split(':')[1];
                if (org == Organization) TpCategory = cat;
            }

            var tSpec =
                "AN:HARD_IP,BI:BIST,CH:SCAN,EF:HARD_IP,FU:HARD_IP,IO:HARD_IP,JT:HARD_IP,SC:SCAN,PW:HARD_IP,XX:OTHERS"
                    .Split(',');
            foreach (var s in tSpec)
            {
                var type = s.Split(':')[0];
                var cat = s.Split(':')[1];
                if (type == TypeSpec)
                {
                    if (cat == "HARD_IP")
                        TpCategory = "HARD_IP";
                    else if (cat == "OTHERS")
                        TpCategory = "OTHERS";
                    else if (TpCategory != "HARD_IP")
                        TpCategory = TpCategory + "_" + cat; //SOC_BIST
                }
            }

            var regPattern = string.Format(@"(?<name>.*_{0})", regDateCode);
            if (Regex.IsMatch(patName, regPattern))
            {
                //如果是標準全名Pattern 最後三碼會是 PatVersion_SiliconRevision_TimeStamp

                patName = Regex.Match(patName, regPattern).Groups["name"].ToString();
                var len = patName.Split('_').Length;

                int resA;
                if (int.TryParse(patName.Split('_')[len - 3], out resA))
                    PatternVersion = resA;

                SiliconVersion = patName.Split('_')[len - 2];
                TimeStamp = patName.Split('_')[len - 1];

                if (Regex.IsMatch(TimeStamp, regDateCode) || Regex.IsMatch(TimeStamp, @"YYMMDDHHMM"))
                {
                    GenericName = Regex.Replace(patName,
                        string.Format("_{0}_{1}_{2}", PatternVersion, SiliconVersion, TimeStamp), "");
                }
            }
            else
                GenericName = patName;
        }


    }
}