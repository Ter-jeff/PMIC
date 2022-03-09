using System.Text.RegularExpressions;

namespace IgxlData.Others.PatternListCsvFile
{
    public class Pattern
    {
        //Only Support N20 Product's Naming Rule!!

        public string OriginalUpperName = "N/A";
        public string FullNameWoMod = "N/A"; //��l�W�٥h���ɦW �M������j�g �h��MOD
        public string GenericName = "N/A"; //�h��PatVersion, SiliconVersion, TimeStamp

        public string Header = "N/A"; //PP DD CZ FA
        public string TimeStamp = "N/A"; //�����int, �]�����ǥ���|�ο���0xxx�}�Y

        //��Org + TypeSpec �P�_TpCategory
        public string Organization = "N/A"; //Organization : 'A:HARD_IP,C:CPU,L:GFX,P:HARD_IP,S:SOC,V:HARD_IP'

        public string TypeSpec = "N/A";
        //TypeSpec : 'AN:HARD_IP,BI:BIST,CH:SCAN,EF:HARD_IP,FU:HARD_IP,IO:HARD_IP,JT:HARD_IP,SC:SCAN,PW:HARD_IP'

        public string ProjectCode = "N/A";

        public string SiliconVersion = "N/A"; //A0 B0
        public int PatternVersion;

        public string TpCategory = "N/A"; //�u��Category���n 

        public bool IsMod;

        public string OpCode = "N/A";


        //�����T��Method�g�b���h �ե��w�g�ˬd���ŦXNaming Rule
        public Pattern(string patName, string regDateCode = @"\d{10}")
        {
            if (patName == "N/A" || patName == string.Empty) return;
            //patName= patName.Replace('/', '\\');
            //var elements = patName.Split('\\').Last().ToUpper().Split('.').First();
            //�h���ɦW �M������j�g
            patName = Regex.Replace(patName, @"\..+\\", "", RegexOptions.IgnoreCase).ToUpper().Trim(); //�h����
            patName = Regex.Replace(patName, @"\..+|\.", "", RegexOptions.IgnoreCase).ToUpper().Trim(); //�h����
            patName = Regex.Replace(patName, @".+/|.+\\", "", RegexOptions.IgnoreCase).ToUpper().Trim(); //�hFull Path 


            OriginalUpperName = patName;

            if (Regex.IsMatch(patName, @"_MOD.?")) //������bFull Name�e �]��Alignment�񪺬OFull Name, �[�W.?�]�����i��_MOD_XXX
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
            var orgAry = "A:HARD_IP,C:CPU,L:GFX,P:HARD_IP,S:SOC,V:HARD_IP,H:SOC".Split(','); //H:SOC�O�v�y���p
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
                //�p�G�O�зǥ��WPattern �̫�T�X�|�O PatVersion_SiliconRevision_TimeStamp

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