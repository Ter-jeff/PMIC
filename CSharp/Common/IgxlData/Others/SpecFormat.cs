using System.Linq;

namespace IgxlData.Others
{
    public class SpecFormat
    {
        #region Field

        public const string GlbSuffix = "GLB";
        public const string VarSuffix = "VAR";
        private const string SuffixPlus = "GLB_Plus";
        private const string SuffixMinus = "GLB_Minus";
        private const string SuffixPlusUHv = "GLB_Plus_UHV";
        private const string SuffixMinusULv = "GLB_Minus_ULV";
        public const string SpecValuePreffixEqual = "=";
        public const string SpecValuePreffix = "_";
        public const string DcSpecValtSuffix = "VOP";
        public const string SuffixHv = "GLB_HV";
        public const string SuffixLv = "GLB_LV";
        public const string SuffixUHv = "GLB_UHV";
        public const string SuffixULv = "GLB_ULV";

        public const string GlbIoPinVil = "Vil";
        public const string GlbIoPinVih = "Vih";
        public const string GlbIoPinVol = "Vol";
        public const string GlbIoPinVoh = "Voh";
        public const string GlbIoPinVt = "Vt";
        public const string GlbIoPinVcl = "Vcl";
        public const string GlbIoPinVch = "Vch";
        public const string GlbIoPinIol = "Iol";
        public const string GlbIoPinIoh = "Ioh";

        #endregion

        #region Gen Spec Symbol

        public static string GenGlbSpecSymbol(string pinName)
        {
            return CreateSpecSymbol(pinName, GlbSuffix);
        }

        public static string GenDcSpecSymbol(string pinName)
        {
            return CreateSpecSymbol(pinName, VarSuffix);
        }

        public static string GenAcSpecSymbol(string pinName)
        {
            return CreateSpecSymbol(pinName, VarSuffix);
        }

        public static string GenGlbMinus(string pinName)
        {
            return CreateSpecSymbol(pinName, SuffixMinus);
        }

        public static string GenGlbPlus(string pinName)
        {
            return CreateSpecSymbol(pinName, SuffixPlus);
        }

        public static string GenGlbMinusULv(string pinName)
        {
            return CreateSpecSymbol(pinName, SuffixMinusULv);
        }

        public static string GenGlbPlusUHv(string pinName)
        {
            return CreateSpecSymbol(pinName, SuffixPlusUHv);
        }

        public static string GenGlbHv(string pinName)
        {
            return CreateSpecSymbol(pinName, SuffixHv);
        }

        public static string GenGlbLv(string pinName)
        {
            return CreateSpecSymbol(pinName, SuffixLv);
        }

        public static string GenGlbULv(string pinName)
        {
            return CreateSpecSymbol(pinName, SuffixULv);
        }

        public static string GenGlbUHv(string pinName)
        {
            return CreateSpecSymbol(pinName, SuffixUHv);
        }

        public static string GenGlbOther(string pinName, string other)
        {
            return CreateSpecSymbol(pinName, GlbSuffix + "_" + other);
        }

        public static string GenGlbRatio(string global, string ratio)
        {
            return "=_" + global + "*_" + ratio;
        }

        public static string GenSpecSymbol(string pinName, string suffix)
        {
            return CreateSpecSymbol(pinName, suffix);
        }

        public static string CreateSpecSymbol(string src, string suffix)
        {
            return string.IsNullOrEmpty(suffix) ? src.Trim() : src.Trim() + "_" + suffix;
        }

        public static string GenGlbSpecSymbolAtLevelSheet(string pinName, string parameterName)
        {
            var src = "_" + pinName + "_" + parameterName;
            return CreateSpecSymbol(src, GlbSuffix);
        }

        public static string GenDcSpecSymbolAtLevelSheet(string pinName, string parameterName, string pinType)
        {
            var src = "_" + pinName + "_" + parameterName;
            return CreateSpecSymbol(src, string.IsNullOrEmpty(pinType) ? VarSuffix : VarSuffix + "_" + pinType);
        }

        public static string GenDcSpecSymbolAtLevelSheet(string pinName, string pinType)
        {
            var src = "_" + pinName;
            return CreateSpecSymbol(src, string.IsNullOrEmpty(pinType) ? VarSuffix : VarSuffix + "_" + pinType);
        }

        #endregion

        #region Gen Spec Value

        public static string GenSpecValueSingleValue(string str)
        {
            if (string.IsNullOrEmpty(str)) return "";
            return SpecValuePreffixEqual + str;
        }

        public static string GenSpecValueSingleSpec(string value)
        {
            return SpecValuePreffixEqual + SpecValuePreffix + value;
        }

        public static string GetRidOfIoPara(string str)
        {
            if (str.Split('_').Last().Equals(GlbIoPinVil) ||
                str.Split('_').Last().Equals(GlbIoPinVih) ||
                str.Split('_').Last().Equals(GlbIoPinVoh) ||
                str.Split('_').Last().Equals(GlbIoPinVol) ||
                str.Split('_').Last().Equals(GlbIoPinVt))
                return str.Replace("_" + str.Split('_').Last(), "");
            return str;
        }

        public static string GetRidOfCorePowerValt(string str)
        {
            if (str.Split('_').Last().Equals("VRS")) return str.Replace("_" + "VRS", "");
            return str;
        }

        #endregion
    }
}