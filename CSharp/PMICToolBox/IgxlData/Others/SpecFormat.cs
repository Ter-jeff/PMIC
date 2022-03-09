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
        public const string SpecValuePreffixEqual = "=";
        public const string SpecValuePreffix = "_";
        public const string DcSpecValtSuffix = "VOP";

        //Add by Edward
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

        #region Member Function

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

        public static string GenGlbOther(string pinName, string other)
        {
            return CreateSpecSymbol(pinName, GlbSuffix + "_" + other);
        }

        public static string GenGlbRatio(string global, string ratio)
        {
            return "=_" + global + "*_" + ratio;
        }

        /// <summary>
        /// Generate spec symbol according to customer setting
        /// </summary>
        /// <param name="pinName"></param>
        /// <param name="suffix"></param>
        /// <returns></returns>
        public static string GenSpecSymbol(string pinName, string suffix)
        {
            return CreateSpecSymbol(pinName, suffix);
        }

        /// <summary>
        /// Change spec symbol
        /// </summary>
        /// <param name="src"></param>
        /// <param name="suffix"></param>
        /// <returns></returns>
        public static string CreateSpecSymbol(string src, string suffix)
        {
            return string.IsNullOrEmpty(suffix) ? src.Trim() : src.Trim() + "_" + suffix;
        }

        /// <summary>
        /// Judge whether aliase name exist
        /// If exist aliase name, then use aliase name
        /// else use the pin name
        /// </summary>
        /// <param name="pinName"></param>
        /// <param name="parameterName"></param>
        /// <returns></returns>
        /// <remarks>
        /// VDD_CPU
        /// VAlt 
        /// =_VDD_CPU_Valt_GLB
        /// </remarks>
        public static string GenGlbSpecSymbolAtLevelSheet(string pinName, string parameterName)
        {
            string lStrSrc = "";
            lStrSrc = "_" + pinName + "_" + parameterName;
            return CreateSpecSymbol(lStrSrc, GlbSuffix);
        }

        /// <summary>
        /// Judge whether aliase name exist
        /// If exist aliase name, then use aliase name
        /// else use the pin name
        /// </summary>
        /// <param name="pinName"></param>
        /// <param name="parameterName"></param>
        /// <returns></returns>
        /// <remarks>
        /// VDD_CPU
        /// Vmain 
        /// H
        /// =_VDD_CPU_Vmain_VAR_H
        /// </remarks>
        public static string GenDcSpecSymbolAtLevelSheet(string pinName, string parameterName, string pinType)
        {
            string lStrSrc = "";
            lStrSrc = "_" + pinName + "_" + parameterName;
            return CreateSpecSymbol(lStrSrc, string.IsNullOrEmpty(pinType) ? VarSuffix : VarSuffix + "_" + pinType);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pinName"></param>
        /// <param name="pinType"></param>
        /// <returns></returns>
        /// /// VDD_CPU
        /// =_VDD_CPU_VAR_H
        public static string GenDcSpecSymbolAtLevelSheet(string pinName, string pinType)
        {
            string lStrSrc = "";
            lStrSrc = "_" + pinName;
            return CreateSpecSymbol(lStrSrc, string.IsNullOrEmpty(pinType) ? VarSuffix : VarSuffix + "_" + pinType);
        }

        #endregion

        #region Gen Spec Value

        /// <summary>
        /// Get sepc new value
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string GenSpecValueSingleValue(string str)
        {
            return SpecValuePreffixEqual + str; // Add "=" before old value
        }

        public static string GenSpecValueSingleSpec(string value)
        {
            return SpecValuePreffixEqual + SpecValuePreffix + value;
        }

        #endregion

        #region

        /// <summary>
        /// If the Input is "Pins_1p1v_vil"
        /// return "Pins_1p1v"
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string GetRidOfIoPara(string str)
        {
            if (str.Split('_').Last().Equals(GlbIoPinVil) ||
                str.Split('_').Last().Equals(GlbIoPinVih) ||
                str.Split('_').Last().Equals(GlbIoPinVoh) ||
                str.Split('_').Last().Equals(GlbIoPinVol) ||
                str.Split('_').Last().Equals(GlbIoPinVt))
            {
                return str.Replace("_" + str.Split('_').Last(), "");
            }
            return str;
        }

        public static string GetRidOfCorePowerValt(string str)
        {
            if (str.Split('_').Last().Equals("VRS"))
            {
                return str.Replace("_" + "VRS", "");
            }
            return str;
        }

        #endregion

        #endregion
    }
}