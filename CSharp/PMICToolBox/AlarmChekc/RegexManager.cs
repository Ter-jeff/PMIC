using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace AlarmChekc
{
    public class RegexManager
    {
        public static Regex RegexFirst = new Regex(@"^\s*TheHdw\.(?<Instrument>Digital)\.Alarm\((?<AlarmCategory>.*?)\)\s*=\s*(?<AlarmBehavior>.*)");//    TheHdw.Digital.Alarm(tlHSDMAlarmAll) =tlAlarmForceBin
        public static Regex RegexSecond = new Regex(@"^\s*TheHdw\.(?<Instrument>Digital)(\.Pins)?\((?<Pins>.*?)\)\.Alarm\((?<AlarmCategory>.*?)\)\s*=\s*(?<AlarmBehavior>.*)");//    TheHdw.Digital.Pins("xx").Alarms(tlHSDMAlarmAll) = tlAlarmForceBin
        public static Regex RegexThird = new Regex(@"^\s*TheHdw\.(?<Instrument>DCVI)(\.Pins)?\((?<Pins>.*?)\)\.Alarm\((?<AlarmCategory>.*?)\)\s*=\s*(?<AlarmBehavior>.*)");// TheHdw.DCVI.Pins("xx").Alarm(tlDCVIAlarmAll) = tlAlarmForceFail
        public static Regex RegexFourth = new Regex(@"^\s*With\s+TheHdw\.(?<Instrument>Digital)(\.Pins)?\((?<Pins>.*)\)\s*$");// With TheHdw.DCVI("BUCK0_FB_UVI80,BUCK1_FB_UVI80,ATB0_UVI80,SPS_TSENSE_UVI80")
        public static Regex RegexFifth = new Regex(@"^\s*With\s+TheHdw\.(?<Instrument>DCVI)(\.Pins)?\((?<Pins>.*)\)\s*$");// With TheHdw.Digital("BUCK0_FB_UVI80,BUCK1_FB_UVI80,ATB0_UVI80,SPS_TSENSE_UVI80")

        public static Regex SubStartRegex = new Regex(@"\s*Sub\s+(?<FunctionName>[^\s]+)\(.*?\)");
        public static Regex SubEndRegex = new Regex(@"\s*End\s+Sub\s*$");
        public static Regex CommentRegex = new Regex(@"^\s*'.*$");

        public static Regex FunctionStartRegex = new Regex(@"\s*Function\s+(?<FunctionName>[^\s]+)\(.*?\)");
        public static Regex FunctionEndRegex = new Regex(@"^\s*End\s+Function\s*$");

        public static Regex RegexAlarmCatagory = new Regex(@"^\s*\.Alarm(\((?<AlarmCategory>.*?)\))?\s*=\s*(?<AlarmBehavior>.*)");
    }
}
