namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenConti.Base
{
    public class DcContiConst
    {
        public const string VbtFuncNameRelayControl = "Relay_Control";
        public const string VbtFuncNameFunctionalT = "Functional_T_updated";
        public const string FlagNamePowerShort = "F_powershort";
        public const string FlagNamePowerOpen = "F_poweropen";
        public const string FlagNameOpen = "F_open";
        public const string FlagNameShort = "F_short";
        public const string FlagNameVoltageClampCheck = "F_Conti_VoltageClamp_Check";
        public const string FlagNameAutoZCheck = "F_AutoZ_check";
        public const string BinNameOpenShort = "Bin_DC_open_short";
        public const string BinNameOpen = "Bin_DC_open";
        public const string BinNameShort = "Bin_DC_short";
        public const string BinNameContiVoltageClampCheck = "Bin_Conti_VoltageClamp_Check";
        public const string BinNamePowerShort = "Bin_DC_powershort";
        public const string BinNamePowerOpen = "Bin_DC_poweropen";
        public const string BinAutoZCheck = "Bin_AutoZ_check";
        public const string RelayWaitTime = "0.003";

        public const string DgsRelayOn = "DGS_Relay_On";
        public const string DgsRelayOff = "DGS_Relay_Off";

        //VBT
        public const string VbtIoContinuityParallel = "IO_Continuity_Parallel";
        public const string VbtIoContinuitySerial = "IO_Continuity_Serial";
        public const string VbtContiWalkingZ = "Conti_WalkingZ";
        public const string VbtPowerContinuityParallel = "Power_Continuity_Parallel";
        public const string VbtPowerContinuitySerial = "Power_Continuity_Serial";
        public const string VbtAnalogContinuityParallel = "Analog_Continuity_Parallel";
        public const string VbtAnalogContinuitySerial = "Analog_Continuity_Serial";

        //InstanceName
        public const string InsNameAutoZOnly = "DC_Continuity_Neg_AutoZOnly";
    }
}